import tkinter as tk
import pandas as pd
from tkinter import filedialog, messagebox, scrolledtext
import sys, json, os, time, threading, keyboard, logging, queue, re, glob, unicodedata
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchWindowException, TimeoutException, StaleElementReferenceException,
    WebDriverException, InvalidSessionIdException
)
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

BASE_DIR = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, 'frozen', False) else __file__))
CONFIG_FILE = os.path.join(BASE_DIR, 'config.json')



# --- 전역 변수 ---
CONFIG_FILE = 'config.json'
driver = None
automation_thread = None
is_running = False
is_verifying = False
_update_last_verify_cb = None
_safe_exit_ui = None

# --- 설정/로깅 ---
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                return {}
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4, ensure_ascii=False)

def clean_text(text):
    if text and text.lower().endswith('.pdf'):
        text = text[:-4]
    return re.sub(r'[^a-zA-Z0-9가-힣]', '', text or '')

def _nfc(s: str) -> str:
    return unicodedata.normalize('NFC', s or '')

def _squash_spaces(s: str) -> str:
    if not s:
        return ''
    s = s.replace('\u00A0', ' ').replace('\u200b', ' ').replace('\u200c', ' ')
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

class QueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue
    def emit(self, record):
        self.log_queue.put(self.format(record))

# --- 문서번호 추출 ---
def extract_docno(purpose_text: str) -> str:
    if not purpose_text:
        return ''
    m = re.search(r'\[([^\[\]]+)\]', purpose_text)
    return (m.group(1).strip() if m else '')

# --- GUI ---
class SettingsGUI:
    def __init__(self, master, log_queue):
        self.master = master
        self.log_queue = log_queue
        self.config = load_config()
        self.pdf_path = tk.StringVar(value=self.config.get("pdf_path", ""))
        self.last_verify_var = tk.StringVar(value='최근 검증: -')

        master.title("e나라도움 파일 업로드 자동화 프로그램 ver.1.0")
        master.geometry("650x560")
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

        settings_frame = tk.Frame(master); settings_frame.pack(pady=10, padx=10, fill="x")
        settings_frame.grid_columnconfigure(0, weight=1)
        tk.Label(settings_frame, text="PDF 파일 폴더 경로:").grid(row=0, column=0, columnspan=2, sticky="w")
        self.path_entry = tk.Entry(settings_frame, textvariable=self.pdf_path, width=70); self.path_entry.grid(row=1, column=0, sticky="ew")
        tk.Button(settings_frame, text="폴더 찾기", command=self.browse_folder).grid(row=1, column=1, padx=5)

        button_frame = tk.Frame(settings_frame); button_frame.grid(row=2, column=0, columnspan=2, pady=5, sticky="w")
        tk.Button(button_frame, text="경로 저장", command=self.save_settings, font=('Helvetica', 10, 'bold')).pack(side="left", padx=(0,5))
        tk.Button(button_frame, text="업로드 검증", command=self.start_verification, font=('Helvetica', 10, 'bold'), bg="#DAF7A6").pack(side="left")

        tk.Label(settings_frame, textvariable=self.last_verify_var, fg='#555').grid(row=3, column=0, columnspan=2, sticky='w', pady=(4,0))

        info_frame = tk.LabelFrame(master, text="단축키 안내"); info_frame.pack(pady=5, padx=10, fill="x")
        tk.Label(info_frame, text="  - [ F1 ] : 디버그 모드 크롬에 연결").pack(anchor="w")
        tk.Label(info_frame, text="  - [ F3 ] : 파일 업로드 자동화 시작 / 중지").pack(anchor="w")
        tk.Label(info_frame, text="  - [ ESC ] : 프로그램 전체 종료").pack(anchor="w")

        tk.Label(master, text="CONTACT: wooyxxng@gmail.com", font=("Helvetica", 8), fg="gray").pack(side="bottom", pady=(0,5))
        log_frame = tk.LabelFrame(master, text="실행 로그"); log_frame.pack(pady=10, padx=10, fill="both", expand=True)
        self.log_widget = scrolledtext.ScrolledText(log_frame, state='disabled', wrap=tk.WORD, font=("Consolas", 9))
        self.log_widget.pack(fill="both", expand=True)

        global _update_last_verify_cb, _safe_exit_ui
        _update_last_verify_cb = self.update_last_verification
        _safe_exit_ui = self.countdown_quit

        self.master.after(100, self.process_queue)

    def update_last_verification(self, when_str: str, filepath: str):
        try:
            base = os.path.basename(filepath) if filepath else '-'
            self.last_verify_var.set(f'최근 검증: {when_str} (파일: {base})')
        except Exception:
            pass

    def process_queue(self):
        try:
            while True:
                record = self.log_queue.get(block=False)
                self.log_widget.configure(state='normal')
                self.log_widget.insert(tk.END, record + '\n')
                self.log_widget.see(tk.END)
                self.log_widget.configure(state='disabled')
        except queue.Empty:
            pass
        self.master.after(100, self.process_queue)

    def browse_folder(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.pdf_path.set(folder_selected)

    def save_settings(self):
        path = self.pdf_path.get()
        if not path or not os.path.isdir(path):
            messagebox.showerror("오류", "유효한 PDF 폴더 경로를 지정해주세요.")
            return
        self.config["pdf_path"] = path
        save_config(self.config)
        messagebox.showinfo("저장 완료", "PDF 폴더 경로가 성공적으로 저장되었습니다.")

    def start_verification(self):
        global driver, is_running, is_verifying
        if driver is None:
            messagebox.showwarning("경고", "먼저 F1로 디버그 모드 크롬에 연결해주세요.")
            return
        if is_running:
            messagebox.showwarning("실행 중", "파일 업로드 작업이 실행 중입니다. 업로드가 끝난 뒤에 검증을 실행해주세요.")
            return
        if is_verifying:
            messagebox.showinfo("안내", "이미 업로드 검증이 실행 중입니다.")
            return

        def _verify_guard():
            global is_verifying
            is_verifying = True
            try:
                run_verification_logic()
            finally:
                is_verifying = False

        threading.Thread(target=_verify_guard, daemon=True).start()

    def on_closing(self):
        if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
            logging.info(">> GUI 창을 닫습니다. 프로그램을 종료합니다...")
            global is_running, driver
            is_running = False
            if driver:
                try: driver.quit()
                except: pass
            self.master.destroy()
            os._exit(0)

    def countdown_quit(self, msg: str, delay_sec: int = 3):
        win = tk.Toplevel(self.master)
        win.title("세션 오류")
        win.transient(self.master)
        win.grab_set()
        win.resizable(False, False)

        frame = tk.Frame(win, padx=18, pady=14)
        frame.pack(fill="both", expand=True)

        lbl_msg = tk.Label(frame, text=msg, justify="left", anchor="w")
        lbl_msg.pack(anchor="w")

        lbl_count = tk.Label(frame, text="", fg="#b00", font=("Helvetica", 11, "bold"))
        lbl_count.pack(anchor="e", pady=(8, 0))

        def tick(n):
            lbl_count.config(text=f"{n}초 후 프로그램이 자동 종료됩니다…")
            if n <= 0:
                try:
                    self.master.destroy()
                except Exception:
                    pass
                os._exit(0)
            else:
                self.master.after(1000, lambda: tick(n-1))

        self.master.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width()//2 - 180)
        y = self.master.winfo_y() + (self.master.winfo_height()//2 - 60)
        win.geometry(f"+{max(0,x)}+{max(0,y)}")

        tick(delay_sec)

# --- 브라우저 연결: F1 = 디버그 모드 연결 ---
def connect_to_existing_browser():
    global driver, _safe_exit_ui
    try:
        if driver and driver.window_handles:
            logging.info(">> 이미 브라우저가 연결되어 있습니다.")
            return
    except (NoSuchWindowException, InvalidSessionIdException, WebDriverException):
        logging.warning(">> 기존 WebDriver 세션이 무효입니다. 프로그램을 종료합니다.")
        driver = None
        if callable(_safe_exit_ui):
            _safe_exit_ui("무효한 WebDriver 세션이 감지되었습니다.\n(이전 창이 닫혔거나 연결이 끊겼습니다.)")
        else:
            print("세션 오류: 3초 후 종료"); time.sleep(3); os._exit(0)
        return

    logging.info("\n>> F1 감지: 기존 크롬 브라우저(디버그 모드)에 연결을 시도합니다...")
    try:
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logging.info(f">> 브라우저에 성공적으로 연결되었습니다. 현재 페이지: {driver.title}")
    except Exception as e:
        logging.error(">> 연결 실패. 디버그 모드로 크롬이 실행 중인지 확인하세요.")
        logging.error(f"   (상세 오류: {e})")
        driver = None

# --- 공용 헬퍼 함수 모음 ---
def get_tr_index(row):
    try:
        return (row.get_attribute('index') or '').strip()
    except Exception:
        return ''

def find_visible_rows():
    grid = driver.find_element(By.ID, "DB007001Q_GridArea")
    rows = grid.find_elements(By.CSS_SELECTOR, ".IBBodyMid .IBSection tr.IBDataRow")
    return rows if rows else grid.find_elements(By.CSS_SELECTOR, ".IBBodyRight .IBSection tr.IBDataRow")

def get_row_key(row):
    for attr in ("id","rowid","rid","r","data-rowid"):
        v = row.get_attribute(attr)
        if v: return f"{attr}:{v}"
    idx = row.get_attribute("index") or ""
    try:
        purpose = row.find_element(By.CSS_SELECTOR, 'td[class*="excutPrposCn"]').text.strip()
    except Exception:
        purpose = ""
    try:
        whole = (row.get_attribute("innerText") or "").strip()[:80]
    except Exception:
        whole = ""
    return f"idx:{idx}|p:{purpose}|w:{whole}"

def get_text_by_class(row, class_keyword):
    try:
        td = row.find_element(By.CSS_SELECTOR, f'td[class*="{class_keyword}"]')
        return td.text.strip(), td
    except Exception:
        return "", None

def get_index_cell_text(row):
    try:
        tds = row.find_elements(By.TAG_NAME, "td")
        if tds:
            return tds[0].text.strip()
    except Exception:
        pass
    try:
        return row.get_attribute("index") or ""
    except Exception:
        return ""

def _wait_any_appear(locators, timeout=8):
    end = time.time() + timeout
    found = None
    while time.time() < end and not found:
        found = _find_element_all_contexts(locators, timeout=0.6)
        if found:
            return found
    return None

def _click_with_retry(locators, tries=3, delay=0.6):
    for _ in range(tries):
        if _smart_click_any(locators, timeout=3):
            return True
        time.sleep(delay)
    return False

def _save_verification_excel(rows_data):
    ts = time.strftime('%Y%m%d_%H%M%S')
    xlsx_path = os.path.join(BASE_DIR, f'upload_verification_{ts}.xlsx')  # ← 여기만 변경

    if pd is not None:
        try:
            df = pd.DataFrame(rows_data)
            df.rename(columns={
                'idx': '인덱스',
                'docno': '문서번호',
                'purpose': '집행용도',
                'attach': '개별첨부'
            }, inplace=True)
            with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='verify')
            return xlsx_path, time.strftime('%Y-%m-%d %H:%M:%S')
        except Exception as e:
            logging.error(f'[엑셀 저장 오류] {e} — CSV로 대체 저장합니다.')
    csv_path = xlsx_path.replace('.xlsx', '.csv')
    try:
        import csv
        with open(csv_path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.DictWriter(f, fieldnames=['인덱스','문서번호','집행용도','개별첨부'])
            writer.writeheader()
            for row in rows_data:
                writer.writerow({
                    '인덱스': row['idx'],
                    '문서번호': row['docno'],
                    '집행용도': row['purpose'],
                    '개별첨부': row['attach']
                })
        return csv_path, time.strftime('%Y-%m-%d %H:%M:%S')
    except Exception as e:
        logging.error(f'[CSV 저장 실패] {e}')
        return '', time.strftime('%Y-%m-%d %H:%M:%S')

def _find_element_all_contexts(locators, timeout=8):
    end = time.time() + timeout
    def _search_here():
        for by, sel in locators:
            try:
                els = driver.find_elements(by, sel)
                for el in els:
                    if el.is_displayed():
                        return el
            except Exception:
                continue
        return None
    while time.time() < end:
        try: driver.switch_to.default_content()
        except Exception: pass
        el = _search_here()
        if el: return el
        try:
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for f in iframes:
                if not f.is_displayed(): continue
                try:
                    driver.switch_to.frame(f)
                    el = _search_here()
                    if el: return el
                finally:
                    try: driver.switch_to.default_content()
                    except Exception: pass
        except Exception: pass
        time.sleep(0.2)
    try: driver.switch_to.default_content()
    except Exception: pass
    return None

def _smart_click_any(locators, timeout=8):
    el = _find_element_all_contexts(locators, timeout=timeout)
    if not el: return False
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el); time.sleep(0.1)
    except Exception:
        pass
    for method in ("normal","actions","js"):
        try:
            if method=="normal": el.click()
            elif method=="actions": ActionChains(driver).move_to_element(el).pause(0.05).click().perform()
            else: driver.execute_script("arguments[0].click();", el)
            return True
        except Exception:
            continue
    return False

def _click_ok_in_any_frame(timeout=8):
    end = time.time() + timeout
    PRIMARY = [
        (By.CSS_SELECTOR, "footer.message button.fn.ok"),
        (By.XPATH, "//footer[contains(@class,'message')]//button[contains(@class,'fn') and contains(@class,'ok')]")
    ]
    FALLBACK = [
        (By.XPATH, "//button[normalize-space(text())='확인']"),
        (By.XPATH, "//*[@role='button' and normalize-space(text())='확인']"),
        (By.XPATH, "//input[@type='button' and (contains(@value,'확인') or contains(@title,'확인'))]"),
        (By.XPATH, "//*[normalize-space(text())='OK' or @value='OK' or contains(@aria-label,'OK')]")
    ]
    def _try_alert():
        try:
            driver.switch_to.alert.accept(); return True
        except Exception:
            return False
    def _click_here():
        for loc in PRIMARY + FALLBACK:
            try:
                for el in driver.find_elements(*loc):
                    if not el.is_displayed(): continue
                    try: driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
                    except Exception: pass
                    for m in ("normal","actions","js"):
                        try:
                            if m=="normal": el.click()
                            elif m=="actions": ActionChains(driver).move_to_element(el).pause(0.05).click().perform()
                            else: driver.execute_script("arguments[0].click();", el)
                            return True
                        except Exception:
                            continue
            except Exception:
                continue
        return False
    while time.time() < end:
        if _try_alert(): return True
        if _click_here(): return True
        try: driver.switch_to.default_content()
        except Exception: pass
        if _click_here(): return True
        try:
            for f in driver.find_elements(By.TAG_NAME, "iframe"):
                if not f.is_displayed(): continue
                try:
                    driver.switch_to.frame(f)
                    if _click_here(): return True
                finally:
                    try: driver.switch_to.default_content()
                    except Exception: pass
        except Exception: pass
        time.sleep(0.2)
    try: driver.switch_to.default_content()
    except Exception: pass
    return False

def _wait_modal_close_or_toast(timeout=6):
    time.sleep(0.6)
    end = time.time() + timeout
    while time.time() < end:
        time.sleep(0.2)
    return True

def _switch_to_new_window_if_any(open_timeout=4, baseline_handles=None):
    if baseline_handles is None:
        baseline_handles = set(driver.window_handles)
    end = time.time() + open_timeout
    while time.time() < end:
        cur = set(driver.window_handles)
        new = list(cur - baseline_handles)
        if new:
            try:
                driver.switch_to.window(new[0])
                return True
            except Exception:
                return False
        time.sleep(0.2)
    return False

def resolve_pdf_path(purpose_text: str, folder: str) -> str:
    if not folder or not os.path.isdir(folder) or not purpose_text:
        return ""
    desired_clean = clean_text(_squash_spaces(_nfc(purpose_text)))

    for path in glob.glob(os.path.join(folder, "*.pdf")):
        name = os.path.splitext(os.path.basename(path))[0]
        name_clean = clean_text(_squash_spaces(_nfc(name)))
        if name_clean == desired_clean:
            return path

    for path in glob.glob(os.path.join(folder, "*.pdf")):
        name = os.path.splitext(os.path.basename(path))[0]
        name_clean = clean_text(_squash_spaces(_nfc(name)))
        if desired_clean and name_clean and (desired_clean in name_clean or name_clean in desired_clean):
            return path

    normalized_target = _squash_spaces(_nfc(purpose_text)).lower()
    best = ""
    for path in glob.glob(os.path.join(folder, "*.pdf")):
        name = os.path.splitext(os.path.basename(path))[0]
        name_norm = _squash_spaces(_nfc(name)).lower()
        if normalized_target == name_norm:
            return path
        if len(normalized_target) >= 10 and normalized_target in name_norm:
            best = path
            break
    return best

def find_magnifier(row):
    try:
        attach_td = row.find_element(By.CSS_SELECTOR, 'td[class*="atchmnflNm"]')
        magnifier_td = attach_td.find_element(By.XPATH, "following-sibling::td[1]")
        for sel in ["img", "button", "a", "[role='button']"]:
            for el in magnifier_td.find_elements(By.CSS_SELECTOR, sel):
                if el.is_displayed():
                    return el
        return magnifier_td
    except Exception:
        pass
    try:
        tds = row.find_elements(By.TAG_NAME, "td")
        for col in reversed(tds[-3:] if len(tds) >= 3 else tds):
            for sel in ["img", "button", "a", "[role='button']"]:
                for el in col.find_elements(By.CSS_SELECTOR, sel):
                    if el.is_displayed():
                        return el
        if tds:
            return tds[-1]
    except Exception:
        pass
    return None

def _ensure_file_input_and_send(pdf_path, pre_timeout=3.0):
    end = time.time() + pre_timeout
    last_err = None

    def _collect_inputs_in_current_context():
        script = r"""
        const out = [];
        function collect(root) {
            if (!root) return;
            const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT, null);
            while (walker.nextNode()) {
                const el = walker.currentNode;
                if (el.tagName && el.tagName.toLowerCase() === 'input' && el.type === 'file') {
                    out.push(el);
                }
                if (el.shadowRoot) {
                    collect(el.shadowRoot);
                }
            }
        }
        collect(document);
        return out;
        """
        try:
            return driver.execute_script(script) or []
        except Exception:
            return []

    while time.time() < end:
        try:
            try: driver.switch_to.default_content()
            except: pass

            candidates = _collect_inputs_in_current_context()

            try:
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
            except Exception:
                iframes = []
            for f in iframes:
                if not f.is_displayed():
                    continue
                try:
                    driver.switch_to.frame(f)
                    candidates.extend(_collect_inputs_in_current_context())
                except Exception:
                    pass
                finally:
                    try: driver.switch_to.default_content()
                    except Exception:
                        pass

            for file_input in candidates:
                try:
                    try:
                        driver.execute_script(
                            "arguments[0].style.display='block';"
                            "arguments[0].style.visibility='visible';"
                            "arguments[0].removeAttribute('disabled');",
                            file_input
                        )
                        time.sleep(0.05)
                    except Exception:
                        pass
                    file_input.send_keys(pdf_path)
                    return True
                except Exception as e:
                    last_err = e
                    continue

            time.sleep(0.2)
        except Exception as e:
            last_err = e
            time.sleep(0.2)

    logging.error(f"  -> 파일 input send_keys 실패(대기 초과): {last_err}")
    return False

# --- 업로드 검증 ---
def run_verification_logic():
    global driver, _update_last_verify_cb
    try:
        logging.info("\n>> 업로드 검증을 시작합니다...")
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#DB007001Q_GridArea tr.IBDataRow"))
        )
        logging.info("-> 첫 데이터 행 발견. 검증을 시작합니다.")

        targets, seen = [], set()
        rows_for_excel = []

        def scan(rows):
            nonlocal targets, seen, rows_for_excel
            for row in rows:
                try:
                    k = get_row_key(row)
                    if k in seen:
                        continue
                    seen.add(k)

                    idx = get_tr_index(row)
                    try:
                        p_td = row.find_element(By.CSS_SELECTOR, 'td[class*="excutPrposCn"]')
                    except Exception:
                        p_td = None
                    try:
                        a_td = row.find_element(By.CSS_SELECTOR, 'td[class*="atchmnflNm"]')
                    except Exception:
                        a_td = None

                    if p_td and a_td:
                        p, a = p_td.text.strip(), a_td.text.strip()
                        if clean_text(p) != clean_text(a):
                            targets.append(f"[{idx}] '{p}'")
                            rows_for_excel.append({
                                "idx": idx,
                                "docno": extract_docno(p),
                                "purpose": p,
                                "attach": a,
                            })
                except StaleElementReferenceException:
                    continue

        scan(find_visible_rows())

        grid_main = driver.find_element(By.CSS_SELECTOR, "#DB007001Q_GridArea .SheetMain")
        ActionChains(driver).move_to_element(grid_main).click().perform()
        logging.info("-> 스크롤을 시작합니다...")

        while True:
            last = len(seen)
            rows_now = find_visible_rows()
            if not rows_now:
                break
            try:
                ActionChains(driver).move_to_element(rows_now[-1]).click().perform()
            except Exception:
                pass

            actions = ActionChains(driver)
            for _ in range(8):
                actions.send_keys(Keys.ARROW_DOWN)
            actions.perform()
            time.sleep(0.8)

            scan(find_visible_rows())

            if len(seen) == last:
                ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
                time.sleep(0.6)
                scan(find_visible_rows())
                if len(seen) == last:
                    break

        logging.info("-> 스크롤 완료.")

        if not targets:
            logging.info(">> [검증 결과] 모든 파일이 정상적으로 첨부되어 있습니다.")
        else:
            logging.info(f">> [검증 결과] 다음 항목들의 업로드/수정이 필요합니다. (총 {len(targets)}건)")
            for t in targets:
                logging.info(t)

        if rows_for_excel:
            saved_path, when_str = _save_verification_excel(rows_for_excel)
            if saved_path:
                logging.info(f">> 검증 결과 파일 저장 완료: {saved_path}")
                try:
                    if callable(_update_last_verify_cb):
                        _update_last_verify_cb(when_str, saved_path)
                except Exception:
                    pass
            else:
                logging.warning(">> 검증 결과 파일 저장에 실패했습니다.")
        else:
            when_str = time.strftime('%Y-%m-%d %H:%M:%S')
            try:
                if callable(_update_last_verify_cb):
                    _update_last_verify_cb(when_str, '')
            except Exception:
                pass

        logging.info(">> 검증이 완료되었습니다.")
    except Exception as e:
        logging.error(f"\n[오류] 검증 작업 중 예외가 발생했습니다: {e}")
        messagebox.showerror("검증 오류", f"검증 중 오류가 발생했습니다:\n{e}")

# --- 업로드 자동화 (한 칸씩 순차 처리) ---
def main_automation_logic(pdf_folder_path):
    global is_running, driver
    try:
        logging.info("\n>> 파일 업로드 자동화를 시작합니다...")
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "#DB007001Q_GridArea tr.IBDataRow")))
        processed = set()  # 안전망(중복 처리 방지)

        def to_int(s):
            return int(s) if s and str(s).isdigit() else None

        def focus_grid():
            try:
                grid_main = driver.find_element(By.CSS_SELECTOR, "#DB007001Q_GridArea .SheetMain")
                ActionChains(driver).move_to_element(grid_main).click().perform()
                time.sleep(0.05)
                return True
            except Exception:
                return False

        def visible_index_map():
            """현재 보이는 행들을 {idx:int -> row WebElement}로, 그리고 (min,max) 반환"""
            rows = find_visible_rows()
            idx_map = {}
            min_i, max_i = None, None
            for r in rows:
                idx_s = get_tr_index(r) or get_index_cell_text(r)
                i = to_int(idx_s)
                if i is None: continue
                idx_map[i] = r
                if min_i is None or i < min_i: min_i = i
                if max_i is None or i > max_i: max_i = i
            return idx_map, min_i, max_i

        def press_keys(key, n=1, pause=0.04):
            actions = ActionChains(driver)
            for _ in range(n):
                actions.send_keys(key)
            actions.perform()
            time.sleep(pause)

        def seek_to_index(target_idx, max_iters=80):
            """
            target_idx가 보일 때까지 위/아래로 조금씩 이동하며 찾기
            """
            if not focus_grid():
                return None
            last_range = None
            for _ in range(max_iters):
                idx_map, mn, mx = visible_index_map()
                if target_idx in idx_map:
                    return idx_map[target_idx]
                if mn is None or mx is None:
                    press_keys(Keys.ARROW_DOWN, 3); continue
                last_range = (mn, mx)
                if target_idx < mn:
                    # 너무 내려와버렸음 → 위로
                    press_keys(Keys.ARROW_UP, 5)
                elif target_idx > mx:
                    # 아직 위쪽 → 아래로
                    press_keys(Keys.ARROW_DOWN, 5)
                else:
                    # 범위 안인데 virtualization으로 DOM에 아직 없음 → 살짝 흔들기
                    press_keys(Keys.ARROW_DOWN, 1)
                    press_keys(Keys.ARROW_UP, 1)
            logging.warning(f"  -> [경고] 인덱스 {target_idx} 노출 실패(가시범위: {last_range})")
            return None

        def _micro_rescan_after_refresh(process_fn, up_steps=2, down_steps=4, pause=0.12):
            try:
                press_keys(Keys.ARROW_UP, up_steps, pause)
                process_fn()  # 위쪽 일부 확인
                press_keys(Keys.ARROW_DOWN, down_steps, pause)
                process_fn()  # 다시 내려오며 확인
            except Exception:
                pass

        def do_upload(purpose_text, attachment_text, trigger, pdf_folder_path):
            global driver
            pdf_path = resolve_pdf_path(purpose_text, pdf_folder_path)
            if not pdf_path or not os.path.exists(pdf_path):
                logging.warning(f"  -> 업로드할 파일을 찾지 못했습니다. 집행용도: '{purpose_text}', 폴더 경로: '{pdf_folder_path}'")
                return False
            logging.info(f"  -> 매칭된 파일: {pdf_path}")

            parent_handle = driver.current_window_handle
            pre_handles = set(driver.window_handles)

            try:
                try:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", trigger)
                    time.sleep(0.1)
                except Exception:
                    pass
                for method in ("normal", "actions", "js", "enter"):
                    try:
                        if method == "normal":
                            trigger.click()
                        elif method == "actions":
                            ActionChains(driver).move_to_element(trigger).pause(0.05).click().perform()
                        elif method == "js":
                            driver.execute_script("arguments[0].click();", trigger)
                        elif method == "enter" and getattr(trigger, 'tag_name', '').lower() == "td":
                            ActionChains(driver).move_to_element(trigger).click().perform()
                            time.sleep(0.05)
                            ActionChains(driver).send_keys(Keys.ENTER).pause(0.03).send_keys(Keys.SPACE).perform()
                        break
                    except Exception:
                        continue

                _switch_to_new_window_if_any(open_timeout=3, baseline_handles=pre_handles)

                if attachment_text:
                    ok = _smart_click_any([
                        (By.ID, "DB003002SfileChange_1"),
                        (By.CSS_SELECTOR, "[id*='fileChange']"),
                        (By.XPATH, "//*[contains(text(),'파일 수정') or contains(text(),'파일수정') or contains(@title,'파일수정')]"),
                    ], timeout=10)
                    if not ok:
                        raise TimeoutException("파일 수정 버튼을 찾지 못했습니다.")
                    time.sleep(0.2)

                    ok = _smart_click_any([
                        (By.ID, "attachFile_1"),
                        (By.CSS_SELECTOR, "[id*='attachFile']"),
                        (By.XPATH, "//*[contains(text(),'파일 선택') or contains(@title,'파일 선택')]"),
                    ], timeout=6)
                    if not ok:
                        raise TimeoutException("파일 선택 버튼을 찾지 못했습니다.")
                    time.sleep(0.2)

                    if not _ensure_file_input_and_send(pdf_path, pre_timeout=6.0):
                        raise TimeoutException("파일 선택 input 주입 실패(첨부 있음).")
                else:
                    if _ensure_file_input_and_send(pdf_path, pre_timeout=3.0):
                        logging.info("  -> 파일 선택 성공")
                    else:
                        _wait_any_appear([
                            (By.ID, "DB003002S_btnUpload"),
                            (By.CSS_SELECTOR, "[id*='btnUpload']"),
                            (By.XPATH, "//*[contains(text(),'파일 추가') or contains(text(),'파일추가') or contains(@title,'파일추가') or contains(text(),'업로드')]"),
                        ], timeout=6)

                        ok = _click_with_retry([
                            (By.ID, "DB003002S_btnUpload"),
                            (By.CSS_SELECTOR, "[id*='btnUpload']"),
                            (By.XPATH, "//*[contains(text(),'파일 추가') or contains(text(),'파일추가') or contains(@title,'파일추가') or contains(text(),'업로드')]"),
                            (By.XPATH, "//footer//button[contains(.,'파일 추가') or contains(.,'파일추가') or contains(.,'업로드')]"),
                        ], tries=3, delay=0.6)

                        if not ok:
                            try:
                                container = _find_element_all_contexts([
                                    (By.CSS_SELECTOR, "#DB003002S"),
                                    (By.CSS_SELECTOR, "body"),
                                ], timeout=1.0)
                                if container:
                                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", container)
                                    time.sleep(0.3)
                            except Exception:
                                pass

                            ok = _smart_click_any([
                                (By.ID, "DB003002S_btnUpload"),
                                (By.CSS_SELECTOR, "[id*='btnUpload']"),
                                (By.XPATH, "//*[contains(text(),'파일 추가') or contains(text(),'파일추가') or contains(@title,'파일추가') or contains(text(),'업로드')]"),
                            ], timeout=4)

                        if not ok:
                            raise TimeoutException("파일 추가 버튼을 찾지 못했습니다.")

                        if not _ensure_file_input_and_send(pdf_path, pre_timeout=4.0):
                            raise TimeoutException("파일 선택 input을 찾지 못했습니다.")

                try: driver.switch_to.default_content()
                except: pass
                _switch_to_new_window_if_any(open_timeout=1)

                ok = _smart_click_any([
                    (By.ID, "DB003002S_btnRegist"),
                    (By.XPATH, "//*[contains(text(),'저장') or contains(text(),'등록') or contains(@id,'btnRegist')]"),
                ], timeout=10)
                if not ok:
                    raise TimeoutException("저장/등록 버튼 클릭 실패")

                time.sleep(0.2)
                if not _click_ok_in_any_frame(timeout=8):
                    logging.warning("  -> [경고] 저장 확인 모달을 찾지 못했습니다.")

                _wait_modal_close_or_toast(timeout=6)

                logging.info("  -> 업로드 완료!")
                # 재렌더링 대비 살짝 위/아래 스윕
                _micro_rescan_after_refresh(lambda: None)
                return True

            except Exception as e:
                logging.error(f"  -> [경고] 업로드 처리 중 오류: {e}")
                return False

            finally:
                try:
                    if parent_handle in driver.window_handles:
                        driver.switch_to.window(parent_handle)
                except Exception:
                    pass

        # === 순차 처리 메인 루프 ===
        # 시작 expected_idx 계산: 현재 보이는 최소 인덱스부터
        idx_map, mn, mx = ({} , None, None)
        for _ in range(10):
            idx_map, mn, mx = visible_index_map()
            if mn is not None:
                break
            press_keys = ActionChains(driver).send_keys
            ActionChains(driver).send_keys(Keys.ARROW_DOWN).perform()
            time.sleep(0.1)
        if mn is None:
            logging.error(">> 보이는 데이터 행을 찾지 못했습니다.")
            return

        expected_idx = mn
        logging.info(f"-> 순차 처리 시작: expected_idx={expected_idx}")

        while is_running:
            row = seek_to_index(expected_idx)
            if row is None:
                logging.info(f"-> 인덱스 {expected_idx} 접근 실패. 종료합니다.")
                break

            k = get_row_key(row)
            if k in processed:
                # 이미 처리된 키면 다음 인덱스로
                expected_idx += 1
                continue
            processed.add(k)

            purpose_text, _ = get_text_by_class(row, "excutPrposCn")
            attach_text,  _ = get_text_by_class(row, "atchmnflNm")
            if not purpose_text:
                logging.info(f"  -> [{expected_idx}] 텍스트 없음. 건너뜁니다.")
                expected_idx += 1
                continue

            if clean_text(purpose_text) == clean_text(attach_text):
                logging.info(f"  -> [{expected_idx}] 이미 올바른 파일이 첨부되어 있습니다. 건너뜁니다.")
            else:
                trigger = find_magnifier(row)
                if not trigger:
                    logging.warning(f"  -> [{expected_idx}] 돋보기 아이콘을 찾지 못했습니다. 행 건너뜀.")
                else:
                    do_upload(purpose_text, attach_text, trigger, pdf_folder_path)

            expected_idx += 1  # 반드시 다음 인덱스로 이동

    except Exception as e:
        logging.error(f"\n[오류] 자동화 작업 중 예외가 발생했습니다: {e}")
    finally:
        is_running = False
        logging.info("\n>> 자동화 작업 스레드가 종료되었습니다. 다시 시작하려면 F3을 누르세요.")

# --- 토글/핫키 ---
def toggle_automation():
    global is_running, automation_thread, driver, is_verifying
    if driver is None:
        logging.warning(">> F3 감지: 먼저 F1로 디버그 모드 크롬에 연결해주세요.")
        return
    if is_verifying and not is_running:
        logging.warning(">> 파일 업로드 시작 불가: 현재 업로드 검증이 실행 중입니다.")
        messagebox.showwarning("실행 중", "업로드 검증이 실행 중입니다. 검증이 끝난 뒤에 파일 업로드를 시작해주세요.")
        return

    if not is_running:
        config = load_config(); pdf_path = config.get("pdf_path")
        if not pdf_path or not os.path.isdir(pdf_path):
            logging.warning(">> F3 감지: 유효한 PDF 경로가 설정되지 않았습니다. 제어판에서 경로를 저장해주세요.")
            return
        is_running = True
        logging.info("\n>> F3 감지: 자동화를 시작합니다. 중지하려면 다시 F3을 누르세요.")
        automation_thread = threading.Thread(target=main_automation_logic, args=(pdf_path,))
        automation_thread.start()
    else:
        logging.info("\n>> F3 감지: 자동화 중지 신호를 보냈습니다. 현재 작업을 마치고 중지합니다...")
        is_running = False

def start_hotkey_listener():
    keyboard.add_hotkey('f1', connect_to_existing_browser)
    keyboard.add_hotkey('f3', toggle_automation)
    keyboard.wait('esc')
    logging.info("\n>> ESC 감지: 프로그램을 종료합니다.")
    os._exit(0)

# --- 엔트리포인트 ---
if __name__ == "__main__":
    log_queue = queue.Queue()
    queue_handler = QueueHandler(log_queue)
    stream_handler = logging.StreamHandler()
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', handlers=[queue_handler, stream_handler])

    hotkey_thread = threading.Thread(target=start_hotkey_listener, daemon=True); hotkey_thread.start()

    root = tk.Tk()
    app = SettingsGUI(root, log_queue)
    root.mainloop()