"""
SAP Release Macro
- tkinter (PSF License) : 상업적 무료 사용 가능
- pyautogui (BSD 3-Clause) : 상업적 무료 사용 가능
- pywin32 / win32gui (PSF License) : 상업적 무료 사용 가능
- PyInstaller (GPL + bootloader exception) : 빌드 도구, 배포물 라이선스 무관
""" 

import sys
import time
import logging
import threading
import tkinter as tk
from tkinter import messagebox
from datetime import datetime

import win32gui
import pyautogui

# ── 로깅 설정 (콘솔 없는 EXE에서도 파일로 기록) ──────────────────────────────
logging.basicConfig(
    filename="macro_log.txt",
    level=logging.ERROR,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)


class MacroWorker:
    """매크로 동작 워커 - 별도 스레드에서 실행"""

    SAP_CLASS = "SAP_FRONTEND_SESSION"

    def __init__(self, on_update, on_finish):
        """
        on_update(status, count, time_left) : UI 업데이트 콜백
        on_finish()                          : 완료 콜백
        """
        self.on_update = on_update
        self.on_finish = on_finish
        self.is_running = False
        self._thread: threading.Thread | None = None

    # ── 외부 API ──────────────────────────────────────────────────────────────

    def start(self, repeat: int, sap_title: str):
        """매크로 시작 (이미 실행 중이면 무시)"""
        if self.is_running:
            return
        self._thread = threading.Thread(
            target=self._run,
            args=(repeat, sap_title),
            daemon=True,  # 메인 창 종료 시 자동 종료
        )
        self._thread.start()

    def stop(self):
        """매크로 중지 요청"""
        self.is_running = False

    def join(self, timeout: float = 2.0):
        """스레드 종료 대기"""
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout)

    # ── 내부 ──────────────────────────────────────────────────────────────────

    def _run(self, repeat: int, sap_title: str):
        self.is_running = True
        start_time = datetime.now()

        try:
            for i in range(repeat):
                if not self.is_running:
                    break

                # 1. SAP 창 존재 확인
                sap_hwnd = win32gui.FindWindow(self.SAP_CLASS, sap_title)
                if not sap_hwnd:
                    self.on_update("상태: SAP 창 없음", "잔여 횟수: 0", "잔여 시간: 00:00:00")
                    logging.error("SAP 창을 찾을 수 없습니다. title=%s", sap_title)
                    break

                # 2. 포커스 확보
                if win32gui.GetForegroundWindow() != sap_hwnd:
                    try:
                        win32gui.SetForegroundWindow(sap_hwnd)
                        time.sleep(0.3)
                    except Exception as e:
                        logging.error("포커스 설정 실패: %s", e)

                # 3. 키 입력 시퀀스
                try:
                    pyautogui.press("enter");   time.sleep(2)
                    pyautogui.press("f8");      time.sleep(2)
                    pyautogui.press("f5");      time.sleep(1)
                    pyautogui.hotkey("shift", "f7"); time.sleep(3)
                except Exception as e:
                    logging.error("키 입력 오류 (반복 %d/%d): %s", i + 1, repeat, e)
                    break

                # 4. 진행 상황 계산 및 UI 업데이트
                remain = repeat - i - 1
                elapsed = (datetime.now() - start_time).total_seconds()
                avg_time = elapsed / (i + 1)
                eta = int(avg_time * remain)

                h, rem = divmod(eta, 3600)
                m, s = divmod(rem, 60)
                self.on_update(
                    "상태: 실행 중",
                    f"잔여 횟수: {remain}",
                    f"잔여 시간: {h:02}:{m:02}:{s:02}",
                )

        except Exception as e:
            logging.error("예상치 못한 오류: %s", e)

        finally:
            self.is_running = False
            self.on_finish()


class ReleaseApp:
    """메인 UI"""

    SAP_TITLE = 'Activities Due for Shipping "Sales Orders, Fast Display"'

    def __init__(self, root: tk.Tk):
        self.root = root
        self.worker = MacroWorker(
            on_update=self._safe_update_labels,
            on_finish=self._safe_on_finish,
        )
        self._build_ui()
        self._fix_position()

    # ── UI 구성 ───────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.root.title("Release")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)  # 항상 최상단

        pad = {"padx": 6, "pady": 3}

        # 반복 횟수 입력
        frame_top = tk.Frame(self.root)
        frame_top.pack(**pad)

        self.entry_repeat = tk.Entry(frame_top, width=8, justify="right")
        self.entry_repeat.insert(0, "500")
        self.entry_repeat.pack(side="left")
        tk.Label(frame_top, text="번 반복").pack(side="left")

        # 시작 / 중지 버튼
        frame_btn = tk.Frame(self.root)
        frame_btn.pack(**pad)

        self.btn_start = tk.Button(frame_btn, text="시작", width=8, command=self._on_start)
        self.btn_stop  = tk.Button(frame_btn, text="중지", width=8, command=self._on_stop)
        self.btn_start.pack(side="left", padx=2)
        self.btn_stop.pack(side="left", padx=2)

        # 상태 라벨
        self.lbl_status = tk.Label(self.root, text="상태: 대기 중", anchor="w")
        self.lbl_status.pack(fill="x", **pad)

        frame_info = tk.Frame(self.root)
        frame_info.pack(**pad)
        self.lbl_count = tk.Label(frame_info, text="잔여 횟수: 0")
        self.lbl_time  = tk.Label(frame_info, text="잔여 시간: 00:00:00")
        self.lbl_count.pack(side="left", padx=4)
        self.lbl_time.pack(side="left", padx=4)

    def _fix_position(self):
        """화면 우측 70% / 상단 20% 위치로 고정"""
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = int(sw * 0.7)
        y = int(sh * 0.2)
        self.root.geometry(f"+{x}+{y}")

    # ── 이벤트 핸들러 ─────────────────────────────────────────────────────────

    def _on_start(self):
        try:
            repeat = int(self.entry_repeat.get())
            if repeat <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("입력 오류", "반복 횟수에 양의 정수를 입력하세요.")
            return

        # 이전 스레드가 살아있으면 종료 대기
        self.worker.join(timeout=2.0)

        self.btn_start.config(state="disabled")
        self.lbl_status.config(text="상태: 실행 중")
        self.worker.start(repeat, self.SAP_TITLE)

    def _on_stop(self):
        self.worker.stop()
        self.lbl_status.config(text="상태: 중지됨")
        self.btn_start.config(state="normal")

    # ── 스레드 → UI 안전 업데이트 (after 사용) ────────────────────────────────

    def _safe_update_labels(self, status: str, count: str, time_left: str):
        self.root.after(0, self._update_labels, status, count, time_left)

    def _update_labels(self, status: str, count: str, time_left: str):
        self.lbl_status.config(text=status)
        self.lbl_count.config(text=count)
        self.lbl_time.config(text=time_left)

    def _safe_on_finish(self):
        self.root.after(0, self._on_finish)

    def _on_finish(self):
        """매크로 정상 완료 시 버튼 복원"""
        self.btn_start.config(state="normal")
        self.lbl_status.config(text="상태: 완료")


def main():
    root = tk.Tk()
    ReleaseApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
