import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import os
import shutil
import threading
import time
from datetime import datetime, timedelta
from queue import Queue, Empty
import winsound
import sys
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

class FileTimerApp:
    def init_timer_state(self):
        self.is_paused = False
        self.is_counting = False
        self._pause_event = threading.Event()
        self._pause_event.set()

    def resource_path(self, relative_path):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.abspath(relative_path)

    def play_alarm(self):
        """播放計時結束音效"""
        try:
            winsound.PlaySound(self.resource_path("finish.wav"), winsound.SND_FILENAME | winsound.SND_ASYNC)
        except Exception as e:
            print(f"音效播放失敗: {e}")

    PREVIEW_MAX_SIZE = (420, 250)

    def __init__(self, root):
        self.root = root
        self.root.title("自動計時丟檔工具")
        # 不設固定高度，讓內容自動撐開
        self.root.resizable(True, True) # 允許視窗縮放
        self.set_icon()

        self.source_file_path = tk.StringVar()
        self.dest_dir_path = tk.StringVar()
        
        self.gui_queue = Queue()

        self.init_timer_state()
        self.create_widgets()
        self.setup_initial_state()
    def set_spinbox_to_now(self):
        now = datetime.now()
        self.hour_spinbox.set(f"{now.hour:02}")
        self.minute_spinbox.set(f"{now.minute:02}")

        if not PIL_AVAILABLE:
            messagebox.showwarning("缺少函式庫", "未找到 Pillow 函式庫 (PIL)。\n圖片預覽功能將無法使用。\n\n請執行 'pip install Pillow' 來安裝。")

    def set_icon(self):
        """設定視窗圖示 (使用 clock.ico，支援 PyInstaller 單檔)"""
        try:
            icon_path = self.resource_path("clock.ico")
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"圖示載入失敗: {e}")


    def setup_initial_state(self):
        """設定程式初始狀態並啟動佇列處理器"""
        self.reset_state()
        self.process_queue() # 啟動佇列檢查

    def process_queue(self):
        """處理從背景執行緒傳來的GUI更新請求"""
        try:
            task = self.gui_queue.get_nowait()
            task()
        except Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)

    def create_widgets(self):
        """建立所有GUI元件"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 來源檔案選擇 ---
        source_frame = ttk.LabelFrame(main_frame, text="1. 選擇來源檔案")
        source_frame.pack(fill=tk.X, pady=5)
        ttk.Label(source_frame, textvariable=self.source_file_path, foreground="gray").pack(pady=5, padx=10, anchor="w")
        ttk.Button(source_frame, text="選擇檔案", command=self.select_source_file).pack(pady=5, padx=10, anchor="w")

        # --- 圖片預覽 (固定高度) ---
        preview_frame = ttk.LabelFrame(main_frame, text="圖片預覽")
        preview_frame.pack(fill=tk.X, pady=5)
        image_container = ttk.Frame(preview_frame, height=260)
        image_container.pack(fill=tk.X)
        image_container.pack_propagate(False)
        self.preview_label = ttk.Label(image_container, anchor="center")
        self.preview_label.pack(fill=tk.BOTH, expand=True)
        
        # --- 電子鐘顯示區 ---
        clock_frame = tk.Frame(main_frame, bg="black", relief="sunken", bd=5)
        clock_frame.pack(fill=tk.X, pady=10)
        self.clock_label = tk.Label(clock_frame, text="00:00:00", bg="black", fg="#39FF14", font=("DS-Digital", 90, "bold")) # 修改這裡的 90 即可改變高度
        self.clock_label.pack(pady=5)

        # --- 時間設定 ---
        time_frame = ttk.LabelFrame(main_frame, text="2. 設定丟出時間 (24小時制)")
        time_frame.pack(fill=tk.X, pady=5)
        time_inner_frame = ttk.Frame(time_frame)
        time_inner_frame.pack(pady=5, padx=10, anchor="w")
        self.hour_spinbox = ttk.Spinbox(time_inner_frame, from_=0, to=23, wrap=True, width=5, format="%02.0f")
        self.hour_spinbox.pack(side=tk.LEFT)
        ttk.Label(time_inner_frame, text="時").pack(side=tk.LEFT, padx=(2, 10))
        self.minute_spinbox = ttk.Spinbox(time_inner_frame, from_=0, to=59, wrap=True, width=5, format="%02.0f")
        self.minute_spinbox.pack(side=tk.LEFT)
        ttk.Label(time_inner_frame, text="分").pack(side=tk.LEFT, padx=2)

        # --- 目的地路徑選擇 ---
        dest_frame = ttk.LabelFrame(main_frame, text="3. 選擇目的地資料夾")
        dest_frame.pack(fill=tk.X, pady=5)
        ttk.Label(dest_frame, textvariable=self.dest_dir_path, foreground="gray").pack(pady=5, padx=10, anchor="w")
        ttk.Button(dest_frame, text="選擇目的地", command=self.select_dest_dir).pack(pady=5, padx=10, anchor="w")

        # --- 執行與狀態 ---
        self.start_button = tk.Button(main_frame, text="開始計時", command=self.start_or_pause_timer, font=("Arial", 24, "bold"))
        self.start_button.pack(pady=10, fill=tk.X)
    def start_or_pause_timer(self):
        if not self.is_counting:
            self.start_timer_thread()
        else:
            if self.is_paused:
                self.resume_timer()
            else:
                self.pause_timer()

    def pause_timer(self):
        self.is_paused = True
        self._pause_event.clear()
        self.start_button.config(text="繼續")

    def resume_timer(self):
        self.is_paused = False
        self._pause_event.set()
        self.start_button.config(text="暫停")

    def select_source_file(self):
        """開啟檔案對話框，讓使用者選擇檔案，並顯示預覽"""
        filepath = filedialog.askopenfilename(
            title="選擇一個檔案"
        )
        if not filepath:
            return

        self.source_file_path.set(filepath)

        if not PIL_AVAILABLE:
            self.preview_label.config(text="Pillow 函式庫未安裝，無法預覽")
            return

        try:
            image = Image.open(filepath)
            image.thumbnail(self.PREVIEW_MAX_SIZE)
            photo = ImageTk.PhotoImage(image)
            
            self.preview_label.config(image=photo, text="")
            self.preview_label.image = photo
        except Exception:
            self.preview_label.config(image=None, text="非圖片檔案或無法預覽")
            self.preview_label.image = None

    def select_dest_dir(self):
        """開啟目錄對話框，讓使用者選擇目的地"""
        dirpath = filedialog.askdirectory(title="選擇目的地資料夾")
        if dirpath:
            self.dest_dir_path.set(dirpath)

    def start_timer_thread(self):
        """啟動計時器執行緒"""
        if not os.path.exists(self.source_file_path.get()) or self.dest_dir_path.get() == "尚未選擇目的地":
            messagebox.showwarning("資訊不完整", "請先選擇來源檔案和目的地資料夾。\n")
            return
        # 檢查 spinbox 值是否為空或不合理，若是則自動設為現在時間
        try:
            hour = int(self.hour_spinbox.get())
            minute = int(self.minute_spinbox.get())
        except Exception:
            self.set_spinbox_to_now()
        self.is_counting = True
        self.is_paused = False
        self._pause_event.set()
        self.start_button.config(text="暫停")
        timer_thread = threading.Thread(target=self.run_countdown_and_move, daemon=True)
        timer_thread.start()

    def run_countdown_and_move(self):
        """在背景執行倒數計時並移動檔案"""
        try:
            source = self.source_file_path.get()
            dest = self.dest_dir_path.get()
            hour = int(self.hour_spinbox.get())
            minute = int(self.minute_spinbox.get())

            now = datetime.now()
            target_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
            if target_time < now:
                target_time += timedelta(days=1)

            while datetime.now() < target_time:
                self._pause_event.wait()  # 若暫停則阻塞
                remaining = target_time - datetime.now()
                remaining_seconds = int(remaining.total_seconds())
                h, rem = divmod(remaining_seconds, 3600)
                m, s = divmod(rem, 60)
                clock_update_func = lambda h=h, m=m, s=s: self.clock_label.config(text=f"{h:02}:{m:02}:{s:02}")
                self.gui_queue.put(clock_update_func)
                time.sleep(1)

            self.gui_queue.put(lambda: [self.clock_label.config(text="00:00:00"), self.play_alarm()])
            filename = os.path.basename(source)
            shutil.move(source, os.path.join(dest, filename))
            completion_time = datetime.now().strftime("%H:%M")
            success_func = lambda fname=filename, d=dest, time=completion_time: messagebox.showinfo(
                "任務完成", f"檔案 '{fname}' 已於 {time} 成功移動到\n'{d}'"
            )
            self.gui_queue.put(success_func)
            self.gui_queue.put(self.reset_state)

        except FileNotFoundError:
            self.gui_queue.put(lambda: messagebox.showerror("錯誤", f"來源檔案不存在：\n{source}\n\n請重新選擇檔案。"))
            self.gui_queue.put(self.reset_state)
        except Exception as e:
            self.gui_queue.put(lambda: messagebox.showerror("發生錯誤", f"發生預期外的錯誤：\n{e}"))
            self.gui_queue.put(self.reset_state)

    def reset_state(self):
        """重設GUI到初始狀態"""
        self.source_file_path.set("尚未選擇檔案")
        self.dest_dir_path.set("尚未選擇目的地")
        self.set_spinbox_to_now()
        self.is_counting = False
        self.is_paused = False
        self._pause_event.set()
        self.start_button.config(text="開始計時", state=tk.NORMAL)
        self.clock_label.config(text="00:00:00")
        
        if PIL_AVAILABLE:
            self.preview_label.config(image=None, text="圖片預覽")
            self.preview_label.image = None

if __name__ == "__main__":
    root = tk.Tk()
    app = FileTimerApp(root)
    root.mainloop()