# Task-Timer
import tkinter as tk
from tkinter import ttk
import time
from datetime import datetime
import pandas as pd
from tkinter import filedialog, messagebox, simpledialog

class TaskManager:
    def __init__(self, root):
        self.root = root
        self.root.title("タスク管理タイマー")

        # スタイルの設定
        style = ttk.Style()
        style.theme_use('default')

        # カスタムカラー
        primary_color = "#4a7a8c"
        secondary_color = "#f0f4f7"
        accent_color = "#dceefb"
        text_color = "#333333"
        button_hover_color = "#5b8ca1"
        button_focus_color = "#729fb8"
        focus_background_color = "#3b6a78"  # フォーカス時の背景色
        focus_text_color = "#ffffff"  # フォーカス時のテキスト色

        style.configure('TFrame', background=secondary_color)
        style.configure('TLabel', background=secondary_color, foreground=text_color, font=("メイリオ", 10))
        style.configure('Header.TLabel', font=("メイリオ", 12, 'bold'))

        # ボタンのスタイル設定
        style.configure('Custom.TButton',
                        background=primary_color,
                        foreground='white',
                        font=("メイリオ", 10),
                        borderwidth=0,
                        focusthickness=3,
                        focuscolor=button_focus_color,
                        padding=6)
        style.map('Custom.TButton',
                  background=[
                      ('active', button_hover_color),
                      ('focus', focus_background_color)
                  ],
                  foreground=[
                      ('focus', focus_text_color)
                  ]
                  )

        # ウィンドウを右端に表示し、縦方向いっぱいに設定
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = 500  # 横幅を450から500に変更
        window_height = int(screen_height * 0.9)
        x_position = screen_width - window_width - 20
        y_position = int((screen_height - window_height) / 2)
        self.root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # ウィンドウのリサイズ許可
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)

        # キャンバスを作成
        canvas = tk.Canvas(root)
        canvas.grid(row=0, column=0, sticky="nsew")

        # スクロールバーを作成
        scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")

        canvas.configure(yscrollcommand=scrollbar.set)

        # メインフレームをキャンバスに配置
        main_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=main_frame, anchor='nw')

        # フレームのサイズが変更されたときにスクロール領域を更新
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        main_frame.bind("<Configure>", on_frame_configure)

        self.tasks = {}
        self.current_task = None
        self.start_time = None
        self.records = []
        self.alert_time = 10 * 60
        self.flash_duration = 10
        self.excel_file_path = None
        self.text_color = text_color  # テキストカラーをクラス変数として保存
        self.is_saved = True  # データが保存されているかのフラグ

        # タスクスタックの初期化
        self.task_stack = []

        # メインフレームのグリッド設定
        main_frame.rowconfigure(list(range(15)), weight=1)
        main_frame.columnconfigure(0, weight=1)

        # タスク入力用ラベル
        self.label = ttk.Label(
            main_frame,
            text="1. 今日のタスクを入力してください。",
            style='Header.TLabel'
        )
        self.label.grid(row=0, column=0, pady=5, sticky="w")

        self.batch_entry = tk.Text(main_frame, height=3, width=40, font=("メイリオ", 10))
        self.batch_entry.grid(row=1, column=0, pady=5, sticky="nsew")
        self.batch_entry.focus_set()

        # タブキーのカスタマイズ
        self.batch_entry.bind('<Tab>', self.focus_next_widget)

        self.add_batch_task_button = ttk.Button(
            main_frame,
            text="一括タスク追加",
            command=self.add_batch_tasks,
            style='Custom.TButton'
        )
        self.add_batch_task_button.grid(row=2, column=0, pady=5, sticky="ew")

        # 今日のタスク一覧ラベル
        self.task_list_label = ttk.Label(
            main_frame,
            text="2. 作業するタスクを選択してください。",
            style='Header.TLabel'
        )
        self.task_list_label.grid(row=3, column=0, pady=5, sticky="w")

        # タスクリストボックス
        self.task_listbox = tk.Listbox(
            main_frame,
            height=3,
            width=50,
            exportselection=False,
            font=("メイリオ", 10)
        )
        self.task_listbox.grid(row=4, column=0, pady=5, sticky="nsew")

        # リストボックスの選択イベントにバインド
        self.task_listbox.bind('<<ListboxSelect>>', self.on_task_select)

        # 目標時間入力用ラベルとエントリ
        self.target_time_label = ttk.Label(
            main_frame,
            text="3. 目標作業時間（分）を入力してください。",
            style='Header.TLabel'
        )
        self.target_time_label.grid(row=5, column=0, pady=5, sticky="w")

        self.target_time_entry = ttk.Entry(main_frame, width=10, font=("メイリオ", 10))
        self.target_time_entry.grid(row=6, column=0, pady=5, sticky="w")
        self.target_time_entry.insert(0, "10")

        # ボタンをフレームに配置
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=7, column=0, pady=10, sticky="ew")
        button_frame.columnconfigure([0, 1, 2], weight=1)

        self.start_timer_button = ttk.Button(
            button_frame,
            text="タイマー開始",
            command=self.start_timer,
            style='Custom.TButton'
        )
        self.start_timer_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.complete_task_button = ttk.Button(
            button_frame,
            text="作業終了",
            command=self.complete_task,
            style='Custom.TButton'
        )
        self.complete_task_button.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # 割り込みタスクボタンの配置
        self.interrupt_task_button = ttk.Button(
            button_frame,
            text="割り込みタスク",
            command=self.start_interrupt_task,
            style='Custom.TButton'
        )
        self.interrupt_task_button.grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        # 現在のタスク名を表示するラベル
        self.current_task_label = ttk.Label(
            main_frame,
            text="現在のタスク: なし",
            font=("メイリオ", 12, 'bold')
        )
        self.current_task_label.grid(row=8, column=0, pady=5, sticky="w")

        self.timer_label = ttk.Label(
            main_frame,
            text="作業時間: 0:00",
            font=("メイリオ", 16, 'bold')
        )
        self.timer_label.grid(row=9, column=0, pady=5, sticky="w")

        self.is_timing = False

        # 作業時間を表示するテキストエリア
        self.time_log_label = ttk.Label(
            main_frame,
            text="作業時間記録:",
            style='Header.TLabel'
        )
        self.time_log_label.grid(row=10, column=0, pady=5, sticky="w")

        self.time_log = tk.Text(
            main_frame,
            height=3,
            width=60,
            state=tk.DISABLED,
            font=("メイリオ", 10)
        )
        self.time_log.grid(row=11, column=0, pady=5, sticky="nsew")

        # 操作方法の説明ラベル
        self.instructions_label = ttk.Label(
            main_frame,
            text="ショートカットキーの操作方法:",
            font=("メイリオ", 10, 'bold')
        )
        self.instructions_label.grid(row=12, column=0, pady=5, sticky="w")

        instructions_text = (
            "・タスクの選択: タスクリストで ↑ / ↓ キーを使用\n"
            "・目標時間を入力後、Enter キーでタイマー開始\n"
            "・作業終了: Delete キー\n"
            "・ボタンにフォーカスがある場合、スペースキーでボタンを押せます"
        )

        self.instructions_detail = ttk.Label(
            main_frame,
            text=instructions_text,
            font=("メイリオ", 10),
            justify='left'
        )
        self.instructions_detail.grid(row=13, column=0, pady=5, sticky="w")

        # キーボード操作のバインド
        self.target_time_entry.bind('<Return>', lambda event: self.start_timer())
        # 修正: bind_all から bind に変更
        self.root.bind('<Delete>', self.handle_key_event)

        # ウィンドウクローズイベントにバインド
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def focus_next_widget(self, event):
        self.root.focus_get().tk_focusNext().focus()
        return 'break'

    def handle_key_event(self, event):
        if event.keysym == 'Delete':
            focused_widget = self.root.focus_get()
            if isinstance(focused_widget, ttk.Button):
                # ボタンがフォーカスされている場合はデフォルトの動作を許可
                return
            else:
                self.complete_task()
                return 'break'  # デフォルトの動作を防ぐ

    def add_batch_tasks(self):
        tasks_text = self.batch_entry.get("1.0", tk.END).strip()
        tasks = tasks_text.splitlines()
        for task in tasks:
            task = task.strip()
            if task and task not in self.tasks:
                self.tasks[task] = 0
                self.task_listbox.insert(tk.END, task)
        self.batch_entry.delete("1.0", tk.END)

    def on_task_select(self, event):
        if self.is_timing:
            # タイマー動作中はタスクの変更を無視
            return
        selected_indices = self.task_listbox.curselection()
        if selected_indices:
            index = selected_indices[0]
            self.current_task = self.task_listbox.get(index)
            self.current_task_label.config(text=f"現在のタスク: {self.current_task}")
        else:
            self.current_task = None
            self.current_task_label.config(text="現在のタスク: なし")

    def start_timer(self):
        if self.current_task is None:
            messagebox.showinfo("情報", "タスクが選択されていません。")
            return
        if self.is_timing:
            messagebox.showinfo("情報", "既にタイマーが動作中です。")
            return
        # タイマー開始
        self.start_time = time.time()
        try:
            self.alert_time = int(self.target_time_entry.get()) * 60
        except ValueError:
            self.alert_time = 10 * 60
        self.is_timing = True
        self.update_timer()

    def start_interrupt_task(self):
        if not self.is_timing:
            messagebox.showinfo("情報", "現在タイマーが動作中のタスクがありません。")
            return
        # 現在のタスクの状態をスタックに保存
        self.task_stack.append({
            'task': self.current_task,
            'start_time': self.start_time
        })
        # 現在のタスクを一時停止し、作業時間を記録
        self.pause_current_task()
        # 割り込みタスクの名前を取得
        interrupt_task_name = simpledialog.askstring("割り込みタスク", "割り込みタスクの名前を入力してください:")
        if interrupt_task_name:
            if interrupt_task_name not in self.tasks:
                self.tasks[interrupt_task_name] = 0
                self.task_listbox.insert(tk.END, interrupt_task_name)
                # ★割り込みタスクを自動的に選択する
                index = self.task_listbox.size() - 1
            else:
                # 既に存在する場合、そのインデックスを取得
                index = self.task_listbox.get(0, tk.END).index(interrupt_task_name)
            # タスクリストで割り込みタスクを選択状態にする
            self.task_listbox.selection_clear(0, tk.END)
            self.task_listbox.selection_set(index)
            self.task_listbox.activate(index)
            self.task_listbox.see(index)

            # 割り込みタスクを開始
            self.current_task = interrupt_task_name
            self.current_task_label.config(text=f"現在のタスク: {self.current_task}")
            self.start_time = time.time()
            self.is_timing = True
            self.update_timer()
        else:
            # キャンセルされた場合、前のタスクを再開
            self.resume_previous_task()

    def pause_current_task(self):
        # 現在のタスクの時間を計測して記録
        elapsed_time = time.time() - self.start_time
        self.tasks[self.current_task] += elapsed_time
        minutes, seconds = divmod(elapsed_time, 60)
        elapsed_str = f"{int(minutes)}:{int(seconds):02d}"
        self.log_time(self.current_task, elapsed_str)
        self.is_timing = False

    def resume_previous_task(self):
        if self.task_stack:
            # スタックから前のタスクの状態を取得
            previous_task_info = self.task_stack.pop()
            self.current_task = previous_task_info['task']
            self.start_time = time.time()
            self.current_task_label.config(text=f"現在のタスク: {self.current_task}")
            # タスクリストで前のタスクを選択状態にする
            index = self.task_listbox.get(0, tk.END).index(self.current_task)
            self.task_listbox.selection_clear(0, tk.END)
            self.task_listbox.selection_set(index)
            self.task_listbox.activate(index)
            self.task_listbox.see(index)

            self.is_timing = True
            self.update_timer()
        else:
            self.current_task = None
            self.current_task_label.config(text="現在のタスク: なし")
            self.is_timing = False

    def stop_timer(self):
        if self.is_timing:
            self.pause_current_task()
            self.current_task = None
            self.current_task_label.config(text="現在のタスク: なし")
            self.is_timing = False
        else:
            messagebox.showinfo("情報", "タイマーが動作していません。")

    def complete_task(self):
        if self.is_timing:
            self.stop_timer()
        selected_task_index = self.task_listbox.curselection()
        if selected_task_index:
            task_name = self.task_listbox.get(selected_task_index)
            self.task_listbox.delete(selected_task_index)
            del self.tasks[task_name]
            # スタックから削除
            self.task_stack = [task_info for task_info in self.task_stack if task_info['task'] != task_name]
            if self.current_task == task_name:
                self.current_task = None
                self.current_task_label.config(text="現在のタスク: なし")
        else:
            messagebox.showinfo("情報", "タスクが選択されていません。")
        self.export_to_excel()

    def update_timer(self):
        if self.is_timing:
            current_time = time.time()
            elapsed_time = current_time - self.start_time
            minutes, seconds = divmod(elapsed_time, 60)
            self.timer_label.config(text=f"作業時間: {int(minutes)}:{int(seconds):02d}")

            # 目標時間を超えた場合にタイマーを赤字にして点滅させる
            if elapsed_time > self.alert_time:
                flash_time = elapsed_time - self.alert_time
                if flash_time <= self.flash_duration:
                    current_color = self.timer_label.cget("foreground")
                    new_color = "red" if current_color != "red" else self.text_color
                    self.timer_label.config(foreground=new_color)
                else:
                    self.timer_label.config(foreground="red")
            else:
                self.timer_label.config(foreground=self.text_color)

            self.root.after(1000, self.update_timer)

    def log_time(self, task, elapsed_str):
        today_date = datetime.now().strftime("%Y-%m-%d")
        start_time_str = datetime.fromtimestamp(self.start_time).strftime("%Y-%m-%d %H:%M:%S")
        end_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_log.config(state=tk.NORMAL)
        log_entry = (
            f"{today_date} - タスク: {task}\n"
            f"  作業時間: {elapsed_str}\n"
            f"  開始: {start_time_str}, 終了: {end_time_str}\n\n"
        )
        self.time_log.insert(tk.END, log_entry)
        self.time_log.config(state=tk.DISABLED)
        self.records.append({
            "Date": today_date,
            "Task": task,
            "Duration": elapsed_str,
            "Start Time": start_time_str,
            "End Time": end_time_str
        })
        self.is_saved = False

    def export_to_excel(self):
        try:
            if not self.records:
                print("作業記録がありません。")
                return

            if self.excel_file_path is None:
                today_date = datetime.now().strftime("%Y-%m-%d")
                default_filename = f"{today_date}_作業記録.xlsx"

                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=default_filename,
                    filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")]
                )

                if not file_path:
                    print("保存がキャンセルされました。")
                    return

                self.excel_file_path = file_path
            else:
                file_path = self.excel_file_path

            df = pd.DataFrame(self.records)
            df.to_excel(file_path, index=False)
            print(f"エクセルに転送しました: {file_path}")
            self.is_saved = True  # データが保存されたことを示す

        except Exception as e:
            print(f"エクセルへの転送中にエラーが発生しました: {e}")

    def on_closing(self):
        if self.records and not self.is_saved:
            if messagebox.askyesno("保存確認", "作業記録が保存されていません。保存しますか？"):
                self.export_to_excel()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    task_manager = TaskManager(root)
    root.mainloop()
