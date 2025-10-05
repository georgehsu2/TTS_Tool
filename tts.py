"""
Chat-Style TTS Panel (Windows/macOS/Linux)
- Tkinter GUI with chat transcript
- pyttsx3 per-utterance engine to avoid 'only speaks once' issues on Windows
- Enter to send & speak; Shift+Enter = newline
- Keeps history; can replay, stop, delete selected, export transcript
- No extra system installs required (besides `pip install pyttsx3`)

Windows → Send audio to Discord without extra drivers:
- If your audio driver exposes "Stereo Mix", set Discord Input Device = Stereo Mix
- Otherwise you will need a virtual audio cable (not included here by request)
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import pyttsx3
import time
from datetime import datetime
from dataclasses import dataclass

# ------------------------ Speech worker ------------------------
class SpeechWorker(threading.Thread):
    def __init__(self, text: str, voice_id: str | None, rate: int, volume: float):
        super().__init__(daemon=True)
        self.text = text
        self.voice_id = voice_id
        self.rate = rate
        self.volume = volume
        self.stop_flag = threading.Event()
        self.exc: Exception | None = None

    def run(self):
        engine = None
        try:
            engine = pyttsx3.init()
            if self.voice_id:
                try:
                    engine.setProperty("voice", self.voice_id)
                except Exception:
                    pass
            try:
                engine.setProperty("rate", int(self.rate))
            except Exception:
                pass
            try:
                engine.setProperty("volume", float(self.volume))
            except Exception:
                pass

            def on_word(name, location, length):
                if self.stop_flag.is_set():
                    raise SystemExit
            engine.connect('started-word', on_word)

            engine.say(self.text)
            engine.runAndWait()
            engine.stop()
        except SystemExit:
            try:
                if engine:
                    engine.stop()
            except Exception:
                pass
        except Exception as e:
            self.exc = e
        finally:
            try:
                del engine
            except Exception:
                pass

    def stop(self):
        self.stop_flag.set()

# ------------------------ Data model --------------------------
@dataclass
class Message:
    id: int
    ts: float
    text: str

# ------------------------ Main App ----------------------------
class ChatTTSApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("橘mouth v1.0")
        self.geometry("760x560")
        self.minsize(680, 520)

        # one-shot probe engine to get defaults & voices
        try:
            _tmp = pyttsx3.init()
            self.voices = _tmp.getProperty("voices") or []
            self.default_rate = _tmp.getProperty("rate") or 180
            self.default_volume = _tmp.getProperty("volume") or 1.0
            # choose a likely Chinese voice if available
            default_voice = None
            for v in self.voices:
                name = (getattr(v, 'name', '') or '').lower()
                langs = ','.join(getattr(v, 'languages', []) or []).lower()
                if any(k in (name + ' ' + langs) for k in ['zh', 'cmn', 'mandarin', 'tw', 'chinese']):
                    default_voice = v
                    break
            if default_voice is None and self.voices:
                default_voice = self.voices[0]
            self.default_voice_id = default_voice.id if default_voice else ''
            try:
                _tmp.stop()
            except Exception:
                pass
            del _tmp
        except Exception as e:
            messagebox.showerror("初始化失敗", f"無法初始化系統語音：\n{e}")
            self.voices = []
            self.default_rate = 180
            self.default_volume = 1.0
            self.default_voice_id = ''

        # state
        self.messages: list[Message] = []
        self.auto_speak_var = tk.BooleanVar(value=True)
        self.current_worker: SpeechWorker | None = None
        self.next_id = 1

        # build UI
        self._build_ui()
        self.after(150, lambda: self.txt_input.focus_set())

    # -------------------- UI --------------------
    def _build_ui(self):
        # top controls
        top = ttk.Frame(self)
        top.pack(fill='x', padx=12, pady=(12, 6))

        ttk.Label(top, text='語者/Voice：').grid(row=0, column=0, sticky='w')
        self.cmb_voice = ttk.Combobox(top, state='readonly', width=40)
        voice_display = []
        for v in self.voices:
            label = getattr(v, 'name', v.id)
            langs = getattr(v, 'languages', [])
            if langs:
                label += f"  ({', '.join(langs)})"
            voice_display.append((label, v.id))
        self.cmb_voice['values'] = [label for (label, _id) in voice_display]
        self.cmb_voice.grid(row=0, column=1, sticky='ew', padx=(6, 12))
        if voice_display:
            # select default
            idx = 0
            for i, (_, vid) in enumerate(voice_display):
                if vid == self.default_voice_id:
                    idx = i
                    break
            self.cmb_voice.current(idx)
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text='語速：').grid(row=0, column=2, sticky='w')
        self.sld_rate = ttk.Scale(top, from_=80, to=260, orient='horizontal')
        self.sld_rate.set(self.default_rate)
        self.sld_rate.grid(row=0, column=3, sticky='ew', padx=(6, 12))

        ttk.Label(top, text='音量：').grid(row=0, column=4, sticky='w')
        self.sld_volume = ttk.Scale(top, from_=0, to=100, orient='horizontal')
        self.sld_volume.set(int(self.default_volume * 100))
        self.sld_volume.grid(row=0, column=5, sticky='ew', padx=(6, 0))

        top.columnconfigure(3, weight=1)
        top.columnconfigure(5, weight=1)

        # transcript area
        mid = ttk.Frame(self)
        mid.pack(fill='both', expand=True, padx=12, pady=(0, 6))

        self.txt_log = tk.Text(mid, wrap='word', state='disabled')
        self.txt_log.pack(fill='both', expand=True, side='left')
        self.scroll = ttk.Scrollbar(mid, command=self.txt_log.yview)
        self.scroll.pack(fill='y', side='right')
        self.txt_log['yscrollcommand'] = self.scroll.set

        # input + buttons
        bottom = ttk.Frame(self)
        bottom.pack(fill='x', padx=12, pady=(0, 12))

        self.txt_input = tk.Text(bottom, height=3, wrap='word')
        self.txt_input.pack(fill='x', expand=True, side='left')
        self.txt_input.bind('<Return>', self._on_enter_send)
        self.txt_input.bind('<Shift-Return>', self._on_shift_enter_newline)

        side = ttk.Frame(bottom)
        side.pack(side='right', padx=(8, 0))
        ttk.Button(side, text='送出並朗讀', command=self.send_message).pack(fill='x')
        ttk.Button(side, text='停止', command=self.stop_speaking).pack(fill='x', pady=(6, 0))
        ttk.Checkbutton(side, text='送出自動朗讀', variable=self.auto_speak_var).pack(anchor='w', pady=(6, 0))
        ttk.Button(side, text='重播上一則', command=self.replay_last).pack(fill='x', pady=(6, 0))
        ttk.Button(side, text='匯出紀錄', command=self.export_log).pack(fill='x', pady=(6, 0))
        ttk.Button(side, text='清除紀錄', command=self.clear_log).pack(fill='x', pady=(6, 0))

        # status bar
        self.lbl_status = ttk.Label(self, text='狀態：待命')
        self.lbl_status.pack(side='bottom', pady=(0, 6))

        # context menu for log
        self._build_log_menu()

    def _build_log_menu(self):
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label='重播此則', command=self._ctx_replay)
        self.menu.add_command(label='複製選取', command=lambda: self.txt_log.event_generate('<<Copy>>'))
        self.menu.add_separator()
        self.menu.add_command(label='刪除此則', command=self._ctx_delete)
        self.txt_log.bind('<Button-3>', self._on_log_right_click)

    # -------------------- Helpers --------------------
    def _current_voice_id(self):
        label = self.cmb_voice.get()
        if not label or not self.voices:
            return None
        for v in self.voices:
            name = getattr(v, 'name', v.id)
            langs = getattr(v, 'languages', [])
            shown = name + (f"  ({', '.join(langs)})" if langs else '')
            if shown == label:
                return v.id
        return self.voices[0].id if self.voices else None

    def _append_log(self, msg: Message):
        ts_str = datetime.fromtimestamp(msg.ts).strftime('%H:%M:%S')
        line = f"[{msg.id:03d} {ts_str}] {msg.text}\n"
        self.txt_log.configure(state='normal')
        self.txt_log.insert('end', line)
        self.txt_log.configure(state='disabled')
        self.txt_log.see('end')

    def _replace_log_all(self, text: str):
        self.txt_log.configure(state='normal')
        self.txt_log.delete('1.0', 'end')
        self.txt_log.insert('end', text)
        self.txt_log.configure(state='disabled')

    def _set_status(self, s: str):
        self.lbl_status.config(text=f'狀態：{s}')
        self.update_idletasks()

    # -------------------- Actions --------------------
    def send_message(self):
        text = self.txt_input.get('1.0', 'end').strip()
        if not text:
            return
        msg = Message(id=self.next_id, ts=time.time(), text=text)
        self.next_id += 1
        self.messages.append(msg)
        self._append_log(msg)
        # 保留內容並自動全選方便覆寫
        self.txt_input.focus_set()
        self.txt_input.tag_add('sel', '1.0', 'end')

        if self.auto_speak_var.get():
            self._speak_text(text)

    def _speak_text(self, text: str):
        # stop current if any
        self.stop_speaking(join=False)
        voice_id = self._current_voice_id()
        rate = int(self.sld_rate.get())
        volume = max(0.0, min(1.0, float(self.sld_volume.get()) / 100.0))
        worker = SpeechWorker(text, voice_id, rate, volume)
        self.current_worker = worker
        self._set_status('朗讀中…')
        worker.start()
        self.after(120, self._poll_worker_done)

    def stop_speaking(self, join=True):
        if self.current_worker and self.current_worker.is_alive():
            self.current_worker.stop()
            if join:
                self.current_worker.join(timeout=0.5)
        self._set_status('已停止')

    def replay_last(self):
        if not self.messages:
            return
        self._speak_text(self.messages[-1].text)

    def export_log(self):
        if not self.messages:
            messagebox.showinfo('匯出', '目前沒有紀錄可匯出。')
            return
        path = filedialog.asksaveasfilename(
            title='匯出聊天紀錄', defaultextension='.txt',
            filetypes=[('Text Files', '*.txt'), ('All Files', '*.*')]
        )
        if not path:
            return
        with open(path, 'w', encoding='utf-8') as f:
            for m in self.messages:
                ts = datetime.fromtimestamp(m.ts).strftime('%Y-%m-%d %H:%M:%S')
                f.write(f"[{m.id:03d} {ts}] {m.text}\n")
        messagebox.showinfo('匯出', f'已儲存到:\n{path}')

    def clear_log(self):
        if not self.messages:
            return
        if messagebox.askyesno('清除紀錄', '確定要清除畫面上的聊天紀錄嗎？此動作不會復原。'):
            self.messages.clear()
            self._replace_log_all('')

    # -------------------- Context menu handlers --------------------
    def _on_log_right_click(self, event):
        try:
            self.rc_click_index = self.txt_log.index(f"@{event.x},{event.y}")
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def _ctx_replay(self):
        # try to get the line text under cursor
        try:
            line_start = self.txt_log.index(f"{self.rc_click_index} linestart")
            line_end = self.txt_log.index(f"{self.rc_click_index} lineend")
            line = self.txt_log.get(line_start, line_end)
            # strip prefix like [001 12:34:56]
            if ']' in line:
                text = line.split(']', 1)[1].strip()
            else:
                text = line.strip()
            if text:
                self._speak_text(text)
        except Exception:
            pass

    def _ctx_delete(self):
        # remove from memory and UI by matching id at line head
        try:
            line_start = self.txt_log.index(f"{self.rc_click_index} linestart")
            line_end = self.txt_log.index(f"{self.rc_click_index} lineend")
            line = self.txt_log.get(line_start, line_end)
            if line.startswith('[') and ']' in line:
                id_part = line.split(']', 1)[0]
                try:
                    msg_id = int(id_part.split()[0][1:])
                    self.messages = [m for m in self.messages if m.id != msg_id]
                except Exception:
                    pass
            # remove line from UI
            self.txt_log.configure(state='normal')
            self.txt_log.delete(line_start, f"{line_end}+1c")
            self.txt_log.configure(state='disabled')
        except Exception:
            pass

    # -------------------- event helpers --------------------
    def _on_enter_send(self, event):
        # Enter = send & speak, Shift+Enter = newline
        if event.state & 0x0001:  # Shift pressed
            return
        self.send_message()
        return "break"  # prevent newline

    def _on_shift_enter_newline(self, event):
        self.txt_input.insert('insert', '\n')
        return "break"

    def _poll_worker_done(self):
        w = self.current_worker
        if not w:
            return
        if w.is_alive():
            self.after(120, self._poll_worker_done)
        else:
            if w.exc:
                messagebox.showerror('朗讀失敗', f'{w.exc}')
            self.current_worker = None
            self._set_status('待命')

    def _build_log_menu(self):
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label='重播此則', command=self._ctx_replay)
        self.menu.add_command(label='複製選取', command=lambda: self.txt_log.event_generate('<<Copy>>'))
        self.menu.add_separator()
        self.menu.add_command(label='刪除此則', command=self._ctx_delete)
        self.txt_log.bind('<Button-3>', self._on_log_right_click)

    def destroy(self):
        try:
            self.stop_speaking(join=True)
        except Exception:
            pass
        super().destroy()

if __name__ == '__main__':
    app = ChatTTSApp()
    app.mainloop()
