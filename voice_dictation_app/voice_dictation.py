import threading
import tkinter as tk
from tkinter import filedialog, messagebox

import speech_recognition as sr


class VoiceDictationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Голосовая диктовка")
        self.root.geometry("700x520")
        self.recognizer = sr.Recognizer()
        self.is_listening = False

        self.text = tk.Text(root, wrap=tk.WORD, font=("Arial", 12))
        self.text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        button_frame = tk.Frame(root)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        self.record_button = tk.Button(button_frame, text="Диктовать", command=self.start_recognition)
        self.record_button.pack(side=tk.LEFT, padx=5)

        self.copy_button = tk.Button(button_frame, text="Копировать", command=self.copy_text)
        self.copy_button.pack(side=tk.LEFT, padx=5)

        self.save_button = tk.Button(button_frame, text="Сохранить", command=self.save_text)
        self.save_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = tk.Button(button_frame, text="Очистить", command=self.clear_text)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        self.status_label = tk.Label(root, text="Готово к диктовке", anchor="w")
        self.status_label.pack(fill=tk.X, padx=10, pady=(0, 10))

    def set_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def start_recognition(self):
        if self.is_listening:
            return
        self.is_listening = True
        self.record_button.config(state=tk.DISABLED)
        self.set_status("Слушаю... Говорите, пожалуйста")

        thread = threading.Thread(target=self.recognize_phrase)
        thread.daemon = True
        thread.start()

    def recognize_phrase(self):
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
                audio = self.recognizer.listen(source, timeout=7, phrase_time_limit=8)

            self.set_status("Обработка...")
            text = self.recognizer.recognize_google(audio, language="ru-RU")
            self.text.insert(tk.END, text + "\n")
            self.set_status("Готово. Текст добавлен")
        except sr.WaitTimeoutError:
            self.set_status("Не услышал. Попробуй ещё раз")
            messagebox.showwarning("Таймаут", "Не было речи за отведённое время.")
        except sr.UnknownValueError:
            self.set_status("Не удалось распознать речь")
            messagebox.showwarning("Ошибка распознавания", "Я не смог разобрать речь.")
        except sr.RequestError as e:
            self.set_status("Ошибка сервиса распознавания")
            messagebox.showerror("Ошибка сервиса", f"Не удалось обратиться к сервису распознавания речи:\n{e}")
        except OSError as e:
            self.set_status("Ошибка с микрофоном")
            messagebox.showerror("Микрофон", f"Не удалось открыть микрофон:\n{e}")
        finally:
            self.record_button.config(state=tk.NORMAL)
            self.is_listening = False

    def copy_text(self):
        content = self.text.get("1.0", tk.END).strip()
        if not content:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.set_status("Текст скопирован в буфер обмена")

    def save_text(self):
        content = self.text.get("1.0", tk.END).strip()
        if not content:
            messagebox.showinfo("Сохранение", "Текст отсутствует для сохранения.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*")],
            title="Сохранить текст"
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self.set_status(f"Сохранено: {path}")

    def clear_text(self):
        self.text.delete("1.0", tk.END)
        self.set_status("Текст очищен")


if __name__ == "__main__":
    root = tk.Tk()
    app = VoiceDictationApp(root)
    root.mainloop()
