from tkinter import filedialog, messagebox
import tkinter as tk
import threading
import openpyxl
import pathlib
import ctypes
import main


class Window:
    def __init__(self):
        threading.Thread.__init__(self)
        self.path = pathlib.Path(__file__).parent.resolve()
        self.label1 = tk.Label(text="Qlik MOS:", width = 8, anchor="e")
        self.label2 = tk.Label(text="SMSS:", width = 8, anchor="e")
        self.label3 = tk.Label(text="Salvar em:", width = 8, anchor="e")
        self.label4 = tk.Label(
            text="",
            font=("Arial", 7),
            justify="center"
        )

        self.label1.grid(column=0, row=0, padx=10, pady=(5, 5))
        self.label2.grid(column=0, row=1, padx=10, pady=(0, 5))
        self.label3.grid(column=0, row=2, padx=10, pady=(0, 5))
        self.label4.grid(column=0, columnspan=4, row=4, padx=10, pady=(5, 5))

        self.entry1 = tk.Entry(width=50)
        self.entry2 = tk.Entry(width=50)
        self.entry3 = tk.Entry(width=50)

        self.entry1.grid(column=1, columnspan=2, row=0, pady=(5, 5))
        self.entry2.grid(column=1, columnspan=2, row=1, pady=(0, 5))
        self.entry3.grid(column=1, columnspan=2, row=2, pady=(0, 5))

        self.button1 = tk.Button(
            text="Escolher", command=lambda: self.open_file(self.entry1), width = 8
        )
        self.button2 = tk.Button(
            text="Escolher", command=lambda: self.open_file(self.entry2), width = 8
        )
        self.button3 = tk.Button(
            text="Escolher", command=lambda: self.save_location(self.entry3), width = 8
        )
        self.button4 = tk.Button(
            text="Exportar", command=lambda: self.threading(self.Exportar)
        )

        self.button1.grid(column=3, row=0, padx=(10, 0), pady=2)
        self.button2.grid(column=3, row=1, padx=(10, 0), pady=2)
        self.button3.grid(column=3, row=2, padx=(10, 0), pady=2)
        self.button4.grid(column=1, columnspan=2, row=3, pady=2, sticky=("EW"))

    def threading(self, target):
        th = threading.Thread(target=target, daemon = True)
        th.start()

    def open_file(self, entry):
        filename = filedialog.askopenfilename(
            initialdir="Downloads",
            title="Select a File",
            filetypes=(("Excel files", "*.xlsx*"), ("all files", "*.*")),
        )
        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def save_location(self, entry):
        filename = filedialog.askdirectory(initialdir="Downloads")
        entry.delete(0, tk.END)
        entry.insert(0, filename)

    def Exportar(self):
        if self.entry1.get() == "" or self.entry2.get() == "" or self.entry3.get() == "":
            messagebox.showinfo(title="Faltando Informações", message="Um ou mais campos não foram preenchidos, por favor solecione um caminho para todos os campos.")
            return
        try:
            self.button4.config(state="disabled")
            self.label4.config(text = "Iniciando")
            main.compare(self.label4, self.entry1.get(), self.entry2.get(), self.entry3.get())
            self.button4.config(state="normal")
        except openpyxl.utils.exceptions.InvalidFileException as e:
            messagebox.showwarning(title = "Arquivo incorreto", message= "Um ou mais arquivos não foram selecionados ou são invalidos")
            self.label4.config(text = "")
        except Exception as e:
            messagebox.showerror(title = "ERROR", message = (f"Favor repassar este erro para o desenvolvedor:\n\n{e}"))
            self.root.destroy()

    root = tk.Tk()
    root.minsize(470, 150)
    root.maxsize(470, 150)
    root.title("Comparar Baixa de MOS com SMSS")


class thread_with_exception(threading.Thread):
    def __init__(self, name):
        threading.Thread.__init__(self)
        self.name = name

    def get_id(self):
        if hasattr(self, "_thread_id"):
            return self._thread_id
        for id, thread in threading._active.items():
            if thread is self:
                return id

    def raise_exception(self):
        thread_id = self.get_id()
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(
            thread_id, ctypes.py_object(SystemExit)
        )
        if res > 1:
            ctypes.pythonapi.PyThreadState_SetAsyncExc(thread_id, 0)

window = Window()
window.root.mainloop()
