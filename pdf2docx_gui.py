import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pdf2docx import Converter


def build_app():
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()

    root.title("PDF → DOCX Converter")
    root.geometry("560x260")
    root.minsize(500, 240)

    pdf_path = {"value": None}
    out_path = {"value": None}

    def pick_pdf():
        f = filedialog.askopenfilename(
            title="Choose a PDF file",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if f:
            set_input(f)

    def set_input(f):
        pdf_path["value"] = f
        lbl_pdf.config(text=os.path.basename(f))
        lbl_pdf_hint.config(text=f)
        default_out = os.path.splitext(f)[0] + ".docx"
        out_path["value"] = default_out
        var_out.set(default_out)
        btn_convert.config(state=tk.NORMAL)

    def pick_out():
        f = filedialog.asksaveasfilename(
            title="Save DOCX as",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")],
            initialfile=os.path.basename(var_out.get() or "output.docx"),
        )
        if f:
            out_path["value"] = f
            var_out.set(f)

    def on_drop(event):
        files = root.tk.splitlist(event.data)
        pdf = next((p for p in files if p.lower().endswith(".pdf")), None)
        if pdf:
            set_input(pdf)
        return event.action

    def convert():
        if not pdf_path["value"]:
            return
        out = var_out.get().strip() or pdf_path["value"]
        if not out.lower().endswith(".docx"):
            out += ".docx"
        out_path["value"] = out
        btn_convert.config(state=tk.DISABLED)
        bar.start()
        status.set("Converting...")
        threading.Thread(target=do_convert, args=(pdf_path["value"], out), daemon=True).start()

    def do_convert(src, dst):
        try:
            cv = Converter(src)
            cv.convert(dst, start=0, end=None)
            cv.close()
            status.set("Done.")
        except Exception as exc:
            status.set("Failed.")
            root.after(0, lambda: messagebox.showerror("Error", str(exc)))
        finally:
            bar.stop()
            btn_convert.config(state=tk.NORMAL)

    pad = {"padx": 12, "pady": 8}

    top = tk.Frame(root)
    top.pack(fill=tk.X, pady=8)
    tk.Label(top, text="PDF → DOCX Converter", font=("Segoe UI", 13, "bold")).pack()

    frame = tk.Frame(root)
    frame.pack(fill=tk.X, padx=16)

    tk.Label(frame, text="Input PDF:", anchor="w").grid(row=0, column=0, sticky="w")
    lbl_pdf = tk.Label(frame, text="— none —", fg="#444")
    lbl_pdf.grid(row=0, column=1, sticky="w")
    tk.Button(frame, text="Browse…", command=pick_pdf).grid(row=0, column=2, padx=6)

    lbl_pdf_hint = tk.Label(frame, text="or drag & drop a .pdf here", fg="#888", anchor="w")
    lbl_pdf_hint.grid(row=1, column=1, columnspan=2, sticky="w")

    tk.Label(frame, text="Output:", anchor="w").grid(row=2, column=0, sticky="w", pady=(14, 0))
    var_out = tk.StringVar()
    tk.Entry(frame, textvariable=var_out).grid(row=2, column=1, sticky="we", pady=(14, 0))
    tk.Button(frame, text="Save as…", command=pick_out).grid(row=2, column=2, padx=6, pady=(14, 0))

    frame.columnconfigure(1, weight=1)

    btn_frame = tk.Frame(root)
    btn_frame.pack(fill=tk.X, padx=16, pady=16)
    status = tk.StringVar()
    status.set("Idle.")
    tk.Label(btn_frame, textvariable=status, fg="#555").pack(side=tk.LEFT)
    btn_convert = tk.Button(btn_frame, text="Convert", command=convert, state=tk.DISABLED, padx=24)
    btn_convert.pack(side=tk.RIGHT)

    bar = ttk.Progressbar(root, mode="indeterminate")
    bar.pack(fill=tk.X, padx=16)

    root.drop_target_register("DND_Files")
    root.dnd_bind("<<Drop>>", on_drop)

    return root


def main():
    root = build_app()
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".pdf"):
        pdf = os.path.abspath(sys.argv[1])
        root.after(100, lambda: set_input(pdf))
    root.mainloop()


if __name__ == "__main__":
    main()