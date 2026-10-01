import os
import sys
import threading

from pdf2docx import Converter

import polish as polish_module
import snip_regions

SNIP_DPI = 300


def _vocab_beside(src):
    """Source files next to the PDF that name the words it uses."""
    stem = os.path.splitext(src)[0]
    return [p for p in (stem + '.tex', stem + '.bbl') if os.path.exists(p)]


def _queue_error_dialog(root, messagebox, error):
    message = str(error)
    root.after(0, lambda: messagebox.showerror("Error", message))


def _polish_summary(counts):
    return (f" Restored {counts['spaces']} spaces,"
            f" removed {counts['page breaks dropped']} page breaks.")


def build_app():
    # Imported here, not at module level: --cli runs on installs without Tk.
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except Exception:
        root = tk.Tk()

    root.title("PDF → DOCX Converter")
    root.geometry("560x330")
    root.minsize(500, 310)

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

            note = "Done."
            if var_snip.get():
                status.set("Cropping tables and equations…")
                replaced, missed = snip_regions.snip(
                    src, dst, dpi=SNIP_DPI,
                    tables=True, equations=True)
                note = f"Done. {replaced} tables and equations kept as images"
                if missed:
                    note += f", {missed} left as converted text"
                note += "."
            if var_polish.get():
                status.set("Tidying the document…")
                counts = polish_module.polish(dst, src, _vocab_beside(src))
                note += _polish_summary(counts)
            status.set(note)
        except Exception as exc:
            status.set("Failed.")
            _queue_error_dialog(root, messagebox, exc)
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

    var_snip = tk.BooleanVar(value=True)
    tk.Checkbutton(
        frame,
        text="Tables and equations as images from the PDF",
        variable=var_snip, anchor="w",
    ).grid(row=3, column=1, columnspan=2, sticky="w", pady=(10, 0))
    tk.Label(
        frame,
        text="Keeps their layout exactly; body text stays editable.",
        fg="#888", anchor="w",
    ).grid(row=4, column=1, columnspan=2, sticky="w")

    var_polish = tk.BooleanVar(value=True)
    tk.Checkbutton(
        frame,
        text="Tidy the document for reading and editing",
        variable=var_polish, anchor="w",
    ).grid(row=5, column=1, columnspan=2, sticky="w", pady=(8, 0))
    tk.Label(
        frame,
        text="Substitutes missing fonts, lets the text flow, restores lost spaces.",
        fg="#888", anchor="w",
    ).grid(row=6, column=1, columnspan=2, sticky="w")

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

    root.set_input = set_input      # so a PDF named on the command line loads
    return root


def convert_cli(src, dst, snip=True, tidy=True):
    """Convert without opening the window, for scripting and for testing."""
    cv = Converter(src)
    cv.convert(dst, start=0, end=None)
    cv.close()
    if snip:
        replaced, missed = snip_regions.snip(src, dst, dpi=SNIP_DPI)
        print(f"{replaced} tables and equations kept as images"
              + (f", {missed} left as converted text" if missed else ""))
    if tidy:
        counts = polish_module.polish(dst, src, _vocab_beside(src))
        print(', '.join(f'{name}: {value}' for name, value in counts.items()))
    print(dst)


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    flags = {a for a in sys.argv[1:] if a.startswith("-")}

    if "--cli" in flags:
        if not args:
            print("usage: pdf2docx_gui.py --cli input.pdf [output.docx] "
                  "[--no-images] [--no-tidy]")
            return 2
        src = os.path.abspath(args[0])
        dst = os.path.abspath(args[1]) if len(args) > 1 \
            else os.path.splitext(src)[0] + ".docx"
        convert_cli(src, dst,
                    snip="--no-images" not in flags,
                    tidy="--no-tidy" not in flags)
        return 0

    root = build_app()
    if args and args[0].lower().endswith(".pdf"):
        pdf = os.path.abspath(args[0])
        root.after(100, lambda: root.set_input(pdf))
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
