#!/usr/bin/env python3
"""Office do PDF. Masowa konwersja dokumentow Word, Excel i PowerPoint do PDF
przez zainstalowany Microsoft Office (wyglad PDF = wydruk z Office).

Uruchomienie: python office_pdf.py            (GUI)
              python office_pdf.py --selftest (test logiki)
"""
import os
import signal
import sys
import threading
import traceback
from collections import Counter

if getattr(sys, "frozen", False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

TYPES = {
    "Word": (".doc", ".docx", ".docm", ".rtf", ".odt"),
    "Excel": (".xls", ".xlsx", ".xlsm", ".ods"),
    "PowerPoint": (".ppt", ".pptx", ".pptm", ".odp"),
}
APP_OF = {ext: app for app, exts in TYPES.items() for ext in exts}
# haslo-atrapa: plik chroniony haslem konczy sie bledem zamiast okienka,
# ktore zawiesiloby niewidoczny Office
NO_PASSWORD = "\x01brak-hasla"
# Excel i PowerPoint przy blednym hasle (i innych okienkach) czekaja w nieskonczonosc;
# po tym czasie ich proces jest zabijany i seria idzie dalej
# ponytail: Word nie ma straznika (sam zglasza bledy), dodac gdy sie zawiesi
TIMEOUT_S = 180
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def encrypted_ooxml(path):
    """docx/xlsx/pptx z haslem to kontener OLE zamiast ZIP-a."""
    if not path.lower().endswith(("x", "m")):  # .doc/.xls/.ppt zawsze sa OLE
        return False
    with open(path, "rb") as f:
        return f.read(8) == OLE_MAGIC

# ------------------------------------------------------------- pliki ----


def find_inputs(inp, subfolders=False):
    """[(sciezka, sciezka_wzgledna_folderu)] - pomija pliki tymczasowe ~$."""
    if os.path.isfile(inp):
        return [(inp, "")]
    out = []
    for folder, dirs, files in os.walk(inp):
        dirs.sort()
        rel = os.path.relpath(folder, inp)
        for f in sorted(files):
            if os.path.splitext(f)[1].lower() in APP_OF and not f.startswith("~$"):
                out.append((os.path.join(folder, f), "" if rel == "." else rel))
        if not subfolders:
            break
    return out


def pdf_names(files, out_dir):
    """Sciezki PDF; 'umowa.docx' i 'umowa.xlsx' w jednym folderze
    daja 'umowa_docx.pdf' i 'umowa_xlsx.pdf' zamiast nadpisywac sie."""
    def key(src, rel):
        return rel, os.path.splitext(os.path.basename(src))[0].lower()
    counts = Counter(key(s, r) for s, r in files)
    out = []
    for src, rel in files:
        stem, ext = os.path.splitext(os.path.basename(src))
        if counts[key(src, rel)] > 1:
            stem += "_" + ext[1:].lower()
        out.append(os.path.join(out_dir, rel, stem + ".pdf"))
    return out


# ------------------------------------------------------------- Office ----


class Office:
    """Uruchamia Worda/Excela/PowerPointa przy pierwszym pliku danego typu
    i trzyma do konca serii. Po bledzie aplikacja jest zamykana i przy
    nastepnym pliku startuje od nowa (np. gdy Office sie wysypal)."""

    def __init__(self):
        self.apps = {}

    def _app(self, name):
        if name not in self.apps:
            import win32com.client
            try:
                app = win32com.client.DispatchEx(name + ".Application")
            except Exception:
                raise RuntimeError("Nie udalo sie uruchomic programu %s"
                                   " - czy jest zainstalowany?" % name)
            if name != "PowerPoint":  # PowerPoint nie pozwala ukryc okna aplikacji
                app.Visible = False
            app.DisplayAlerts = 0 if name != "Excel" else False
            self.apps[name] = app
        return self.apps[name]

    @staticmethod
    def _pid(app, name):
        import win32process
        hwnd = {"Excel": "Hwnd", "PowerPoint": "HWND"}.get(name)
        try:
            return win32process.GetWindowThreadProcessId(getattr(app, hwnd))[1] if hwnd else None
        except Exception:
            return None

    def convert(self, src, dst):
        name = APP_OF[os.path.splitext(src)[1].lower()]
        if encrypted_ooxml(src):
            raise RuntimeError("plik chroniony haslem - zdejmij haslo w Office i sprobuj ponownie")
        src, dst = os.path.abspath(src), os.path.abspath(dst)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        app = self._app(name)
        pid = self._pid(app, name)
        killed = []

        def kill():
            killed.append(1)
            os.kill(pid, signal.SIGTERM)  # na Windows = TerminateProcess

        watchdog = threading.Timer(TIMEOUT_S, kill) if pid else None
        if watchdog:
            watchdog.start()
        try:
            if name == "Word":
                d = app.Documents.Open(src, ConfirmConversions=False, ReadOnly=True,
                                       AddToRecentFiles=False, PasswordDocument=NO_PASSWORD,
                                       Visible=False)
                try:
                    d.SaveAs2(dst, FileFormat=17)  # wdFormatPDF
                finally:
                    d.Close(False)
            elif name == "Excel":
                wb = app.Workbooks.Open(src, UpdateLinks=0, ReadOnly=True,
                                        Password=NO_PASSWORD, AddToMru=False)
                try:
                    wb.ExportAsFixedFormat(0, dst)  # xlTypePDF, wszystkie arkusze
                finally:
                    wb.Close(False)
            else:
                p = app.Presentations.Open(src, ReadOnly=True, Untitled=False, WithWindow=False)
                try:
                    p.SaveAs(dst, 32)  # ppSaveAsPDF
                finally:
                    p.Close()
        except Exception:
            self._quit(name)
            if killed:
                raise RuntimeError("%s nie odpowiadal przez %d s (okienko z pytaniem albo"
                                   " haslo?) - plik pominiety" % (name, TIMEOUT_S))
            raise
        finally:
            if watchdog:
                watchdog.cancel()

    def _quit(self, name):
        try:
            self.apps.pop(name).Quit()
        except Exception:
            pass

    def close(self):
        for name in list(self.apps):
            self._quit(name)


def com_error(e):
    """Czytelny opis bledu COM (pywintypes.com_error ma opis gleboko w args)."""
    try:
        return str(e.args[2][2] or e.args[1]).strip()
    except Exception:
        return str(e)


# ------------------------------------------------------------------ wsad ----


def run(inp, out_dir, subfolders=False, log=print):
    files = find_inputs(inp, subfolders)
    if not files:
        log("Brak dokumentow Office w: %s" % inp)
        return 0
    import pythoncom  # COM w watku innym niz glowny (GUI)
    pythoncom.CoInitialize()
    office = Office()
    ok = 0
    try:
        for (src, rel), dst in zip(files, pdf_names(files, out_dir)):
            log("%s -> %s" % (os.path.join(rel, os.path.basename(src)), os.path.basename(dst)))
            try:
                office.convert(src, dst)
                ok += 1
            except RuntimeError as e:
                log("  BLAD: %s" % e)
            except Exception as e:
                log("  BLAD: %s (plik chroniony haslem albo uszkodzony?)" % com_error(e))
    finally:
        office.close()
    log("Zakonczono: %d z %d plikow -> %s" % (ok, len(files), out_dir))
    return ok


# -------------------------------------------------------------------- GUI ----


def gui():
    import tkinter as tk
    from tkinter import filedialog, ttk, scrolledtext

    root = tk.Tk()
    root.title("Office do PDF (Word, Excel, PowerPoint)")
    root.geometry("780x480")
    pad = dict(padx=6, pady=3)

    v_in = tk.StringVar(value=os.path.join(APP_DIR, "INPUT"))
    v_out = tk.StringVar(value=os.path.join(APP_DIR, "OUTPUT"))
    v_sub = tk.BooleanVar(value=False)
    all_exts = " ".join("*" + e for e in APP_OF)

    f = ttk.Frame(root)
    f.pack(fill="x", **pad)
    ttk.Label(f, text="Plik lub folder:").grid(row=0, column=0, sticky="w", **pad)
    ttk.Entry(f, textvariable=v_in, width=60).grid(row=0, column=1, **pad)
    ttk.Button(f, text="Plik...", command=lambda: v_in.set(filedialog.askopenfilename(
        filetypes=[("Dokumenty Office", all_exts)]) or v_in.get())).grid(row=0, column=2, **pad)
    ttk.Button(f, text="Folder...", command=lambda: v_in.set(
        filedialog.askdirectory() or v_in.get())).grid(row=0, column=3, **pad)
    ttk.Label(f, text="Folder wyjsciowy:").grid(row=1, column=0, sticky="w", **pad)
    ttk.Entry(f, textvariable=v_out, width=60).grid(row=1, column=1, **pad)
    ttk.Button(f, text="Wybierz...", command=lambda: v_out.set(
        filedialog.askdirectory() or v_out.get())).grid(row=1, column=2, **pad)
    ttk.Checkbutton(f, text="Razem z podfolderami (uklad folderow zostaje zachowany)",
                    variable=v_sub).grid(row=2, column=1, sticky="w", **pad)

    log_box = scrolledtext.ScrolledText(root, height=16)

    def log(msg):
        def put():
            log_box.insert("end", str(msg) + "\n")
            log_box.see("end")
        root.after(0, put)

    btn = ttk.Button(root, text="Konwertuj do PDF")
    btn.pack(pady=6)
    log_box.pack(fill="both", expand=True, **pad)

    def start():
        inp, out = v_in.get().strip('" '), v_out.get().strip('" ')
        log_box.delete("1.0", "end")
        if not os.path.exists(inp):
            return log("Wskaz istniejacy plik lub folder.")
        if not out:
            return log("Wskaz folder wyjsciowy.")
        btn.config(state="disabled")

        def work():
            try:
                run(inp, out, v_sub.get(), log)
            except Exception:
                log("BLAD:\n" + traceback.format_exc())
            finally:
                root.after(0, lambda: btn.config(state="normal"))

        threading.Thread(target=work, daemon=True).start()

    btn.config(command=start)
    if "--selftest" in sys.argv:
        root.after(200, root.destroy)
    import aktualizacja
    aktualizacja.start(root, "DawidBochno/Office-do-PDF", "main", "office_pdf.py")
    root.mainloop()


# --------------------------------------------------------------- selftest ----


def office_installed(name):
    import winreg
    try:
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name + ".Application"))
        return True
    except OSError:
        return False


def selftest():
    import shutil
    import tempfile

    tmp = tempfile.mkdtemp()
    src = os.path.join(tmp, "in")
    for rel in ("umowa.docx", "umowa.xlsx", "prezentacja.pptx", "notatka.txt",
                "~$umowa.docx", os.path.join("2024", "pismo.doc")):
        os.makedirs(os.path.dirname(os.path.join(src, rel)), exist_ok=True)
        open(os.path.join(src, rel), "wb").close()
    files = find_inputs(src)
    assert [os.path.basename(f) for f, _ in files] == \
        ["prezentacja.pptx", "umowa.docx", "umowa.xlsx"], files
    files = find_inputs(src, subfolders=True)
    assert len(files) == 4 and files[-1][1] == "2024", files
    names = [os.path.relpath(p, "OUT") for p in pdf_names(files, "OUT")]
    assert names == ["prezentacja.pdf", "umowa_docx.pdf", "umowa_xlsx.pdf",
                     os.path.join("2024", "pismo.pdf")], names
    assert find_inputs(os.path.join(src, "umowa.docx")) == [(os.path.join(src, "umowa.docx"), "")]
    with open(os.path.join(src, "umowa.xlsx"), "wb") as fh:
        fh.write(OLE_MAGIC + b"...")
    assert encrypted_ooxml(os.path.join(src, "umowa.xlsx"))
    assert not encrypted_ooxml(os.path.join(src, "umowa.docx"))
    try:
        Office().convert(os.path.join(src, "umowa.xlsx"), os.path.join(tmp, "x.pdf"))
        raise AssertionError("plik z haslem nie zostal odrzucony")
    except RuntimeError as e:
        assert "haslem" in str(e)

    # prawdziwa konwersja - tylko tam, gdzie jest Office (nie ma go w CI)
    import win32com.client
    real = os.path.join(tmp, "real")
    os.makedirs(real)
    made = []
    if office_installed("Word"):
        w = win32com.client.DispatchEx("Word.Application")
        try:
            d = w.Documents.Add()
            d.Content.Text = "Zażółć gęślą jaźń - test"
            d.SaveAs2(os.path.join(real, "test.docx"), FileFormat=16)
            d.Close(False)
        finally:
            w.Quit()
        made.append("test.docx")
    if office_installed("Excel"):
        x = win32com.client.DispatchEx("Excel.Application")
        x.DisplayAlerts = False
        try:
            wb = x.Workbooks.Add()
            wb.Worksheets(1).Range("A1").Value = "test"
            wb.SaveAs(os.path.join(real, "test.xlsx"), FileFormat=51)
            # stary .xls z haslem: Excel czeka na haslo -> musi zadzialac straznik
            wb.SaveAs(os.path.join(real, "haslo.xls"), FileFormat=56, Password="abc")
            wb.Close(False)
        finally:
            x.Quit()
        made.append("test.xlsx")
    if made:
        with open(os.path.join(real, "zepsuty.docx"), "wb") as fh:
            fh.write(b"to nie jest dokument Word")
        log = []
        global TIMEOUT_S
        TIMEOUT_S, saved = 10, TIMEOUT_S
        try:
            ok = run(real, os.path.join(tmp, "out"), log=log.append)
        finally:
            TIMEOUT_S = saved
        if "test.xlsx" in made:
            assert any("nie odpowiadal" in x for x in log), log
        out = sorted(os.listdir(os.path.join(tmp, "out")))
        assert ok == len(made), log
        assert len(out) == len(made) and "zepsuty.pdf" not in out, out
        for p in out:
            with open(os.path.join(tmp, "out", p), "rb") as fh:
                assert fh.read(5) == b"%PDF-", p
        assert any("zepsuty.docx" in x for x in log) and any("BLAD" in x for x in log), log
        print("konwersja Office OK:", ", ".join(made))
    else:
        print("brak Microsoft Office - pomijam test prawdziwej konwersji")
    shutil.rmtree(tmp, ignore_errors=True)

    import aktualizacja
    aktualizacja.selftest()
    print("selftest OK")


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        selftest()
        if "--gui" in sys.argv:
            gui()
    else:
        gui()
