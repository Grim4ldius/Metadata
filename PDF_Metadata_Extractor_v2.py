import os
import subprocess
import platform
import sys
from tkinter import Tk, Label, filedialog, messagebox, Frame, Scrollbar, VERTICAL, RIGHT, Y, LEFT, StringVar, Entry, Button
from tkinter.ttk import Treeview, Separator
from datetime import datetime
import tkinter.font as tkFont
from tkinter import Menu
import pandas as pd
from pdf_extractor import extract_metadata  # Utilise le module séparé

APP_VERSION = "1.3"  # Incrémente ce numéro à chaque évolution majeure ou mineure

# Variable globale pour filtrage
all_metadata = []

# ----------- TOOLTIP CLASS -----------
class ToolTip(object):
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)
    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert") if self.widget.winfo_ismapped() else (0, 0, 0, 0)
        x = x + self.widget.winfo_rootx() + 25
        y = y + self.widget.winfo_rooty() + 20
        self.tipwindow = tw = Tk()
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = Label(tw, text=self.text, background="#ffffe0", relief="solid", borderwidth=1, font=("Arial", 10))
        label.pack()
    def hide_tip(self, event=None):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def show_about():
    import tkinter
    import pypdf
    import pandas
    info = (
        "PDF Metadata Extractor\n"
        "Auteur : Xavier Minotte\n"
        f"Version : {APP_VERSION}\n\n"
        f"Python : {platform.python_version()}\n"
        f"Tkinter : {tkinter.TkVersion}\n"
        f"pypdf : {pypdf.__version__}\n"
        f"pandas : {pandas.__version__}\n"
    )
    messagebox.showinfo("À propos", info)

def convert_pdf_date(pdf_date):
    if not pdf_date:
        return ""
    try:
        if pdf_date.startswith("D:"):
            pdf_date = pdf_date[2:]
        if "+" in pdf_date or "-" in pdf_date:
            sign = "+" if "+" in pdf_date else "-"
            main, tz = pdf_date.split(sign)
            tz = tz.replace("'", "")
            pdf_date = f"{main}{sign}{tz}"
            dt = datetime.strptime(pdf_date, "%Y%m%d%H%M%S%z")
        else:
            dt = datetime.strptime(pdf_date[:14], "%Y%m%d%H%M%S")
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return pdf_date

def open_selected_folder():
    folder = folder_label.cget("text")
    if folder and os.path.isdir(folder):
        if sys.platform.startswith("win"):
            os.startfile(folder)
        elif sys.platform.startswith("darwin"):
            subprocess.call(["open", folder])
        else:
            subprocess.call(["xdg-open", folder])

def filter_metadata(*args):
    global all_metadata
    search = search_var.get().lower()
    if not search or search == "recherche...":
        display_metadata(all_metadata)
        return
    filtered = [
        m for m in all_metadata
        if any(search in str(m.get(col, "")).lower() for col in ["File Name", "Title", "Author", "Subject"])
    ]
    display_metadata(filtered)

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_label.config(text=folder_path)
        global all_metadata
        all_metadata = extract_metadata(folder_path)
        display_metadata(all_metadata)
        if not all_metadata:
            messagebox.showinfo("Aucun PDF", "Aucun fichier PDF trouvé dans ce dossier.")

def auto_resize_columns(tree):
    style = tree.cget("style") or "Treeview"
    font_name = tree.tk.call("ttk::style", "lookup", style, "font")
    font = tkFont.nametofont(font_name) if font_name else tkFont.Font()
    for col in tree["columns"]:
        if col == "Full Path":
            tree.column(col, width=0, stretch=False, minwidth=0)
        else:
            tree.heading(col, text=col)
            tree.column(col, width=100, stretch=True, minwidth=50)

# ----------- INTERFACE GRAPHIQUE -----------
root = Tk()
root.title("PDF Metadata Extractor - Sélection et Export")
root.geometry("1100x600")

# Police par défaut plus grande
default_font = tkFont.nametofont("TkDefaultFont")
default_font.configure(size=11)
root.option_add("*Font", default_font)

folder_label = Label(root, text="")
folder_label.pack(pady=(2, 0))

# ----------- BARRE HAUTE : recherche, résultats, boutons -----------
top_bar_frame = Frame(root)
top_bar_frame.pack(fill="x", padx=10, pady=(8, 2))

search_var = StringVar()
def clear_placeholder(event):
    if search_entry.get() == "Recherche...":
        search_entry.delete(0, "end")
def add_placeholder(event):
    if not search_entry.get():
        search_entry.insert(0, "Recherche...")
search_var.trace("w", filter_metadata)
search_entry = Entry(top_bar_frame, textvariable=search_var, bg="#f0f4f8", width=32)
search_entry.pack(side=LEFT, padx=(0, 10))
search_entry.insert(0, "Recherche...")
search_entry.bind("<FocusIn>", clear_placeholder)
search_entry.bind("<FocusOut>", add_placeholder)

results_label = Label(top_bar_frame, text="Résultats", font=("Arial", 13, "bold"))
results_label.pack(side=LEFT, padx=(0, 16), anchor="w")

counter_var = StringVar()
counter_label = Label(top_bar_frame, textvariable=counter_var, font=("Arial", 10, "italic"))
counter_label.pack(side=LEFT, padx=(0, 16))

select_btn = Button(top_bar_frame, text="Sélectionner un dossier", command=select_folder, width=20, bg="#4a90e2", fg="white")
select_btn.pack(side=LEFT, padx=4)

def export_menu_popup(event=None):
    exportmenu.tk_popup(select_btn.winfo_rootx() + select_btn.winfo_width(), select_btn.winfo_rooty())

export_btn = Button(top_bar_frame, text="Exporter", width=14, bg="#4a90e2", fg="white", command=export_menu_popup)
export_btn.pack(side=LEFT, padx=4)

ToolTip(select_btn, "Choisir un dossier à analyser")
ToolTip(export_btn, "Exporter la sélection (Excel ou CSV)")

sep = Separator(root, orient="horizontal")
sep.pack(fill="x", padx=10, pady=6)

context_menu = Menu(root, tearoff=0)
context_menu.add_command(label="Ouvrir le PDF", command=lambda: None)
context_menu.add_command(label="Ouvrir le dossier", command=lambda: None)
context_menu.add_separator()
context_menu.add_command(label="Copier", command=lambda: None)

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Sélectionner un dossier", command=select_folder)
filemenu.add_command(label="Ouvrir dossier dans l'explorateur", command=open_selected_folder)
filemenu.add_separator()
filemenu.add_command(label="Quitter", command=root.quit)
menubar.add_cascade(label="Fichier", menu=filemenu)

exportmenu = Menu(menubar, tearoff=0)
exportmenu.add_command(label="Exporter vers Excel", command=lambda: export_to_excel())
exportmenu.add_command(label="Exporter vers CSV", command=lambda: export_to_csv())
menubar.add_cascade(label="Export", menu=exportmenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="À propos", command=show_about)
menubar.add_cascade(label="Aide", menu=helpmenu)
root.config(menu=menubar)

# ----------- TREEVIEW -----------
frame = Frame(root)
frame.pack(fill="both", expand=True, pady=(5, 0))

tree = Treeview(
    frame,
    columns=(
        "File Name", "Title", "Author", "Subject", "Creator", "Producer", "Creation Date", "Modification Date", "Full Path"
    ),
    show="headings",
    selectmode="browse"
)
for col in tree["columns"]:
    if col == "Full Path":
        tree.column(col, width=0, stretch=False, minwidth=0)  # Masque la colonne Full Path
        # Ne PAS créer d'en-tête pour "Full Path"
    else:
        tree.heading(col, text=col)
        tree.column(col, width=100, stretch=True, minwidth=50)

tree.pack(side=LEFT, fill="both", expand=True)
tree.tag_configure("evenrow", background="#f6f6f6")
tree.tag_configure("oddrow", background="#ffffff")

scrollbar = Scrollbar(frame, orient=VERTICAL, command=tree.yview)
scrollbar.pack(side=RIGHT, fill=Y)
tree.configure(yscrollcommand=scrollbar.set)

def display_metadata(metadata_list):
    for row in tree.get_children():
        tree.delete(row)
    if not metadata_list:
        counter_var.set("0 fichier affiché / 0 au total")
        return
    for idx, metadata in enumerate(metadata_list):
        tag = "evenrow" if idx % 2 == 0 else "oddrow"
        tree.insert(
            "", "end",
            values=(
                metadata.get("File Name", ""),
                metadata.get("Title", ""),
                metadata.get("Author", ""),
                metadata.get("Subject", ""),
                metadata.get("Creator", ""),
                metadata.get("Producer", ""),
                metadata.get("Creation Date", ""),
                metadata.get("Modification Date", ""),
                metadata.get("Full Path", "")  # Full Path stocké mais non affiché
            ),
            tags=(tag,)
        )
    auto_resize_columns(tree)
    counter_var.set(f"{len(metadata_list)} fichiers affichés / {len(all_metadata)} au total")

def export_to_excel():
    folder_path = folder_label.cget("text")
    if not folder_path:
        messagebox.showerror("Erreur", "Aucun dossier sélectionné.")
        return
    metadata_list = []
    for row in tree.get_children():
        values = tree.item(row)["values"]
        metadata_list.append(values)  # Inclut Full Path
    if not metadata_list:
        messagebox.showinfo("Aucun résultat", "Aucune donnée à exporter.")
        return
    df = pd.DataFrame(metadata_list, columns=[
        "File Name", "Title", "Author", "Subject", "Creator", "Producer", "Creation Date", "Modification Date", "Full Path"
    ])
    default_filename = f"PDF_Metadata_{os.path.basename(folder_path)}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        initialfile=default_filename,
        filetypes=[("Excel files", "*.xlsx")]
    )
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Succès", f"Metadata exported to {file_path}")

def export_to_csv():
    folder_path = folder_label.cget("text")
    if not folder_path:
        messagebox.showerror("Erreur", "Aucun dossier sélectionné.")
        return
    metadata_list = []
    for row in tree.get_children():
        values = tree.item(row)["values"]
        metadata_list.append(values)  # Inclut Full Path
    if not metadata_list:
        messagebox.showinfo("Aucun résultat", "Aucune donnée à exporter.")
        return
    df = pd.DataFrame(metadata_list, columns=[
        "File Name", "Title", "Author", "Subject", "Creator", "Producer", "Creation Date", "Modification Date", "Full Path"
    ])
    default_filename = f"PDF_Metadata_{os.path.basename(folder_path)}_{datetime.now().strftime('%Y%m%d')}.csv"
    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        initialfile=default_filename,
        filetypes=[("CSV files", "*.csv")]
    )
    if file_path:
        df.to_csv(file_path, index=False)
        messagebox.showinfo("Succès", f"Metadata exported to {file_path}")

def on_tree_double_click(event):
    row = tree.identify_row(event.y)
    if row:
        values = tree.item(row)["values"]
        file_path = values[-1]  # Full Path
        if file_path:
            try:
                if sys.platform.startswith("win"):
                    os.startfile(file_path)
                elif sys.platform.startswith("darwin"):
                    subprocess.call(["open", file_path])
                else:
                    subprocess.call(["xdg-open", file_path])
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier : {e}")
        else:
            messagebox.showinfo("Info", "Chemin du fichier non disponible.")
tree.bind("<Double-1>", on_tree_double_click)

def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    try:
        l.sort(key=lambda t: t[0].lower(), reverse=reverse)
    except Exception:
        l.sort(reverse=reverse)
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

root.mainloop()