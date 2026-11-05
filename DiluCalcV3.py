import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import math
from datetime import datetime
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageDraw, ImageFont
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_FILE = os.path.join(BASE_DIR, "dilucalc_library.json")

default_data = {
    "elements": {
        "Lead (Pb)": 207.2, "Copper (Cu)": 63.546, "Iron (Fe)": 55.845,
        "Cadmium (Cd)": 112.41, "Zinc (Zn)": 65.38, "Gold (Au)": 196.97,
        "Indium (In)": 114.82, "Bismuth (Bi)": 208.98
    },
    "acids": {
        "Nitric Acid (HNO3)": {"purity": 65.0, "density": 1.42},
        "Hidrochloric Acid (HCl)": {"purity": 37.0, "density": 1.19},
        "Sulfuric Acid (H2SO4)": {"purity": 98.0, "density": 1.84}
    },
    "solutes": ["Distilled Water", "Ethanol (%99.9)", "Methanol"]
}

def load_data():
    if not os.path.exists(DATA_FILE):
        save_data(default_data)
        return default_data
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

db = load_data()

def sort_treeview(tree, col, reverse):
    l = [(tree.set(k, col), k) for k in tree.get_children('')]
    try:
        l.sort(key=lambda t: float(t[0]), reverse=reverse)
    except ValueError:
        l.sort(key=lambda t: t[0].lower(), reverse=reverse)
    for index, (val, k) in enumerate(l):
        tree.move(k, '', index)

    tree.heading(col, command=lambda: sort_treeview(tree, col, not reverse))

# APP WINDOW

root = tk.Tk()
root.title("DiluCalc v3 - Solution, Alloy & Dilution Calculator")
root.geometry("1368x900")
root.minsize(1150, 760)
root.configure(bg="#f4f6f9")

style = ttk.Style()
if "clam" in style.theme_names():
    style.theme_use("clam")
style.configure("TFrame", background="#f4f6f9")
style.configure("TLabelframe", background="#ffffff", font=("Segoe UI", 10, "bold"))
style.configure("TLabelframe.Label", background="#ffffff", foreground="#1e293b", font=("Segoe UI", 11, "bold"))
style.configure("TLabel", background="#ffffff", font=("Segoe UI", 10))
style.configure("TCheckbutton", background="#ffffff", font=("Segoe UI", 10))
style.configure("TRadiobutton", background="#ffffff", font=("Segoe UI", 10))
style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), background="#2563eb", foreground="white")
style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[14, 6])

root.tk.eval("""
bind ComboboxListbox <KeyPress> {
    set key [string tolower %A]
    if {[string match {[a-z]} $key]} {
        set lb %W
        set list [$lb get 0 end]
        set curr [$lb curselection]
        if {$curr eq ""} { set curr -1 }
        set found -1
        for {set i [expr {$curr + 1}]} {$i < [llength $list]} {incr i} {
            if {[string match "${key}*" [string tolower [lindex $list $i]]]} {
                set found $i
                break
            }
        }
        if {$found == -1} {
            for {set i 0} {$i <= $curr} {incr i} {
                if {[string match "${key}*" [string tolower [lindex $list $i]]]} {
                    set found $i
                    break
                }
            }
        }
        if {$found != -1} {
            $lb selection clear 0 end
            $lb selection set $found
            $lb activate $found
            $lb see $found
        }
        break
    }
}
""")
# LIBRARY WINDOW
def open_library_window():
    lib_win = tk.Toplevel(root)
    lib_win.title("📚 Library Management (Element / Acid / Solute)")
    lib_win.geometry("620x450")
    lib_win.configure(bg="#f4f6f9")
    lib_win.grab_set()

    lbl_info = tk.Label(lib_win, text="You can add new elements, solutes or acids or change the current list.", bg="#f4f6f9", font=("Segoe UI", 10))
    lbl_info.pack(pady=8)

    notebook = ttk.Notebook(lib_win)
    notebook.pack(fill="both", expand=True, padx=15, pady=5)

    sub_tab1 = ttk.Frame(notebook)
    sub_tab2 = ttk.Frame(notebook)
    sub_tab3 = ttk.Frame(notebook)

    notebook.add(sub_tab1, text="🔹 Elements")
    notebook.add(sub_tab2, text="🧪 Acids")
    notebook.add(sub_tab3, text="💧 Solutes")

    # ELEMENT MANAGEMENT
    tk.Label(sub_tab1, text="Current Elements and their Atomic Masses", font=("Segoe UI", 10, "bold"), bg="#f4f6f9").pack(pady=5)
    tree_elem = ttk.Treeview(sub_tab1, columns=("Element", "Mass"), show="headings", height=10)
    tree_elem.heading("Element", text="Name and Symbol of the Element", command=lambda: sort_treeview(tree_elem, "Element", False))
    tree_elem.heading("Mass", text="Atomic Mass (g/mol)", command=lambda: sort_treeview(tree_elem, "Mass", False))
    tree_elem.pack(fill="both", expand=True, padx=10, pady=5)

    def refresh_elem_tree():
        for row in tree_elem.get_children():
            tree_elem.delete(row)
        for k, v in db["elements"].items():
            tree_elem.insert("", "end", values=(k, v))
    refresh_elem_tree()

    def popup_add_element():
        pop = tk.Toplevel(lib_win)
        pop.title("➕ Add New Element")
        pop.geometry("360x200")
        pop.grab_set()

        tk.Label(pop, text="Name and Symbol of the Element (Ex: Gold (Au)):").pack(pady=(15, 2))
        ent_name = ttk.Entry(pop, width=35)
        ent_name.pack()

        tk.Label(pop, text="Atomic Mass (g/mol):").pack(pady=(10, 2))
        ent_mass = ttk.Entry(pop, width=35)
        ent_mass.insert(0, "196.97")
        ent_mass.pack()

        def save_elem():
            new_element_name = ent_name.get().strip()
            try:
                new_element_atkes = float(ent_mass.get())
                if new_element_name:
                    db["elements"][new_element_name] = new_element_atkes
                    save_data(db)
                    refresh_elem_tree()
                    update_main_dropdowns()
                    messagebox.showinfo("Success", f"'{new_element_name}' Successfully added to the library!", parent=pop)
                    pop.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid number for Atomic Mass.", parent=pop)

        ttk.Button(pop, text="Add Element", command=save_elem).pack(pady=15)

    def remove_selected_element():
        selected = tree_elem.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an element from the list to remove.", parent=lib_win)
            return
        if len(db["elements"]) <= 1:
            messagebox.showwarning("Warning", "At least one element must remain in the library!", parent=lib_win)
            return
        elem_name = tree_elem.item(selected[0])["values"][0]
        if messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove '{elem_name}'?", parent=lib_win):
            del db["elements"][elem_name]
            save_data(db)
            refresh_elem_tree()
            update_main_dropdowns()

    btn_frame_elem = ttk.Frame(sub_tab1)
    btn_frame_elem.pack(pady=10)
    ttk.Button(btn_frame_elem, text="➕ Add New Element", command=popup_add_element).pack(side="left", padx=5)
    ttk.Button(btn_frame_elem, text="🗑️ Remove Selected", command=remove_selected_element).pack(side="left", padx=5)

    # ACID MANAGEMENT
    tk.Label(sub_tab2, text="Added Acids, Purities and Densities", font=("Segoe UI", 10, "bold"), bg="#f4f6f9").pack(pady=5)
    tree_acid = ttk.Treeview(sub_tab2, columns=("Acid", "Purity", "Density"), show="headings", height=10)
    tree_acid.heading("Acid", text="Acid Name", command=lambda: sort_treeview(tree_acid, "Acid", False))
    tree_acid.heading("Purity", text="Default Label Purity (%)", command=lambda: sort_treeview(tree_acid, "Purity", False))
    tree_acid.heading("Density", text="Stock Density (g/mL)", command=lambda: sort_treeview(tree_acid, "Density", False))
    tree_acid.column("Acid", width=220)
    tree_acid.column("Purity", width=150)
    tree_acid.column("Density", width=150)
    tree_acid.pack(fill="both", expand=True, padx=10, pady=5)

    def refresh_acid_tree():
        for row in tree_acid.get_children():
            tree_acid.delete(row)
        for k, v in db["acids"].items():
            tree_acid.insert("", "end", values=(k, v["purity"], v["density"]))
    refresh_acid_tree()

    def popup_add_acid():
        pop = tk.Toplevel(lib_win)
        pop.title("➕ Add New Acid")
        pop.geometry("360x250")
        pop.grab_set()

        tk.Label(pop, text="Acid Name (Ex: Phosphoric Acid (H3PO4)):").pack(pady=(10, 2))
        ent_name = ttk.Entry(pop, width=35)
        ent_name.pack()

        tk.Label(pop, text="Default Label Purity (%):").pack(pady=(10, 2))
        ent_purity = ttk.Entry(pop, width=35)
        ent_purity.insert(0, "85.0")
        ent_purity.pack()

        tk.Label(pop, text="Stock Density (g/mL):").pack(pady=(10, 2))
        ent_density = ttk.Entry(pop, width=35)
        ent_density.insert(0, "1.69")
        ent_density.pack()

        def save_acid():
            new_acid_name = ent_name.get().strip()
            try:
                new_acid_purity = float(ent_purity.get())
                new_acid_density = float(ent_density.get())
                if new_acid_name:
                    db["acids"][new_acid_name] = {
                        "purity": new_acid_purity,
                        "density": new_acid_density
                    }
                    save_data(db)
                    refresh_acid_tree()
                    update_main_dropdowns()
                    messagebox.showinfo("Success", f"'{new_acid_name}' Successfully added to the library!", parent=pop)
                    pop.destroy()
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers.", parent=pop)

        ttk.Button(pop, text="Add Acid", command=save_acid).pack(pady=15)

    def remove_selected_acid():
        selected = tree_acid.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an acid from the list to remove.", parent=lib_win)
            return
        acid_name = tree_acid.item(selected[0])["values"][0]
        if messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove '{acid_name}'?", parent=lib_win):
            del db["acids"][acid_name]
            save_data(db)
            refresh_acid_tree()
            update_main_dropdowns()

    btn_frame_acid = ttk.Frame(sub_tab2)
    btn_frame_acid.pack(pady=10)
    ttk.Button(btn_frame_acid, text="➕ Add New Acid", command=popup_add_acid).pack(side="left", padx=5)
    ttk.Button(btn_frame_acid, text="🗑️ Remove Selected", command=remove_selected_acid).pack(side="left", padx=5)

    # SOLUTE MANAGEMENT
    tk.Label(sub_tab3, text="Added Solutes", font=("Segoe UI", 10, "bold"), bg="#f4f6f9").pack(pady=5)
    list_solute = tk.Listbox(sub_tab3, font=("Segoe UI", 10), height=10)
    list_solute.pack(fill="both", expand=True, padx=10, pady=5)

    def refresh_solute_list():
        list_solute.delete(0, tk.END)
        for s in db["solutes"]:
            list_solute.insert(tk.END, s)
    refresh_solute_list()

    def popup_add_solute():
        pop = tk.Toplevel(lib_win)
        pop.title("➕ Add New Solute")
        pop.geometry("340x160")
        pop.grab_set()

        tk.Label(pop, text="Solute Name (Ex: Acetone):").pack(pady=(15, 2))
        ent_name = ttk.Entry(pop, width=35)
        ent_name.pack()

        def save_solute():
            new_solute_name = ent_name.get().strip()
            if new_solute_name:
                if new_solute_name not in db["solutes"]:
                    db["solutes"].append(new_solute_name)
                    save_data(db)
                    refresh_solute_list()
                    update_main_dropdowns()
                    messagebox.showinfo("Success", f"'{new_solute_name}' Successfully added to the library!", parent=pop)
                    pop.destroy()
                else:
                    messagebox.showwarning("Warning", "This solute has already been added!", parent=pop)

        ttk.Button(pop, text="Add Solute", command=save_solute).pack(pady=15)

    def remove_selected_solute():
        selected = list_solute.curselection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a solute from the list to remove.", parent=lib_win)
            return
        if len(db["solutes"]) <= 1:
            messagebox.showwarning("Warning", "At least one solute must remain in the library!", parent=lib_win)
            return
        solute_name = list_solute.get(selected[0])
        if messagebox.askyesno("Confirm Remove", f"Are you sure you want to remove '{solute_name}'?", parent=lib_win):
            db["solutes"].remove(solute_name)
            save_data(db)
            refresh_solute_list()
            update_main_dropdowns()

    btn_frame_solute = ttk.Frame(sub_tab3)
    btn_frame_solute.pack(pady=10)
    ttk.Button(btn_frame_solute, text="➕ Add New Solute", command=popup_add_solute).pack(side="left", padx=5)
    ttk.Button(btn_frame_solute, text="🗑️ Remove Selected", command=remove_selected_solute).pack(side="left", padx=5)

# HELP & ABOUT WINDOWS
def open_help_window():
    help_win = tk.Toplevel(root)
    help_win.title("📖 How to Use DiluCalc v3")
    help_win.geometry("640x560")
    help_win.configure(bg="#ffffff")
    help_win.grab_set()

    tk.Label(
        help_win, text="📖 User Guide & Instructions (v3)", 
        font=("Segoe UI", 13, "bold"), bg="#ffffff", fg="#1e293b"
    ).pack(anchor="w", padx=20, pady=(15, 5))

    guide_text = (
        "1. Standard Solution Preparation (Tab 1):\n"
        "   • Select Target Element, Concentration Unit (ppm, ppb, ppt), and Main Solute.\n"
        "   • Check 'Acid-only solution (No dissolved element)' if preparing a pure blank acid.\n\n"
        "2. Alloy Preparation (Tab 2):\n"
        "   • Set Desired Total Alloy Weight and Unit (mg or g).\n"
        "   • Add elements using '➕ Add Element' and check 'Balance' on one element.\n"
        "   • Enter 'Actual Weighed' values to dynamically compensate remaining elements.\n\n"
        "3. Solution Dilution & Adjustment (Tab 3):\n"
        "   • Check 'Acid-only solution (No element)' to dilute or concentrate pure acid solutions.\n"
        "   • Adjust Final Volume: Dilute or concentrate by adding directly to the full volume.\n"
        "   • Keep Volume Constant: Withdraw a calculated volume from your vessel and replace\n"
        "     with pure solute, stock acid, or spiking agent to maintain constant final volume.\n\n"
        "4. Managing the Library (Data -> See Library):\n"
        "   • View or add custom Elements, Acids, or Solutes saved to 'dilucalc_library.json'."
    )

    lbl_guide = tk.Label(
        help_win, text=guide_text, font=("Segoe UI", 9), 
        bg="#f8fafc", fg="#334155", justify="left", padx=15, pady=15,
        highlightbackground="#e2e8f0", highlightthickness=1
    )
    lbl_guide.pack(fill="both", expand=True, padx=20, pady=10)

    ttk.Button(help_win, text="Close", command=help_win.destroy).pack(pady=(0, 15))

def open_about_window():
    messagebox.showinfo(
        "About DiluCalc v3",
        "🧪 DiluCalc v3 - Solution, Alloy & Dilution Calculator\n\n"
        "A standalone laboratory utility for standard solution, acid matrix, alloy preparation, and concentration adjustment.\n\n"
        "© 2026 Mehmet Işık / resawlin. All rights reserved.\n"
        "Free for student and laboratory use.\n\n"
        "Commercial use, resale, or integration into any commercial product, in whole or in part, is strictly prohibited."
    )

# HEADER & MAIN NOTEBOOK (TABS)
header_frame = tk.Frame(root, bg="#1e293b", pady=10, padx=20)
header_frame.pack(fill="x")
tk.Label(header_frame, text="🧪 DiluCalc v3 - Laboratory Solution, Alloy & Dilution Calculator", font=("Segoe UI", 15, "bold"), fg="white", bg="#1e293b").pack(anchor="w")
tk.Label(header_frame, text="Precision calculations for standard preparations, multi-element alloys, and solution concentration/matrix adjustments.", font=("Segoe UI", 10), fg="#cbd5e1", bg="#1e293b").pack(anchor="w")

main_notebook = ttk.Notebook(root)
main_notebook.pack(fill="both", expand=True, padx=12, pady=8)

tab_solution = ttk.Frame(main_notebook, padding=10)
tab_alloy = ttk.Frame(main_notebook, padding=10)
tab_adjust = ttk.Frame(main_notebook, padding=10)

main_notebook.add(tab_solution, text="🧪 Standard Solution Prep")
main_notebook.add(tab_alloy, text="⚙️ Alloy Preparation")
main_notebook.add(tab_adjust, text="📉 Solution Dilution & Adjustment")

# ==============================================================================
# TAB 1: STANDARD SOLUTION PREPARATION
# ==============================================================================
tab_solution.columnconfigure(0, weight=1)
tab_solution.columnconfigure(1, weight=1)
tab_solution.rowconfigure(0, weight=1)

col_left = ttk.LabelFrame(tab_solution, text="⚙️ Parameters", padding=15)
col_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
col_left.columnconfigure(1, weight=1)

var_tab1_no_elem = tk.BooleanVar(value=False)
chk_tab1_no_elem = ttk.Checkbutton(col_left, text="Acid-only solution (No dissolved element / Blank matrix)", variable=var_tab1_no_elem)
chk_tab1_no_elem.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))

lbl_t1_elem = ttk.Label(col_left, text="Target Element:")
lbl_t1_elem.grid(row=1, column=0, sticky="w", pady=6)
cb_element = ttk.Combobox(col_left, state="readonly", values=list(db["elements"].keys()))
if db["elements"]:
    cb_element.current(0)
cb_element.grid(row=1, column=1, sticky="ew", pady=6, padx=5)

lbl_t1_conc = ttk.Label(col_left, text="Desired Concentration Value:")
lbl_t1_conc.grid(row=2, column=0, sticky="w", pady=6)
ent_conc = ttk.Entry(col_left)
ent_conc.insert(0, "114.0")
ent_conc.grid(row=2, column=1, sticky="ew", pady=6, padx=5)

lbl_t1_unit = ttk.Label(col_left, text="Concentration Unit:")
lbl_t1_unit.grid(row=3, column=0, sticky="w", pady=6)
cb_unit = ttk.Combobox(col_left, state="readonly", values=["ppb (µg/L)", "ppm (mg/L)", "ppt (ng/L)"])
cb_unit.current(0)
cb_unit.grid(row=3, column=1, sticky="ew", pady=6, padx=5)

ttk.Label(col_left, text="Final Total Volume (mL):").grid(row=4, column=0, sticky="w", pady=6)
ent_vol = ttk.Entry(col_left)
ent_vol.insert(0, "1000.0")
ent_vol.grid(row=4, column=1, sticky="ew", pady=6, padx=5)

ttk.Label(col_left, text="Main Solute:").grid(row=5, column=0, sticky="w", pady=6)
cb_solute = ttk.Combobox(col_left, state="readonly", values=db["solutes"])
if db["solutes"]:
    cb_solute.current(0)
cb_solute.grid(row=5, column=1, sticky="ew", pady=6, padx=5)

ttk.Label(col_left, text="Matrix Acid:").grid(row=6, column=0, sticky="w", pady=6)
cb_acid = ttk.Combobox(col_left, state="readonly", values=["No Acid"] + list(db["acids"].keys()))
cb_acid.current(0)
cb_acid.grid(row=6, column=1, sticky="ew", pady=6, padx=5)

acid_frame = tk.Frame(col_left, bg="#f8fafc", padx=8, pady=8, highlightbackground="#e2e8f0", highlightthickness=1)
acid_frame.grid(row=7, column=0, columnspan=2, sticky="ew", pady=6)
acid_frame.columnconfigure(1, weight=1)

tk.Label(acid_frame, text="Desired Acid Matrix Level (%):", bg="#f8fafc", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=3)
ent_acid_percent = ttk.Entry(acid_frame)
ent_acid_percent.insert(0, "2.0")
ent_acid_percent.grid(row=0, column=1, sticky="ew", pady=3, padx=5)

tk.Label(acid_frame, text="Purity of chosen acid (%):", bg="#f8fafc", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=3)
ent_acid_purity = ttk.Entry(acid_frame)
ent_acid_purity.insert(0, "65.0")
ent_acid_purity.grid(row=1, column=1, sticky="ew", pady=3, padx=5)

tk.Label(acid_frame, text="Acid density (g/mL):", bg="#f8fafc", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=3)
ent_acid_density = ttk.Entry(acid_frame)
ent_acid_density.insert(0, "1.42")
ent_acid_density.grid(row=2, column=1, sticky="ew", pady=3, padx=5)

var_beaker = tk.BooleanVar(value=False)
chk_beaker = ttk.Checkbutton(col_left, text="Include container tare mass", variable=var_beaker)
chk_beaker.grid(row=8, column=0, columnspan=2, sticky="w", pady=(10, 4))

lbl_beaker = ttk.Label(col_left, text="Mass of Beaker (g):")
ent_beaker = ttk.Entry(col_left)
ent_beaker.insert(0, "15.0")

# Column Right (Solution Instructions)
col_right = ttk.LabelFrame(tab_solution, text="📋 Instructions", padding=15)
col_right.grid(row=0, column=1, sticky="nsew")

weighing_box = tk.Frame(col_right, bg="#ffffff", padx=15, pady=12, highlightbackground="#cbd5e1", highlightthickness=1)
weighing_box.pack(fill="x", pady=(0, 12))

tk.Label(weighing_box, text="⚖️ WEIGHING INFO", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#64748b").pack(anchor="w")
lbl_weighing_summary = tk.Label(weighing_box, text="", font=("Segoe UI", 9), bg="#ffffff", fg="#334155", justify="left")
lbl_weighing_summary.pack(anchor="w", pady=(2, 8))

mass_highlight = tk.Frame(weighing_box, bg="#eff6ff", padx=12, pady=8, highlightbackground="#93c5fd", highlightthickness=1)
mass_highlight.pack(fill="x")
tk.Label(mass_highlight, text="🔹 Needed Pure Element Mass:", font=("Segoe UI", 9, "bold"), bg="#eff6ff", fg="#1e40af").pack(anchor="w")
lbl_mass_big = tk.Label(mass_highlight, text="", font=("Segoe UI", 15, "bold"), bg="#eff6ff", fg="#1d4ed8")
lbl_mass_big.pack(anchor="w", pady=(2, 0))

scale_frame = tk.Frame(weighing_box, bg="#dcfce7", padx=12, pady=8, highlightbackground="#86efac", highlightthickness=1)
tk.Label(scale_frame, text="📌 Scale Screen (Beaker + Element):", font=("Segoe UI", 9, "bold"), bg="#dcfce7", fg="#166534").pack(anchor="w")
lbl_scale_big = tk.Label(scale_frame, text="", font=("Segoe UI", 15, "bold"), bg="#dcfce7", fg="#15803d")
lbl_scale_big.pack(anchor="w", pady=(2, 0))

liquid_box = tk.Frame(col_right, bg="#ffffff", padx=15, pady=12, highlightbackground="#cbd5e1", highlightthickness=1)
liquid_box.pack(fill="x", pady=(0, 12))

tk.Label(liquid_box, text="🧪 LIQUID & MATRIX COMPONENTS", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#64748b").pack(anchor="w")
lbl_solute_title = tk.Label(liquid_box, text="", font=("Segoe UI", 9, "bold"), bg="#ffffff", fg="#334155")
lbl_solute_title.pack(anchor="w", pady=(2, 8))

acid_vol_frame = tk.Frame(liquid_box, bg="#fffbeb", padx=12, pady=8, highlightbackground="#fde68a", highlightthickness=1)
lbl_acid_title = tk.Label(acid_vol_frame, text="", font=("Segoe UI", 9, "bold"), bg="#fffbeb", fg="#b45309")
lbl_acid_title.pack(anchor="w")
lbl_acid_big = tk.Label(acid_vol_frame, text="", font=("Segoe UI", 14, "bold"), bg="#fffbeb", fg="#d97706")
lbl_acid_big.pack(anchor="w", pady=(2, 0))

solute_vol_frame = tk.Frame(liquid_box, bg="#f0fdfa", padx=12, pady=8, highlightbackground="#99f6e4", highlightthickness=1)
solute_vol_frame.pack(fill="x")
lbl_solute_vol_title = tk.Label(solute_vol_frame, text="", font=("Segoe UI", 9, "bold"), bg="#f0fdfa", fg="#0f766e")
lbl_solute_vol_title.pack(anchor="w")
lbl_solute_big = tk.Label(solute_vol_frame, text="", font=("Segoe UI", 13, "bold"), bg="#f0fdfa", fg="#0d9488")
lbl_solute_big.pack(anchor="w", pady=(2, 0))

hint_box = tk.Frame(col_right, bg="#fef9c3", padx=12, pady=10, highlightbackground="#fde047", highlightthickness=1)
hint_box.pack(fill="x", side="bottom")
tk.Label(
    hint_box, 
    text="⚠️ Warning: First add a portion of the solute into the beaker, then slowly add acid,\ndissolve the element (if applicable) and finally complete it to the final volume.", 
    font=("Segoe UI", 9, "italic"), bg="#fef9c3", fg="#854d0e", justify="left"
).pack(anchor="w")


# ==============================================================================
# TAB 2: ALLOY PREPARATION
# ==============================================================================
tab_alloy.columnconfigure(0, weight=1)
tab_alloy.columnconfigure(1, weight=1)
tab_alloy.rowconfigure(0, weight=1)

alloy_left = ttk.LabelFrame(tab_alloy, text="⚙️ Alloy Parameters & Elements", padding=12)
alloy_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
alloy_left.columnconfigure(0, weight=1)
alloy_left.rowconfigure(2, weight=1)

alloy_top_frame = tk.Frame(alloy_left, bg="#ffffff")
alloy_top_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
alloy_top_frame.columnconfigure(1, weight=1)

ttk.Label(alloy_top_frame, text="Desired Total Alloy Weight:").grid(row=0, column=0, sticky="w", pady=4)
ent_alloy_total = ttk.Entry(alloy_top_frame, width=14)
ent_alloy_total.insert(0, "10.0")
ent_alloy_total.grid(row=0, column=1, sticky="ew", pady=4, padx=5)

cb_alloy_unit = ttk.Combobox(alloy_top_frame, state="readonly", values=["mg", "g"], width=6)
cb_alloy_unit.current(0)
cb_alloy_unit.grid(row=0, column=2, sticky="w", pady=4)

var_alloy_tare = tk.BooleanVar(value=False)
chk_alloy_tare = ttk.Checkbutton(alloy_top_frame, text="Include container tare mass", variable=var_alloy_tare)
chk_alloy_tare.grid(row=1, column=0, columnspan=3, sticky="w", pady=(4, 2))

lbl_alloy_tare = ttk.Label(alloy_top_frame, text="Container / Beaker Mass:")
ent_alloy_tare = ttk.Entry(alloy_top_frame, width=14)
ent_alloy_tare.insert(0, "15.0")
cb_alloy_tare_unit = ttk.Combobox(alloy_top_frame, state="readonly", values=["g", "mg"], width=6)
cb_alloy_tare_unit.current(0)

alloy_btn_bar = tk.Frame(alloy_left, bg="#ffffff")
alloy_btn_bar.grid(row=1, column=0, sticky="ew", pady=(0, 6))
ttk.Button(alloy_btn_bar, text="➕ Add Element", command=lambda: add_alloy_element_row()).pack(side="left", padx=(0, 6))
ttk.Button(alloy_btn_bar, text="🧹 Clear Actual Weights", command=lambda: clear_alloy_actual_weights()).pack(side="left")

canvas_frame = tk.Frame(alloy_left, bg="#f8fafc", highlightbackground="#e2e8f0", highlightthickness=1)
canvas_frame.grid(row=2, column=0, sticky="nsew")
canvas_frame.columnconfigure(0, weight=1)
canvas_frame.rowconfigure(0, weight=1)

alloy_canvas = tk.Canvas(canvas_frame, bg="#f8fafc", highlightthickness=0)
alloy_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=alloy_canvas.yview)
alloy_rows_container = tk.Frame(alloy_canvas, bg="#f8fafc", padx=6, pady=6)

alloy_rows_container.bind(
    "<Configure>",
    lambda e: alloy_canvas.configure(scrollregion=alloy_canvas.bbox("all"))
)
canvas_win = alloy_canvas.create_window((0, 0), window=alloy_rows_container, anchor="nw")
alloy_canvas.bind(
    "<Configure>",
    lambda e: alloy_canvas.itemconfig(canvas_win, width=e.width)
)
alloy_canvas.configure(yscrollcommand=alloy_scrollbar.set)

alloy_canvas.grid(row=0, column=0, sticky="nsew")
alloy_scrollbar.grid(row=0, column=1, sticky="ns")

# Right Panel (Alloy Instructions & Actual Results - Modern Card Style)
alloy_right = ttk.LabelFrame(tab_alloy, text="📋 Instructions & Composition", padding=12)
alloy_right.grid(row=0, column=1, sticky="nsew")
alloy_right.columnconfigure(0, weight=1)
alloy_right.rowconfigure(0, weight=1)

right_canvas_frame = tk.Frame(alloy_right, bg="#f4f6f9")
right_canvas_frame.grid(row=0, column=0, sticky="nsew")
right_canvas_frame.columnconfigure(0, weight=1)
right_canvas_frame.rowconfigure(0, weight=1)

alloy_right_canvas = tk.Canvas(right_canvas_frame, bg="#f4f6f9", highlightthickness=0)
alloy_right_scroll = ttk.Scrollbar(right_canvas_frame, orient="vertical", command=alloy_right_canvas.yview)
alloy_right_content = tk.Frame(alloy_right_canvas, bg="#f4f6f9")

alloy_right_content.bind(
    "<Configure>",
    lambda e: alloy_right_canvas.configure(scrollregion=alloy_right_canvas.bbox("all"))
)
right_canvas_win = alloy_right_canvas.create_window((0, 0), window=alloy_right_content, anchor="nw")
alloy_right_canvas.bind(
    "<Configure>",
    lambda e: alloy_right_canvas.itemconfig(right_canvas_win, width=e.width)
)
alloy_right_canvas.configure(yscrollcommand=alloy_right_scroll.set)

alloy_right_canvas.grid(row=0, column=0, sticky="nsew")
alloy_right_scroll.grid(row=0, column=1, sticky="ns")

alloy_weighing_box = tk.Frame(alloy_right_content, bg="#ffffff", padx=15, pady=12, highlightbackground="#cbd5e1", highlightthickness=1)
alloy_weighing_box.pack(fill="x", pady=(0, 12))

tk.Label(alloy_weighing_box, text="⚖️ WEIGHING INFO", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#64748b").pack(anchor="w")
lbl_alloy_summary = tk.Label(alloy_weighing_box, text="", font=("Segoe UI", 9), bg="#ffffff", fg="#334155", justify="left")
lbl_alloy_summary.pack(anchor="w", pady=(2, 8))

alloy_cards_container = tk.Frame(alloy_weighing_box, bg="#ffffff")
alloy_cards_container.pack(fill="x")

alloy_comp_box = tk.Frame(alloy_right_content, bg="#ffffff", padx=15, pady=12, highlightbackground="#cbd5e1", highlightthickness=1)
alloy_comp_box.pack(fill="x", pady=(0, 12))

tk.Label(alloy_comp_box, text="🧪 ALLOY COMPOSITION & SUMMARY", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#64748b").pack(anchor="w")
lbl_alloy_status = tk.Label(alloy_comp_box, text="", font=("Segoe UI", 9, "bold"), bg="#ffffff", fg="#334155", justify="left")
lbl_alloy_status.pack(anchor="w", pady=(2, 8))

alloy_total_card = tk.Frame(alloy_comp_box, bg="#f0fdfa", padx=12, pady=8, highlightbackground="#99f6e4", highlightthickness=1)
alloy_total_card.pack(fill="x", pady=(0, 8))
lbl_alloy_total_title = tk.Label(alloy_total_card, text="💧 Resulting Total Alloy Mass:", font=("Segoe UI", 9, "bold"), bg="#f0fdfa", fg="#0f766e")
lbl_alloy_total_title.pack(anchor="w")
lbl_alloy_total_big = tk.Label(alloy_total_card, text="", font=("Segoe UI", 14, "bold"), bg="#f0fdfa", fg="#0d9488", justify="left")
lbl_alloy_total_big.pack(anchor="w", pady=(2, 0))

alloy_comp_card = tk.Frame(alloy_comp_box, bg="#eff6ff", padx=12, pady=8, highlightbackground="#93c5fd", highlightthickness=1)
alloy_comp_card.pack(fill="x")
tk.Label(alloy_comp_card, text="📊 Element Breakdown (Weight %  |  Atomic %):", font=("Segoe UI", 9, "bold"), bg="#eff6ff", fg="#1e40af").pack(anchor="w")
lbl_alloy_comp_details = tk.Label(alloy_comp_card, text="", font=("Segoe UI", 11, "bold"), bg="#eff6ff", fg="#1d4ed8", justify="left")
lbl_alloy_comp_details.pack(anchor="w", pady=(4, 0))

alloy_rows = []
updating_alloy = False

def on_alloy_tare_toggle():
    if var_alloy_tare.get():
        lbl_alloy_tare.grid(row=2, column=0, sticky="w", pady=4)
        ent_alloy_tare.grid(row=2, column=1, sticky="ew", pady=4, padx=5)
        cb_alloy_tare_unit.grid(row=2, column=2, sticky="w", pady=4)
    else:
        lbl_alloy_tare.grid_remove()
        ent_alloy_tare.grid_remove()
        cb_alloy_tare_unit.grid_remove()
    calculate_alloy()

chk_alloy_tare.config(command=on_alloy_tare_toggle)

def add_alloy_element_row(default_elem_idx=0, default_pct="50.0", is_balance=False):
    row_idx = len(alloy_rows) + 1
    card = tk.Frame(alloy_rows_container, bg="#ffffff", padx=10, pady=8, highlightbackground="#cbd5e1", highlightthickness=1)
    card.pack(fill="x", pady=4)
    card.columnconfigure(1, weight=1, minsize=155)

    lbl_num = tk.Label(card, text=f"#{row_idx}", font=("Segoe UI", 9, "bold"), bg="#ffffff", fg="#475569")
    lbl_num.grid(row=0, column=0, sticky="w", padx=(0, 6))

    elem_list = list(db["elements"].keys())
    cb_el = ttk.Combobox(card, state="readonly", values=elem_list, width=18)
    if elem_list:
        cb_el.current(min(default_elem_idx, len(elem_list) - 1))
    cb_el.grid(row=0, column=1, sticky="ew", padx=(0, 8))

    tk.Label(card, text="Target %:", bg="#ffffff", font=("Segoe UI", 9)).grid(row=0, column=2, sticky="e", padx=2)
    ent_pct = ttk.Entry(card, width=7)
    ent_pct.insert(0, default_pct)
    ent_pct.grid(row=0, column=3, sticky="w", padx=2)

    var_bal = tk.BooleanVar(value=is_balance)
    chk_bal = ttk.Checkbutton(card, text="Balance", variable=var_bal)
    chk_bal.grid(row=0, column=4, sticky="w", padx=6)

    btn_del = ttk.Button(card, text="✖", width=3)
    btn_del.grid(row=0, column=5, sticky="e", padx=(4, 0))

    act_frame = tk.Frame(card, bg="#ffffff")
    act_frame.grid(row=1, column=0, columnspan=6, sticky="w", pady=(8, 0))

    tk.Label(act_frame, text="Actual Weighed:", bg="#ffffff", fg="#64748b", font=("Segoe UI", 9, "italic")).pack(side="left", padx=(24, 6))
    ent_act = ttk.Entry(act_frame, width=14)
    ent_act.pack(side="left", padx=(0, 6))

    cb_act_u = ttk.Combobox(act_frame, state="readonly", values=["mg", "g"], width=5)
    cb_act_u.current(0)
    cb_act_u.pack(side="left", padx=(0, 6))

    tk.Label(act_frame, text="(optional override)", bg="#ffffff", fg="#94a3b8", font=("Segoe UI", 8)).pack(side="left")

    row_data = {
        "frame": card,
        "lbl_num": lbl_num,
        "cb_elem": cb_el,
        "ent_pct": ent_pct,
        "var_bal": var_bal,
        "chk_bal": chk_bal,
        "btn_del": btn_del,
        "ent_act": ent_act,
        "cb_act_unit": cb_act_u
    }
    alloy_rows.append(row_data)

    def remove_this_row():
        if len(alloy_rows) <= 2:
            messagebox.showwarning("Warning", "At least 2 elements are required for an alloy!")
            return
        card.destroy()
        alloy_rows.remove(row_data)
        for i, r in enumerate(alloy_rows, 1):
            r["lbl_num"].config(text=f"#{i}")
        sync_balance_states()
        calculate_alloy()

    btn_del.config(command=remove_this_row)
    chk_bal.config(command=lambda: on_balance_check_change())
    cb_el.bind("<<ComboboxSelected>>", lambda e: calculate_alloy())
    cb_act_u.bind("<<ComboboxSelected>>", lambda e: calculate_alloy())
    ent_pct.bind("<KeyRelease>", lambda e: calculate_alloy())
    ent_act.bind("<KeyRelease>", lambda e: calculate_alloy())

    sync_balance_states()
    calculate_alloy()

def clear_alloy_actual_weights():
    for r in alloy_rows:
        r["ent_act"].delete(0, tk.END)
    calculate_alloy()

def on_balance_check_change():
    sync_balance_states()
    calculate_alloy()

def sync_balance_states():
    any_bal = any(r["var_bal"].get() for r in alloy_rows)
    for r in alloy_rows:
        if any_bal:
            if r["var_bal"].get():
                r["chk_bal"].config(state="normal")
                r["ent_pct"].config(state="disabled")
            else:
                r["chk_bal"].config(state="disabled")
                r["ent_pct"].config(state="normal")
        else:
            r["chk_bal"].config(state="normal")
            r["ent_pct"].config(state="normal")

def calculate_alloy(*args):
    global updating_alloy
    if updating_alloy:
        return
    updating_alloy = True
    try:
        total_val = float(ent_alloy_total.get() or 0.0)
        total_unit = cb_alloy_unit.get()
        desired_total_mg = total_val * 1000.0 if total_unit == "g" else total_val

        bal_row = None
        other_pct_sum = 0.0
        for r in alloy_rows:
            if r["var_bal"].get():
                bal_row = r
            else:
                try:
                    val = float(r["ent_pct"].get() or 0.0)
                except ValueError:
                    val = 0.0
                other_pct_sum += val

        if bal_row is not None:
            rem_pct = max(0.0, 100.0 - other_pct_sum)
            bal_row["ent_pct"].config(state="normal")
            bal_row["ent_pct"].delete(0, tk.END)
            bal_row["ent_pct"].insert(0, f"{rem_pct:.4f}".rstrip("0").rstrip("."))
            bal_row["ent_pct"].config(state="disabled")

        elements_info = []
        total_pct_sum = 0.0
        weighed_mg_sum = 0.0
        unweighed_pct_sum = 0.0
        weighed_pct_sum = 0.0

        for idx, r in enumerate(alloy_rows, 1):
            el_name = r["cb_elem"].get()
            at_mass = db["elements"].get(el_name, 1.0)
            try:
                pct = float(r["ent_pct"].get() or 0.0)
            except ValueError:
                pct = 0.0
            total_pct_sum += pct

            act_str = r["ent_act"].get().strip()
            act_unit = r["cb_act_unit"].get()
            has_actual = False
            act_mg = 0.0
            if act_str != "":
                try:
                    act_val = float(act_str)
                    act_mg = act_val * 1000.0 if act_unit == "g" else act_val
                    has_actual = True
                    weighed_mg_sum += act_mg
                    weighed_pct_sum += pct
                except ValueError:
                    has_actual = False

            if not has_actual:
                unweighed_pct_sum += pct

            elements_info.append({
                "idx": idx,
                "name": el_name,
                "at_mass": at_mass,
                "target_pct": pct,
                "theo_mg": desired_total_mg * (pct / 100.0),
                "has_actual": has_actual,
                "act_mg": act_mg
            })

        tare_mg = 0.0
        if var_alloy_tare.get():
            try:
                t_val = float(ent_alloy_tare.get() or 0.0)
            except ValueError:
                t_val = 0.0
            tare_mg = t_val * 1000.0 if cb_alloy_tare_unit.get() == "g" else t_val

        summary_text = f"Desired Total Alloy: {desired_total_mg:.4f} mg ({desired_total_mg/1000.0:.6f} g)   |   Elements: {len(elements_info)}"
        if abs(total_pct_sum - 100.0) > 1e-3:
            summary_text += f"\n⚠️ Warning: Target percentages sum to {total_pct_sum:.2f}% (not 100%)!"
        lbl_alloy_summary.config(text=summary_text)

        for widget in alloy_cards_container.winfo_children():
            widget.destroy()

        any_weighed = any(e["has_actual"] for e in elements_info)
        all_weighed = all(e["has_actual"] for e in elements_info)

        running_scale_mg = tare_mg
        for e in elements_info:
            if e["has_actual"]:
                running_scale_mg += e["act_mg"]
                c_frame = tk.Frame(alloy_cards_container, bg="#dcfce7", padx=12, pady=8, highlightbackground="#86efac", highlightthickness=1)
                c_frame.pack(fill="x", pady=(0, 8))
                tk.Label(
                    c_frame,
                    text=f"✓ #{e['idx']} {e['name']} (Target: {e['target_pct']:.2f}%  |  A.M: {e['at_mass']} g/mol) — WEIGHED:",
                    font=("Segoe UI", 9, "bold"), bg="#dcfce7", fg="#166534"
                ).pack(anchor="w")
                tk.Label(
                    c_frame,
                    text=f"{e['act_mg']/1000.0:.6f} g  ({e['act_mg']:.4f} mg)",
                    font=("Segoe UI", 14, "bold"), bg="#dcfce7", fg="#15803d", justify="left"
                ).pack(anchor="w", pady=(2, 0))
                if var_alloy_tare.get():
                    tk.Label(
                        c_frame,
                        text=f"📌 Scale Screen (Cumulative): {running_scale_mg/1000.0:.6f} g  ({running_scale_mg:.4f} mg)",
                        font=("Segoe UI", 9, "bold"), bg="#dcfce7", fg="#166534"
                    ).pack(anchor="w", pady=(4, 0))

            elif not any_weighed:
                running_scale_mg += e["theo_mg"]
                c_frame = tk.Frame(alloy_cards_container, bg="#eff6ff", padx=12, pady=8, highlightbackground="#93c5fd", highlightthickness=1)
                c_frame.pack(fill="x", pady=(0, 8))
                tk.Label(
                    c_frame,
                    text=f"🔹 #{e['idx']} Needed {e['name']} Mass (Target: {e['target_pct']:.2f}%  |  A.M: {e['at_mass']} g/mol):",
                    font=("Segoe UI", 9, "bold"), bg="#eff6ff", fg="#1e40af"
                ).pack(anchor="w")
                tk.Label(
                    c_frame,
                    text=f"{e['theo_mg']/1000.0:.6f} g\n({e['theo_mg']:.4f} mg)",
                    font=("Segoe UI", 14, "bold"), bg="#eff6ff", fg="#1d4ed8", justify="left"
                ).pack(anchor="w", pady=(2, 0))
                if var_alloy_tare.get():
                    tk.Label(
                        c_frame,
                        text=f"📌 Scale Screen (Beaker + Elements): {running_scale_mg/1000.0:.6f} g  ({running_scale_mg:.4f} mg)",
                        font=("Segoe UI", 9, "bold"), bg="#eff6ff", fg="#1e40af"
                    ).pack(anchor="w", pady=(4, 0))

            else:
                rem_for_fixed_total = max(0.0, desired_total_mg - weighed_mg_sum)
                fixed_total_need_mg = rem_for_fixed_total * (e["target_pct"] / unweighed_pct_sum) if unweighed_pct_sum > 0 else 0.0

                if weighed_pct_sum > 0:
                    implied_total_mg = weighed_mg_sum / (weighed_pct_sum / 100.0)
                    keep_ratio_need_mg = implied_total_mg * (e["target_pct"] / 100.0)
                else:
                    keep_ratio_need_mg = e["theo_mg"]

                running_scale_mg += fixed_total_need_mg
                c_frame = tk.Frame(alloy_cards_container, bg="#fffbeb", padx=12, pady=8, highlightbackground="#fde68a", highlightthickness=1)
                c_frame.pack(fill="x", pady=(0, 8))
                tk.Label(
                    c_frame,
                    text=f"⚡ #{e['idx']} Adjusted Needed {e['name']} Mass (Target: {e['target_pct']:.2f}%):",
                    font=("Segoe UI", 9, "bold"), bg="#fffbeb", fg="#b45309"
                ).pack(anchor="w")
                tk.Label(
                    c_frame,
                    text=f"For {desired_total_mg:g} mg total:  {fixed_total_need_mg:.4f} mg  ({fixed_total_need_mg/1000.0:.6f} g)\n"
                         f"To keep exact % ratio:  {keep_ratio_need_mg:.4f} mg  ({keep_ratio_need_mg/1000.0:.6f} g)",
                    font=("Segoe UI", 12, "bold"), bg="#fffbeb", fg="#d97706", justify="left"
                ).pack(anchor="w", pady=(2, 0))
                if var_alloy_tare.get():
                    tk.Label(
                        c_frame,
                        text=f"📌 Scale Screen (for {desired_total_mg:g} mg total): {running_scale_mg/1000.0:.6f} g  ({running_scale_mg:.4f} mg)",
                        font=("Segoe UI", 9, "bold"), bg="#fffbeb", fg="#b45309"
                    ).pack(anchor="w", pady=(4, 0))

        final_masses = [(e, e["act_mg"] if e["has_actual"] else e["theo_mg"]) for e in elements_info]
        actual_total_mg = sum(m for _, m in final_masses)
        total_moles = sum((m / 1000.0) / e["at_mass"] for e, m in final_masses if e["at_mass"] > 0)

        if all_weighed:
            status_str = "Status: Final Weighed Alloy (All elements weighed)"
        elif any_weighed:
            status_str = "Status: Partial Weighing (Unweighed elements assumed at target)"
        else:
            status_str = "Status: Theoretical Recipe Preview"
        lbl_alloy_status.config(text=status_str)

        tot_text = f"{actual_total_mg/1000.0:.6f} g   ({actual_total_mg:.4f} mg)"
        if var_alloy_tare.get():
            final_scale = tare_mg + actual_total_mg
            tot_text += f"\n📌 Final Scale Screen: {final_scale/1000.0:.6f} g   ({final_scale:.4f} mg)"
        lbl_alloy_total_big.config(text=tot_text)

        comp_lines = []
        for e, m in final_masses:
            wt_pct = (m / actual_total_mg * 100.0) if actual_total_mg > 0 else 0.0
            moles = (m / 1000.0) / e["at_mass"] if e["at_mass"] > 0 else 0.0
            at_pct = (moles / total_moles * 100.0) if total_moles > 0 else 0.0
            comp_lines.append(f"• {e['name']}:   {wt_pct:.2f} wt%  (Target: {e['target_pct']:g}%)   |   {at_pct:.2f} at%   [{m:.4f} mg]")

        lbl_alloy_comp_details.config(text="\n".join(comp_lines))

    except Exception as ex:
        print("Alloy Calc Error:", ex)
    finally:
        updating_alloy = False

ent_alloy_total.bind("<KeyRelease>", calculate_alloy)
cb_alloy_unit.bind("<<ComboboxSelected>>", calculate_alloy)
ent_alloy_tare.bind("<KeyRelease>", calculate_alloy)
cb_alloy_tare_unit.bind("<<ComboboxSelected>>", calculate_alloy)


# ==============================================================================
# TAB 3: SOLUTION DILUTION & CONCENTRATION ADJUSTMENT
# ==============================================================================
tab_adjust.columnconfigure(0, weight=1)
tab_adjust.columnconfigure(1, weight=1)
tab_adjust.rowconfigure(0, weight=1)

adj_left = ttk.LabelFrame(tab_adjust, text="⚙️ Existing Solution & Target Adjustment Parameters", padding=15)
adj_left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
adj_left.columnconfigure(1, weight=1)

# 1. Volume Handling Mode
mode_frame = tk.Frame(adj_left, bg="#f8fafc", padx=10, pady=8, highlightbackground="#cbd5e1", highlightthickness=1)
mode_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
tk.Label(mode_frame, text="🧭 Volume Handling Mode:", font=("Segoe UI", 9, "bold"), bg="#f8fafc", fg="#1e293b").pack(anchor="w", pady=(0, 4))

var_adj_mode = tk.StringVar(value="adjust_vol")
rb_adj_vol = ttk.Radiobutton(
    mode_frame,
    text="Adjust Final Volume (Auto-calculate V2 or set custom final volume)",
    variable=var_adj_mode, value="adjust_vol"
)
rb_adj_vol.pack(anchor="w", pady=2)

# YENİ: Adjust Final Volume altındaki özel hedef hacim (Custom V2) seçeneği
custom_v2_frame = tk.Frame(mode_frame, bg="#f8fafc")
custom_v2_frame.pack(anchor="w", padx=(22, 0), pady=(2, 4))

var_adj_custom_v2 = tk.BooleanVar(value=False)
chk_adj_custom_v2 = ttk.Checkbutton(
    custom_v2_frame,
    text="Specify custom target final volume (V2, mL):",
    variable=var_adj_custom_v2
)
chk_adj_custom_v2.pack(side="left")

ent_adj_v2 = ttk.Entry(custom_v2_frame, width=10, state="disabled")
ent_adj_v2.insert(0, "1000.0")
ent_adj_v2.pack(side="left", padx=6)

rb_keep_vol = ttk.Radiobutton(
    mode_frame,
    text="Keep Volume Constant (Withdraw excess solution & top back up to V1)",
    variable=var_adj_mode, value="keep_vol"
)
rb_keep_vol.pack(anchor="w", pady=2)

# Checkbox: Acid-Only Solution (No Element)
var_adj_no_elem = tk.BooleanVar(value=False)
chk_adj_no_elem = ttk.Checkbutton(
    adj_left,
    text="Acid-only solution (No dissolved element — Dilute or concentrate acid only)",
    variable=var_adj_no_elem
)
chk_adj_no_elem.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 6))

# 2. Element & Solute Selection
lbl_adj_elem = ttk.Label(adj_left, text="Target Element:")
lbl_adj_elem.grid(row=2, column=0, sticky="w", pady=4)
cb_adj_element = ttk.Combobox(adj_left, state="readonly", values=list(db["elements"].keys()))
if db["elements"]:
    cb_adj_element.current(0)
cb_adj_element.grid(row=2, column=1, columnspan=2, sticky="ew", pady=4, padx=5)

ttk.Label(adj_left, text="Main Solute:").grid(row=3, column=0, sticky="w", pady=4)
cb_adj_solute = ttk.Combobox(adj_left, state="readonly", values=db["solutes"])
if db["solutes"]:
    cb_adj_solute.current(0)
cb_adj_solute.grid(row=3, column=1, columnspan=2, sticky="ew", pady=4, padx=5)

# 3. Current & Target Concentration / Volume
ttk.Label(adj_left, text="Current Solution Volume (V1, mL):").grid(row=4, column=0, sticky="w", pady=4)
ent_adj_v1 = ttk.Entry(adj_left)
ent_adj_v1.insert(0, "500.0")
ent_adj_v1.grid(row=4, column=1, columnspan=2, sticky="ew", pady=4, padx=5)

lbl_adj_c1 = ttk.Label(adj_left, text="Current Element Conc. (C1):")
lbl_adj_c1.grid(row=5, column=0, sticky="w", pady=4)
ent_adj_c1 = ttk.Entry(adj_left)
ent_adj_c1.insert(0, "100.0")
ent_adj_c1.grid(row=5, column=1, sticky="ew", pady=4, padx=5)
cb_adj_u1 = ttk.Combobox(adj_left, state="readonly", values=["ppm (mg/L)", "ppb (µg/L)", "ppt (ng/L)"], width=12)
cb_adj_u1.current(0)
cb_adj_u1.grid(row=5, column=2, sticky="w", pady=4)

lbl_adj_c2 = ttk.Label(adj_left, text="Target Element Conc. (C2):")
lbl_adj_c2.grid(row=6, column=0, sticky="w", pady=4)
ent_adj_c2 = ttk.Entry(adj_left)
ent_adj_c2.insert(0, "25.0")
ent_adj_c2.grid(row=6, column=1, sticky="ew", pady=4, padx=5)
cb_adj_u2 = ttk.Combobox(adj_left, state="readonly", values=["ppm (mg/L)", "ppb (µg/L)", "ppt (ng/L)"], width=12)
cb_adj_u2.current(0)
cb_adj_u2.grid(row=6, column=2, sticky="w", pady=4)

# 4. Spiking Configuration Frame (Shown automatically when C2 > C1 and element mode active)
spike_frame = tk.Frame(adj_left, bg="#eff6ff", padx=10, pady=8, highlightbackground="#93c5fd", highlightthickness=1)
spike_frame.grid(row=7, column=0, columnspan=3, sticky="ew", pady=5)
spike_frame.columnconfigure(1, weight=1)

tk.Label(spike_frame, text="📈 Concentrating / Spiking Source (C2 > C1):", font=("Segoe UI", 9, "bold"), bg="#eff6ff", fg="#1e40af").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

var_spike_type = tk.StringVar(value="solid")
rb_spike_solid = tk.Radiobutton(spike_frame, text="Pure Solid Element (g/mg)", variable=var_spike_type, value="solid", bg="#eff6ff", font=("Segoe UI", 9))
rb_spike_solid.grid(row=1, column=0, sticky="w")
rb_spike_stock = tk.Radiobutton(spike_frame, text="Concentrated Stock Solution", variable=var_spike_type, value="stock", bg="#eff6ff", font=("Segoe UI", 9))
rb_spike_stock.grid(row=1, column=1, columnspan=2, sticky="w", padx=8)

lbl_spike_stock = tk.Label(spike_frame, text="Stock Solution Conc. (Cs):", bg="#eff6ff", font=("Segoe UI", 9))
ent_spike_stock = ttk.Entry(spike_frame, width=12)
ent_spike_stock.insert(0, "1000.0")
cb_spike_unit = ttk.Combobox(spike_frame, state="readonly", values=["ppm (mg/L)", "ppb (µg/L)", "ppt (ng/L)"], width=12)
cb_spike_unit.current(0)

lbl_spike_acid = tk.Label(spike_frame, text="Stock Acid Level (%):", bg="#eff6ff", font=("Segoe UI", 9))
ent_spike_acid = ttk.Entry(spike_frame, width=12)
ent_spike_acid.insert(0, "2.0")

# 5. Matrix Acid Adjustment Frame
ttk.Label(adj_left, text="Matrix Acid:").grid(row=8, column=0, sticky="w", pady=4)
cb_adj_acid = ttk.Combobox(adj_left, state="readonly", values=["No Acid"] + list(db["acids"].keys()))
cb_adj_acid.current(0)
cb_adj_acid.grid(row=8, column=1, columnspan=2, sticky="ew", pady=4, padx=5)

adj_acid_frame = tk.Frame(adj_left, bg="#f8fafc", padx=8, pady=8, highlightbackground="#e2e8f0", highlightthickness=1)
adj_acid_frame.grid(row=9, column=0, columnspan=3, sticky="ew", pady=5)
adj_acid_frame.columnconfigure(1, weight=1)

tk.Label(adj_acid_frame, text="Current Solution Acid (%):", bg="#f8fafc", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=2)
ent_adj_curr_acid = ttk.Entry(adj_acid_frame, width=10)
ent_adj_curr_acid.insert(0, "5.0")
ent_adj_curr_acid.grid(row=0, column=1, sticky="ew", pady=2, padx=5)

tk.Label(adj_acid_frame, text="Target Final Acid (%):", bg="#f8fafc", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=2)
ent_adj_targ_acid = ttk.Entry(adj_acid_frame, width=10)
ent_adj_targ_acid.insert(0, "2.0")
ent_adj_targ_acid.grid(row=1, column=1, sticky="ew", pady=2, padx=5)

tk.Label(adj_acid_frame, text="Stock Acid Purity (%):", bg="#f8fafc", font=("Segoe UI", 9)).grid(row=2, column=0, sticky="w", pady=2)
ent_adj_acid_pur = ttk.Entry(adj_acid_frame, width=10)
ent_adj_acid_pur.insert(0, "65.0")
ent_adj_acid_pur.grid(row=2, column=1, sticky="ew", pady=2, padx=5)

tk.Label(adj_acid_frame, text="Stock Acid Density (g/mL):", bg="#f8fafc", font=("Segoe UI", 9)).grid(row=3, column=0, sticky="w", pady=2)
ent_adj_acid_den = ttk.Entry(adj_acid_frame, width=10)
ent_adj_acid_den.insert(0, "1.42")
ent_adj_acid_den.grid(row=3, column=1, sticky="ew", pady=2, padx=5)

# Tare option for Solid Spiking
var_adj_tare = tk.BooleanVar(value=False)
chk_adj_tare = ttk.Checkbutton(adj_left, text="Include container tare mass (for solid spiking)", variable=var_adj_tare)
chk_adj_tare.grid(row=10, column=0, columnspan=3, sticky="w", pady=(6, 2))

lbl_adj_tare = ttk.Label(adj_left, text="Mass of Beaker (g):")
ent_adj_tare = ttk.Entry(adj_left)
ent_adj_tare.insert(0, "15.0")

# Right Column (Tab 3 Instructions - Modern Card Style)
adj_right = ttk.LabelFrame(tab_adjust, text="📋 Adjustment & Dilution Instructions", padding=15)
adj_right.grid(row=0, column=1, sticky="nsew")

adj_step_box = tk.Frame(adj_right, bg="#ffffff", padx=15, pady=12, highlightbackground="#cbd5e1", highlightthickness=1)
adj_step_box.pack(fill="x", pady=(0, 12))

tk.Label(adj_step_box, text="⚖️ CONCENTRATION & VOLUME ADJUSTMENT STEPS", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#64748b").pack(anchor="w")
lbl_adj_summary = tk.Label(adj_step_box, text="", font=("Segoe UI", 9), bg="#ffffff", fg="#334155", justify="left")
lbl_adj_summary.pack(anchor="w", pady=(2, 8))

# Withdraw / Keep Card
adj_withdraw_card = tk.Frame(adj_step_box, bg="#fef2f2", padx=12, pady=8, highlightbackground="#fca5a5", highlightthickness=1)
lbl_adj_withdraw_title = tk.Label(adj_withdraw_card, text="📤 Step 1 — Solution Volume to Withdraw / Remove:", font=("Segoe UI", 9, "bold"), bg="#fef2f2", fg="#991b1b")
lbl_adj_withdraw_title.pack(anchor="w")
lbl_adj_withdraw_big = tk.Label(adj_withdraw_card, text="", font=("Segoe UI", 14, "bold"), bg="#fef2f2", fg="#dc2626", justify="left")
lbl_adj_withdraw_big.pack(anchor="w", pady=(2, 0))

# Spiking / Element Addition Card
adj_spike_card = tk.Frame(adj_step_box, bg="#eff6ff", padx=12, pady=8, highlightbackground="#93c5fd", highlightthickness=1)
lbl_adj_spike_title = tk.Label(adj_spike_card, text="🔹 Element / Stock Addition:", font=("Segoe UI", 9, "bold"), bg="#eff6ff", fg="#1e40af")
lbl_adj_spike_title.pack(anchor="w")
lbl_adj_spike_big = tk.Label(adj_spike_card, text="", font=("Segoe UI", 14, "bold"), bg="#eff6ff", fg="#1d4ed8", justify="left")
lbl_adj_spike_big.pack(anchor="w", pady=(2, 0))

# Tare Scale Card (for Solid Spiking)
adj_scale_card = tk.Frame(adj_step_box, bg="#dcfce7", padx=12, pady=8, highlightbackground="#86efac", highlightthickness=1)
tk.Label(adj_scale_card, text="📌 Scale Screen (Beaker + Added Element):", font=("Segoe UI", 9, "bold"), bg="#dcfce7", fg="#166534").pack(anchor="w")
lbl_adj_scale_big = tk.Label(adj_scale_card, text="", font=("Segoe UI", 14, "bold"), bg="#dcfce7", fg="#15803d", justify="left")
lbl_adj_scale_big.pack(anchor="w", pady=(2, 0))

# Liquid & Acid Matrix Box
adj_liquid_box = tk.Frame(adj_right, bg="#ffffff", padx=15, pady=12, highlightbackground="#cbd5e1", highlightthickness=1)
adj_liquid_box.pack(fill="x", pady=(0, 12))

tk.Label(adj_liquid_box, text="🧪 SOLVENT & ACID MATRIX BALANCING", font=("Segoe UI", 10, "bold"), bg="#ffffff", fg="#64748b").pack(anchor="w")
lbl_adj_liq_summary = tk.Label(adj_liquid_box, text="", font=("Segoe UI", 9, "bold"), bg="#ffffff", fg="#334155", justify="left")
lbl_adj_liq_summary.pack(anchor="w", pady=(2, 8))

adj_acid_card = tk.Frame(adj_liquid_box, bg="#fffbeb", padx=12, pady=8, highlightbackground="#fde68a", highlightthickness=1)
lbl_adj_acid_title = tk.Label(adj_acid_card, text="", font=("Segoe UI", 9, "bold"), bg="#fffbeb", fg="#b45309")
lbl_adj_acid_title.pack(anchor="w")
lbl_adj_acid_big = tk.Label(adj_acid_card, text="", font=("Segoe UI", 14, "bold"), bg="#fffbeb", fg="#d97706", justify="left")
lbl_adj_acid_big.pack(anchor="w", pady=(2, 0))

adj_solute_card = tk.Frame(adj_liquid_box, bg="#f0fdfa", padx=12, pady=8, highlightbackground="#99f6e4", highlightthickness=1)
adj_solute_card.pack(fill="x")
lbl_adj_solute_title = tk.Label(adj_solute_card, text="", font=("Segoe UI", 9, "bold"), bg="#f0fdfa", fg="#0f766e")
lbl_adj_solute_title.pack(anchor="w")
lbl_adj_solute_big = tk.Label(adj_solute_card, text="", font=("Segoe UI", 13, "bold"), bg="#f0fdfa", fg="#0d9488", justify="left")
lbl_adj_solute_big.pack(anchor="w", pady=(2, 0))

adj_hint_box = tk.Frame(adj_right, bg="#fef9c3", padx=12, pady=10, highlightbackground="#fde047", highlightthickness=1)
adj_hint_box.pack(fill="x", side="bottom")
lbl_adj_hint = tk.Label(
    adj_hint_box,
    text="💡 Tip: In 'Keep Volume Constant' mode, withdraw the specified volume first,\nthen add required acid/spike and complete to the original volume with main solute.",
    font=("Segoe UI", 9, "italic"), bg="#fef9c3", fg="#854d0e", justify="left"
)
lbl_adj_hint.pack(anchor="w")


# ==============================================================================
# HELPER & CALCULATOR FUNCTIONS
# ==============================================================================
def unit_to_ppm_factor(unit_str):
    if "ppm" in unit_str:
        return 1.0
    elif "ppb" in unit_str:
        return 1e-3
    elif "ppt" in unit_str:
        return 1e-6
    return 1.0

def format_exact_decimal(val, min_decimals=6):
    if val == 0:
        return f"{0.0:.{min_decimals}f}"
    if abs(val) >= 1e-4:
        return f"{val:.{min_decimals}f}"
    needed_decimals = max(min_decimals, -int(math.floor(math.log10(abs(val)))) + 2)
    needed_decimals = min(needed_decimals, 15)
    return f"{val:.{needed_decimals}f}"

def format_smart_subunits(mass_g):
    mass_mg = mass_g * 1e3
    mass_ug = mass_g * 1e6
    mass_ng = mass_g * 1e9
    mass_pg = mass_g * 1e12

    if mass_g == 0:
        return "0.0000 mg"
    elif mass_mg >= 0.01:
        return f"{mass_mg:.4f} mg  |  {mass_ug:.2f} µg"
    elif mass_ug >= 0.01:
        return f"{mass_ug:.4f} µg  |  {mass_ng:.2f} ng"
    elif mass_ng >= 0.01:
        return f"{mass_ng:.4f} ng  |  {mass_pg:.1f} pg"
    else:
        return f"{mass_ng:.6f} ng  |  {mass_pg:.3f} pg"

# TAB 1 CALCULATOR
def on_tab1_no_elem_toggle():
    no_elem = var_tab1_no_elem.get()
    if no_elem:
        cb_element.config(state="disabled")
        ent_conc.config(state="disabled")
        cb_unit.config(state="disabled")
        chk_beaker.config(state="disabled")
        lbl_beaker.grid_remove()
        ent_beaker.grid_remove()
        # Automatically select first acid if "No Acid" was selected
        if cb_acid.get() == "No Acid" and len(cb_acid["values"]) > 1:
            cb_acid.current(1)
            on_acid_change()
    else:
        cb_element.config(state="readonly")
        ent_conc.config(state="normal")
        cb_unit.config(state="readonly")
        chk_beaker.config(state="normal")
        on_beaker_toggle()
    calculate()

def calculate(*args):
    try:
        no_elem = var_tab1_no_elem.get()
        chosen_element = cb_element.get()
        atkes = db["elements"].get(chosen_element, 0.0)

        concentration_value = float(ent_conc.get() or 0.0)
        concentration_unit = cb_unit.get()
        final_volume_ml = float(ent_vol.get() or 0.0)
        solute = cb_solute.get()
        acid_choosing = cb_acid.get()

        if no_elem:
            weighing_box.pack_forget()
        else:
            weighing_box.pack(fill="x", pady=(0, 12), before=liquid_box)
            factor = unit_to_ppm_factor(concentration_unit)
            needed_element_mg = (concentration_value * factor) * (final_volume_ml / 1000.0)
            needed_element_g = needed_element_mg / 1000.0

            w_summary = (
                f"Target Element: {chosen_element} (A.M: {atkes} g/mol)\n"
                f"Target Concentration: {concentration_value} {concentration_unit}   |   Final Volume: {final_volume_ml} mL"
            )
            lbl_weighing_summary.config(text=w_summary)

            g_str = format_exact_decimal(needed_element_g, min_decimals=6)
            sub_str = format_smart_subunits(needed_element_g)
            lbl_mass_big.config(text=f"{g_str} g\n({sub_str})", justify="left")

            if var_beaker.get():
                beaker_mass = float(ent_beaker.get() or 0.0)
                total_desired_weight_g = beaker_mass + needed_element_g
                total_desired_weight_mg = total_desired_weight_g * 1000.0
                
                elem_decimals = len(g_str.split(".")[1]) if "." in g_str else 6
                mg_decimals = max(4, elem_decimals - 3)
                
                lbl_scale_big.config(
                    text=f"{total_desired_weight_g:.{elem_decimals}f} g\n({total_desired_weight_mg:.{mg_decimals}f} mg)",
                    justify="left"
                )
                scale_frame.pack(fill="x", pady=(8, 0))
            else:
                scale_frame.pack_forget()

        lbl_solute_title.config(text=f"Main Solute: {solute}   |   Final Volume: {final_volume_ml} mL")

        if acid_choosing != "No Acid":
            acid_percent = float(ent_acid_percent.get() or 0.0)
            st_acid_concentration = float(ent_acid_purity.get() or 1.0)
            acid_density = float(ent_acid_density.get() or 1.0)

            solution_mass_g = final_volume_ml
            needed_pure_acid_g = solution_mass_g * (acid_percent / 100.0)
            needed_acid_g = needed_pure_acid_g / (st_acid_concentration / 100.0)
            needed_pure_acid_ml = needed_acid_g / acid_density
            solute_left_ml = final_volume_ml - needed_pure_acid_ml

            lbl_acid_title.config(text=f"🧪 Stock Acid to be added (%{st_acid_concentration} {acid_choosing}):")
            lbl_acid_big.config(text=f"{needed_pure_acid_ml:.4f} mL")
            acid_vol_frame.pack(fill="x", pady=(0, 8), before=solute_vol_frame)

            lbl_solute_vol_title.config(text=f"💧 Needed {solute} Volume:")
            lbl_solute_big.config(text=f"Approx. {solute_left_ml:.4f} mL")
        else:
            acid_vol_frame.pack_forget()
            lbl_solute_vol_title.config(text=f"💧 Needed {solute} Volume:")
            lbl_solute_big.config(text=f"To complete the volume to {final_volume_ml} mL")

    except Exception as e:
        print("Calc Error:", e)

def on_acid_change(event=None):
    acid_choosing = cb_acid.get()
    if acid_choosing != "No Acid":
        acid_frame.grid(row=7, column=0, columnspan=2, sticky="ew", pady=6)
        default_purity = db["acids"][acid_choosing]["purity"]
        default_density = db["acids"][acid_choosing]["density"]
        ent_acid_purity.delete(0, tk.END)
        ent_acid_purity.insert(0, str(default_purity))
        ent_acid_density.delete(0, tk.END)
        ent_acid_density.insert(0, str(default_density))
    else:
        acid_frame.grid_remove()
    calculate()

def on_beaker_toggle():
    if var_beaker.get() and not var_tab1_no_elem.get():
        lbl_beaker.grid(row=9, column=0, sticky="w", pady=4)
        ent_beaker.grid(row=9, column=1, sticky="ew", pady=4, padx=5)
    else:
        lbl_beaker.grid_remove()
        ent_beaker.grid_remove()
    calculate()


# TAB 3 CALCULATOR (SOLUTION DILUTION & ADJUSTMENT)
def on_adj_no_elem_toggle():
    no_elem = var_adj_no_elem.get()
    if no_elem:
        cb_adj_element.config(state="disabled")
        ent_adj_c1.config(state="disabled")
        cb_adj_u1.config(state="disabled")
        ent_adj_c2.config(state="disabled")
        cb_adj_u2.config(state="disabled")
        chk_adj_tare.config(state="disabled")
        lbl_adj_tare.grid_remove()
        ent_adj_tare.grid_remove()
        spike_frame.grid_remove()
        if cb_adj_acid.get() == "No Acid" and len(cb_adj_acid["values"]) > 1:
            cb_adj_acid.current(1)
            on_adj_acid_change()
    else:
        cb_adj_element.config(state="readonly")
        ent_adj_c1.config(state="normal")
        cb_adj_u1.config(state="readonly")
        ent_adj_c2.config(state="normal")
        cb_adj_u2.config(state="readonly")
        chk_adj_tare.config(state="normal")
        on_adj_tare_toggle()
    calculate_adjustment()

def on_adj_acid_change(event=None):
    acid_choosing = cb_adj_acid.get()
    if acid_choosing != "No Acid":
        adj_acid_frame.grid(row=9, column=0, columnspan=3, sticky="ew", pady=5)
        default_purity = db["acids"][acid_choosing]["purity"]
        default_density = db["acids"][acid_choosing]["density"]
        ent_adj_acid_pur.delete(0, tk.END)
        ent_adj_acid_pur.insert(0, str(default_purity))
        ent_adj_acid_den.delete(0, tk.END)
        ent_adj_acid_den.insert(0, str(default_density))
    else:
        adj_acid_frame.grid_remove()
    calculate_adjustment()

def on_spike_type_change():
    if var_spike_type.get() == "stock":
        lbl_spike_stock.grid(row=2, column=0, sticky="w", pady=3)
        ent_spike_stock.grid(row=2, column=1, sticky="ew", pady=3, padx=5)
        cb_spike_unit.grid(row=2, column=2, sticky="w", pady=3)
        if cb_adj_acid.get() != "No Acid":
            lbl_spike_acid.grid(row=3, column=0, sticky="w", pady=3)
            ent_spike_acid.grid(row=3, column=1, sticky="ew", pady=3, padx=5)
        else:
            lbl_spike_acid.grid_remove()
            ent_spike_acid.grid_remove()
    else:
        lbl_spike_stock.grid_remove()
        ent_spike_stock.grid_remove()
        cb_spike_unit.grid_remove()
        lbl_spike_acid.grid_remove()
        ent_spike_acid.grid_remove()
    calculate_adjustment()

def on_adj_tare_toggle():
    if var_adj_tare.get() and not var_adj_no_elem.get():
        lbl_adj_tare.grid(row=11, column=0, sticky="w", pady=4)
        ent_adj_tare.grid(row=11, column=1, columnspan=2, sticky="ew", pady=4, padx=5)
    else:
        lbl_adj_tare.grid_remove()
        ent_adj_tare.grid_remove()
    calculate_adjustment()
def on_adj_mode_or_v2_toggle():
    if var_adj_mode.get() == "adjust_vol":
        chk_adj_custom_v2.config(state="normal")
        if var_adj_custom_v2.get():
            ent_adj_v2.config(state="normal")
        else:
            ent_adj_v2.config(state="disabled")
    else:
        chk_adj_custom_v2.config(state="disabled")
        ent_adj_v2.config(state="disabled")
    calculate_adjustment()

def calculate_adjustment(*args):
    try:
        no_elem = var_adj_no_elem.get()
        elem = cb_adj_element.get()
        at_mass = db["elements"].get(elem, 0.0)
        solute = cb_adj_solute.get()
        acid_name = cb_adj_acid.get()
        mode = var_adj_mode.get()  # "adjust_vol" or "keep_vol"
        use_custom_v2 = (mode == "adjust_vol" and var_adj_custom_v2.get())
        v1 = float(ent_adj_v1.get() or 0.0)
        custom_v2 = float(ent_adj_v2.get() or 0.0) if use_custom_v2 else 0.0

        adj_withdraw_card.pack_forget()
        adj_spike_card.pack_forget()
        adj_scale_card.pack_forget()
        adj_acid_card.pack_forget()

        if v1 <= 0 or (use_custom_v2 and custom_v2 <= 0):
            lbl_adj_summary.config(text="Please enter valid positive volume values.")
            return

        # ======================================================================
        # MODE 1: ACID-ONLY SOLUTION (NO ELEMENT)
        # ======================================================================
        if no_elem:
            spike_frame.grid_remove()
            if acid_name == "No Acid":
                lbl_adj_summary.config(text="⚠️ Please select a Matrix Acid to dilute or concentrate in Acid-Only mode.")
                lbl_adj_liq_summary.config(text=f"Main Solute: {solute}")
                lbl_adj_solute_title.config(text=f"💧 Needed {solute} Volume:")
                lbl_adj_solute_big.config(text="Select an acid from 'Matrix Acid' dropdown on the left.")
                return

            curr_a = float(ent_adj_curr_acid.get() or 0.0)
            targ_a = float(ent_adj_targ_acid.get() or 0.0)
            st_pur = float(ent_adj_acid_pur.get() or 1.0)
            st_den = float(ent_adj_acid_den.get() or 1.0)
            eff_stock_pct = st_pur * st_den

            if targ_a <= 0 or curr_a < 0:
                lbl_adj_summary.config(text="Please enter valid acid percentages.")
                return

            # Subcase 1A: Acid Dilution (Target Acid % < Current Acid %)
            if targ_a < curr_a:
                op_label = f"ACID-ONLY DILUTION ({curr_a}% ➔ {targ_a}% {acid_name})"
                v_final = custom_v2 if use_custom_v2 else (v1 * (curr_a / targ_a) if mode == "adjust_vol" else v1)
                v_needed_from_v1 = v_final * (targ_a / curr_a)

                if v_needed_from_v1 <= v1 + 1e-6:
                    # Elimizdeki V1 yeterli (içinden v_needed_from_v1 kadar kullanacağız)
                    v_kept = min(v1, v_needed_from_v1)
                    v_withdraw = v1 - v_kept
                    solute_add = v_final - v_kept

                    if v_withdraw > 1e-6:
                        lbl_adj_withdraw_title.config(text="📤 Step 1 — Take / Withdraw From Initial Acid Solution:")
                        lbl_adj_withdraw_big.config(
                            text=f"Use {v_kept:.4f} mL of current solution\n(Withdraw / set aside excess {v_withdraw:.4f} mL)"
                        )
                        adj_withdraw_card.pack(fill="x", pady=(0, 8))
                    else:
                        lbl_adj_spike_title.config(text="🔹 Step 1 — Keep Full Initial Acid Solution:", bg="#eff6ff", fg="#1e40af")
                        lbl_adj_spike_big.config(text=f"Keep all {v1:.2f} mL   →   Expand to {v_final:.2f} mL", bg="#eff6ff", fg="#1d4ed8")
                        adj_spike_card.pack(fill="x", pady=(0, 8))

                    lbl_adj_solute_title.config(text=f"💧 Step 2 — Needed {solute} Volume to Add:")
                    lbl_adj_solute_big.config(text=f"{solute_add:.4f} mL   (Complete to {v_final:.2f} mL)")
                    lbl_adj_liq_summary.config(text=f"Main Solute: {solute}   |   No extra stock acid required")
                else:
                    # Custom V2 çok büyük seçildi, V1'in içindeki asit yetmiyor -> V1'in tamamını kullan + eksik asidi stoktan tamamla!
                    v_kept = v1
                    pure_acid_deficit_g = (v_final * targ_a - v1 * curr_a) / 100.0
                    stock_acid_ml = (pure_acid_deficit_g / (st_pur / 100.0)) / st_den
                    solute_add = v_final - v1 - stock_acid_ml

                    lbl_adj_spike_title.config(text="🔹 Step 1 — Keep Full Initial Acid Solution:", bg="#eff6ff", fg="#1e40af")
                    lbl_adj_spike_big.config(text=f"Keep all {v1:.2f} mL (Target V2={v_final:.2f} mL requires extra stock acid)", bg="#eff6ff", fg="#1d4ed8")
                    adj_spike_card.pack(fill="x", pady=(0, 8))

                    lbl_adj_acid_title.config(text=f"🧪 Step 2 — Extra Stock Acid Needed (%{st_pur} {acid_name}):")
                    lbl_adj_acid_big.config(text=f"{stock_acid_ml:.4f} mL")
                    adj_acid_card.pack(fill="x", pady=(0, 8), before=adj_solute_card)

                    lbl_adj_solute_title.config(text=f"💧 Step 3 — Needed {solute} Volume:")
                    lbl_adj_solute_big.config(text=f"{max(0.0, solute_add):.4f} mL   (Complete to {v_final:.2f} mL)")
                    lbl_adj_liq_summary.config(text=f"Main Solute: {solute}   |   Extra Stock Acid Added")

                lbl_adj_summary.config(text=f"Operation: {op_label}\nInitial Volume: {v1:.2f} mL   |   Final Volume: {v_final:.2f} mL")

            # Subcase 1B: Acid Concentrating (Target Acid % > Current Acid %)
            else:
                op_label = f"ACID-ONLY CONCENTRATING ({curr_a}% ➔ {targ_a}% {acid_name})"
                if eff_stock_pct <= targ_a:
                    lbl_adj_summary.config(text=f"⚠️ Error: Stock acid strength ({st_pur}% @ {st_den} g/mL) is too low to reach {targ_a}%!")
                    return

                if mode == "adjust_vol" and not use_custom_v2:
                    v_stock_acid = v1 * (targ_a - curr_a) / (eff_stock_pct - targ_a)
                    v_final = v1 + v_stock_acid
                    lbl_adj_acid_title.config(text=f"🧪 Step 1 — Stock Acid to Add Directly (%{st_pur} {acid_name}):")
                    lbl_adj_acid_big.config(text=f"{v_stock_acid:.4f} mL   (New Final Volume: {v_final:.4f} mL)")
                    adj_acid_card.pack(fill="x", pady=(0, 8), before=adj_solute_card)
                    lbl_adj_solute_title.config(text=f"💧 Needed {solute} Volume:")
                    lbl_adj_solute_big.config(text="0.0000 mL   (Only stock acid is added)")
                else:
                    v_final = custom_v2 if use_custom_v2 else v1
                    v_kept = min(v1, v_final * (eff_stock_pct - targ_a) / (eff_stock_pct - curr_a))
                    v_withdraw = v1 - v_kept
                    pure_acid_deficit_g = (v_final * targ_a - v_kept * curr_a) / 100.0
                    v_stock_acid = (pure_acid_deficit_g / (st_pur / 100.0)) / st_den
                    solute_add = max(0.0, v_final - v_kept - v_stock_acid)

                    if v_withdraw > 1e-6:
                        lbl_adj_withdraw_title.config(text="📤 Step 1 — Withdraw From Current Acid Solution:")
                        lbl_adj_withdraw_big.config(text=f"Withdraw:  {v_withdraw:.4f} mL\n(Keep {v_kept:.4f} mL in the vessel)")
                        adj_withdraw_card.pack(fill="x", pady=(0, 8))

                    lbl_adj_acid_title.config(text=f"🧪 Step 2 — Add Stock Acid (%{st_pur} {acid_name}):")
                    lbl_adj_acid_big.config(text=f"Add {v_stock_acid:.4f} mL")
                    adj_acid_card.pack(fill="x", pady=(0, 8), before=adj_solute_card)

                    lbl_adj_solute_title.config(text=f"💧 Step 3 — Needed {solute} Volume:")
                    lbl_adj_solute_big.config(text=f"{solute_add:.4f} mL   (Complete to {v_final:.2f} mL)")

                lbl_adj_summary.config(text=f"Operation: {op_label}\nInitial Volume: {v1:.2f} mL   |   Final Volume: {v_final:.2f} mL")
                lbl_adj_liq_summary.config(text=f"Stock Acid: {acid_name} (Purity: %{st_pur}, Density: {st_den} g/mL)")
            return

        # ======================================================================
        # MODE 2: STANDARD ELEMENT + MATRIX SOLUTION ADJUSTMENT
        # ======================================================================
        c1_val = float(ent_adj_c1.get() or 0.0)
        u1 = cb_adj_u1.get()
        c2_val = float(ent_adj_c2.get() or 0.0)
        u2 = cb_adj_u2.get()

        c1_ppm = c1_val * unit_to_ppm_factor(u1)
        c2_ppm = c2_val * unit_to_ppm_factor(u2)

        if c1_ppm < 0 or c2_ppm <= 0:
            lbl_adj_summary.config(text="Please enter valid positive values for concentrations.")
            return

        is_dilution = c2_ppm < c1_ppm
        is_concentrating = c2_ppm > c1_ppm

        if is_concentrating:
            spike_frame.grid(row=7, column=0, columnspan=3, sticky="ew", pady=5)
            if var_spike_type.get() == "stock":
                lbl_spike_stock.grid(row=2, column=0, sticky="w", pady=3)
                ent_spike_stock.grid(row=2, column=1, sticky="ew", pady=3, padx=5)
                cb_spike_unit.grid(row=2, column=2, sticky="w", pady=3)
                if acid_name != "No Acid":
                    lbl_spike_acid.grid(row=3, column=0, sticky="w", pady=3)
                    ent_spike_acid.grid(row=3, column=1, sticky="ew", pady=3, padx=5)
                else:
                    lbl_spike_acid.grid_remove()
                    ent_spike_acid.grid_remove()
        else:
            spike_frame.grid_remove()

        v_withdraw = 0.0
        v_kept = v1
        v_final = v1
        v_stock_spike = 0.0
        added_elem_g = 0.0
        spike_acid_pct = 0.0

        # CASE A: DILUTION (C2 < C1)
        if is_dilution:
            op_label = "DILUTION (Lowering Element Concentration)"
            max_possible_v2 = (c1_ppm * v1) / c2_ppm

            if mode == "adjust_vol" and not use_custom_v2:
                v_final = max_possible_v2
                v_kept = v1
                lbl_adj_spike_title.config(text="🔹 Step 1 — Keep Full Initial Solution Volume:", bg="#eff6ff", fg="#1e40af")
                lbl_adj_spike_big.config(
                    text=f"Keep all {v1:.2f} mL   →   Expand to {v_final:.2f} mL\n(Total Volume to Add: {v_final - v1:.4f} mL)",
                    bg="#eff6ff", fg="#1d4ed8"
                )
                adj_spike_card.pack(fill="x", pady=(0, 8))
            else:
                v_final = custom_v2 if use_custom_v2 else v1
                v_needed_from_v1 = (c2_ppm * v_final) / c1_ppm

                if v_needed_from_v1 > v1 + 1e-6:
                    lbl_adj_summary.config(
                        text=f"⚠️ Warning: Your initial solution ({v1:g} mL @ {c1_val:g} {u1}) only contains enough element\n"
                             f"to make at most {max_possible_v2:.2f} mL at {c2_val:g} {u2} (without spiking)! Lower Custom V2 to ≤ {max_possible_v2:.2f} mL."
                    )
                    return

                v_kept = v_needed_from_v1
                v_withdraw = v1 - v_kept

                if v_withdraw > 1e-6:
                    lbl_adj_withdraw_title.config(text="📤 Step 1 — Take / Withdraw From Current Vessel:")
                    lbl_adj_withdraw_big.config(
                        text=f"Keep/Use:  {v_kept:.4f} mL   (Withdraw excess: {v_withdraw:.4f} mL)\nTarget Final Volume: {v_final:.2f} mL"
                    )
                    adj_withdraw_card.pack(fill="x", pady=(0, 8))
                else:
                    lbl_adj_spike_title.config(text="🔹 Step 1 — Keep Full Initial Solution Volume:", bg="#eff6ff", fg="#1e40af")
                    lbl_adj_spike_big.config(text=f"Keep all {v1:.2f} mL   →   Complete to {v_final:.2f} mL", bg="#eff6ff", fg="#1d4ed8")
                    adj_spike_card.pack(fill="x", pady=(0, 8))

        # CASE B: CONCENTRATING / SPIKING (C2 > C1)
        elif is_concentrating:
            op_label = "CONCENTRATING / SPIKING (Increasing Element Concentration)"
            spike_mode = var_spike_type.get()

            if spike_mode == "solid":
                v_final = custom_v2 if use_custom_v2 else v1
                v_kept = min(v1, v_final)
                v_withdraw = v1 - v_kept
                needed_mg = (c2_ppm * (v_final / 1000.0)) - (c1_ppm * (v_kept / 1000.0))
                added_elem_g = max(0.0, needed_mg / 1000.0)

                if v_withdraw > 1e-6:
                    lbl_adj_withdraw_title.config(text="📤 Step 1 — Adjust Initial Volume:")
                    lbl_adj_withdraw_big.config(text=f"Withdraw: {v_withdraw:.4f} mL  (Keep {v_kept:.4f} mL)")
                    adj_withdraw_card.pack(fill="x", pady=(0, 8))

                g_str = format_exact_decimal(added_elem_g, min_decimals=6)
                sub_str = format_smart_subunits(added_elem_g)

                lbl_adj_spike_title.config(text=f"🔹 Dissolve Pure {elem} Into Solution:", bg="#eff6ff", fg="#1e40af")
                lbl_adj_spike_big.config(text=f"{g_str} g\n({sub_str})", bg="#eff6ff", fg="#1d4ed8")
                adj_spike_card.pack(fill="x", pady=(0, 8))

                if var_adj_tare.get():
                    beaker_g = float(ent_adj_tare.get() or 0.0)
                    tot_g = beaker_g + added_elem_g
                    elem_dec = len(g_str.split(".")[1]) if "." in g_str else 6
                    mg_dec = max(4, elem_dec - 3)
                    lbl_adj_scale_big.config(text=f"{tot_g:.{elem_dec}f} g\n({tot_g*1000.0:.{mg_dec}f} mg)")
                    adj_scale_card.pack(fill="x", pady=(0, 8))

            else:
                cs_val = float(ent_spike_stock.get() or 0.0)
                cs_u = cb_spike_unit.get()
                cs_ppm = cs_val * unit_to_ppm_factor(cs_u)
                spike_acid_pct = float(ent_spike_acid.get() or 0.0) if acid_name != "No Acid" else 0.0

                if cs_ppm <= c2_ppm:
                    lbl_adj_summary.config(text=f"⚠️ Error: Stock concentration ({cs_val} {cs_u}) must be greater than Target C2 ({c2_val} {u2})!")
                    return

                if mode == "adjust_vol" and not use_custom_v2:
                    v_stock_spike = v1 * (c2_ppm - c1_ppm) / (cs_ppm - c2_ppm)
                    v_final = v1 + v_stock_spike
                    v_kept = v1

                    lbl_adj_spike_title.config(text=f"🔹 Step 1 — Add Concentrated Stock ({cs_val} {cs_u}):", bg="#eff6ff", fg="#1e40af")
                    lbl_adj_spike_big.config(
                        text=f"Add {v_stock_spike:.4f} mL Stock Solution\n(New Final Volume: {v_final:.4f} mL)",
                        bg="#eff6ff", fg="#1d4ed8"
                    )
                    adj_spike_card.pack(fill="x", pady=(0, 8))
                else:
                    v_final = custom_v2 if use_custom_v2 else v1
                    # Keep as much of V1 as possible without exceeding V_final after adding stock
                    max_v_kept = v_final * (cs_ppm - c2_ppm) / (cs_ppm - c1_ppm)
                    v_kept = min(v1, max_v_kept)
                    v_withdraw = v1 - v_kept
                    v_stock_spike = (c2_ppm * v_final - c1_ppm * v_kept) / cs_ppm

                    if v_withdraw > 1e-6:
                        lbl_adj_withdraw_title.config(text="📤 Step 1 — Withdraw From Current Solution:")
                        lbl_adj_withdraw_big.config(text=f"Withdraw:  {v_withdraw:.4f} mL\n(Remaining Volume: {v_kept:.4f} mL)")
                        adj_withdraw_card.pack(fill="x", pady=(0, 8))

                    lbl_adj_spike_title.config(text=f"🔹 Step 2 — Add Concentrated Stock ({cs_val} {cs_u}):", bg="#eff6ff", fg="#1e40af")
                    lbl_adj_spike_big.config(
                        text=f"Add {v_stock_spike:.4f} mL Stock Solution\n(Target Final Volume: {v_final:.2f} mL)",
                        bg="#eff6ff", fg="#1d4ed8"
                    )
                    adj_spike_card.pack(fill="x", pady=(0, 8))

        else:
            op_label = "MATRIX / ACID ADJUSTMENT ONLY (C1 = C2)"
            v_final = custom_v2 if use_custom_v2 else v1
            v_kept = min(v1, v_final)

        summary_str = (
            f"Operation: {op_label}\n"
            f"Element: {elem} (A.M: {at_mass} g/mol)   |   {c1_val} {u1}  ➔  {c2_val} {u2}\n"
            f"Initial Volume: {v1:.2f} mL   |   Final Volume: {v_final:.2f} mL"
        )
        lbl_adj_summary.config(text=summary_str)

        # ACID MATRIX & SOLVENT BALANCING
        volume_to_fill = v_final - v_kept - v_stock_spike

        if acid_name != "No Acid":
            curr_a_pct = float(ent_adj_curr_acid.get() or 0.0)
            targ_a_pct = float(ent_adj_targ_acid.get() or 0.0)
            st_pur = float(ent_adj_acid_pur.get() or 1.0)
            st_den = float(ent_adj_acid_den.get() or 1.0)

            pure_acid_present_g = v_kept * (curr_a_pct / 100.0) + v_stock_spike * (spike_acid_pct / 100.0)
            pure_acid_target_g = v_final * (targ_a_pct / 100.0)
            pure_acid_deficit_g = pure_acid_target_g - pure_acid_present_g

            lbl_adj_liq_summary.config(text=f"Main Solute: {solute}   |   Acid Matrix: {curr_a_pct}% ➔ {targ_a_pct}% {acid_name}")

            if pure_acid_deficit_g >= -1e-6:
                pure_acid_deficit_g = max(0.0, pure_acid_deficit_g)
                stock_acid_ml = (pure_acid_deficit_g / (st_pur / 100.0)) / st_den
                solute_needed_ml = volume_to_fill - stock_acid_ml

                lbl_adj_acid_title.config(text=f"🧪 Stock Acid to Add (%{st_pur} {acid_name}):")
                lbl_adj_acid_big.config(text=f"{stock_acid_ml:.4f} mL")
                adj_acid_card.pack(fill="x", pady=(0, 8), before=adj_solute_card)

                lbl_adj_solute_title.config(text=f"💧 Needed {solute} Volume:")
                if solute_needed_ml >= 0:
                    lbl_adj_solute_big.config(text=f"Approx. {solute_needed_ml:.4f} mL  (Completes to {v_final:.2f} mL)")
                else:
                    lbl_adj_solute_big.config(
                        text=f"0.00 mL (⚠️ Adding {stock_acid_ml:.2f} mL acid slightly exceeds {v_final:.2f} mL fixed volume)"
                    )
            else:
                lbl_adj_acid_title.config(text=f"⚠️ Acid Matrix Notice ({acid_name}):")
                lbl_adj_acid_big.config(
                    text=f"Current acid in kept solution exceeds target {targ_a_pct}%!\n"
                         f"Dilute further or lower initial acid percentage."
                )
                adj_acid_card.pack(fill="x", pady=(0, 8), before=adj_solute_card)
                lbl_adj_solute_title.config(text=f"💧 Needed {solute} Volume:")
                lbl_adj_solute_big.config(text=f"Approx. {max(0.0, volume_to_fill):.4f} mL")
        else:
            lbl_adj_liq_summary.config(text=f"Main Solute: {solute}   |   Matrix Acid: No Acid")
            lbl_adj_solute_title.config(text=f"💧 Needed {solute} Volume to Add:")
            if volume_to_fill > 1e-6:
                lbl_adj_solute_big.config(text=f"{volume_to_fill:.4f} mL   (Complete to {v_final:.2f} mL)")
            else:
                lbl_adj_solute_big.config(text=f"0.0000 mL   (Final volume remains {v_final:.2f} mL)")

    except Exception as ex:
        print("Adjustment Calc Error:", ex)


def update_main_dropdowns():
    elem_keys = list(db["elements"].keys())
    for cb in (cb_element, cb_adj_element):
        cb["values"] = elem_keys
        if cb.get() not in elem_keys and elem_keys:
            cb.current(0)

    for cb in (cb_solute, cb_adj_solute):
        cb["values"] = db["solutes"]
        if cb.get() not in db["solutes"] and db["solutes"]:
            cb.current(0)

    acid_keys = ["No Acid"] + list(db["acids"].keys())
    if cb_acid.get() not in acid_keys:
        cb_acid.current(0)
        on_acid_change()
    cb_acid["values"] = acid_keys

    if cb_adj_acid.get() not in acid_keys:
        cb_adj_acid.current(0)
        on_adj_acid_change()
    cb_adj_acid["values"] = acid_keys

    for r in alloy_rows:
        r["cb_elem"]["values"] = elem_keys
        if r["cb_elem"].get() not in elem_keys and elem_keys:
            r["cb_elem"].current(0)

    calculate()
    calculate_alloy()
    calculate_adjustment()

# ==============================================================================
# EXPORT & RESET FUNCTIONS (SUPPORTS ALL 3 TABS)
# ==============================================================================
def reset_parameters():
    current_tab = main_notebook.index(main_notebook.select())
    if current_tab == 0:
        var_tab1_no_elem.set(False)
        ent_conc.delete(0, tk.END)
        ent_conc.insert(0, "114.0")
        ent_vol.delete(0, tk.END)
        ent_vol.insert(0, "1000.0")
        cb_acid.current(0)
        var_beaker.set(False)
        on_tab1_no_elem_toggle()
        on_acid_change()
        on_beaker_toggle()
        calculate()
    elif current_tab == 1:
        ent_alloy_total.delete(0, tk.END)
        ent_alloy_total.insert(0, "10.0")
        cb_alloy_unit.current(0)
        var_alloy_tare.set(False)
        on_alloy_tare_toggle()
        clear_alloy_actual_weights()
    else:
        var_adj_mode.set("adjust_vol")
        var_adj_no_elem.set(False)
        ent_adj_v1.delete(0, tk.END)
        ent_adj_v1.insert(0, "500.0")
        ent_adj_c1.delete(0, tk.END)
        ent_adj_c1.insert(0, "100.0")
        cb_adj_u1.current(0)
        ent_adj_c2.delete(0, tk.END)
        ent_adj_c2.insert(0, "25.0")
        cb_adj_u2.current(0)
        cb_adj_acid.current(0)
        var_adj_tare.set(False)
        on_adj_no_elem_toggle()
        on_adj_acid_change()
        on_adj_tare_toggle()
        calculate_adjustment()

def get_current_recipe_data():
    current_tab = main_notebook.index(main_notebook.select())
    if current_tab == 0:
        no_elem = var_tab1_no_elem.get()
        chosen_element = cb_element.get()
        atkes = db["elements"].get(chosen_element, 0.0)
        conc_val = float(ent_conc.get() or 0.0)
        conc_unit = cb_unit.get()
        final_vol = float(ent_vol.get() or 0.0)
        solute = cb_solute.get()
        acid = cb_acid.get()

        data = {
            "Mode": "Acid-Only Blank Preparation" if no_elem else "Standard Solution Preparation",
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Final Volume": f"{final_vol} mL",
            "Main Solute": solute
        }

        if not no_elem:
            factor = unit_to_ppm_factor(conc_unit)
            needed_mg = (conc_val * factor) * (final_vol / 1000.0)
            needed_g = needed_mg / 1000.0
            g_str = format_exact_decimal(needed_g, min_decimals=6)
            sub_str = format_smart_subunits(needed_g)

            data["Target Element"] = f"{chosen_element} (A.M: {atkes} g/mol)"
            data["Target Concentration"] = f"{conc_val} {conc_unit}"
            data["Needed Pure Element Mass"] = f"{g_str} g ({sub_str})"

            if var_beaker.get():
                beaker_m = float(ent_beaker.get() or 0.0)
                total_g = beaker_m + needed_g
                elem_decimals = len(g_str.split(".")[1]) if "." in g_str else 6
                mg_decimals = max(4, elem_decimals - 3)
                data["Beaker Mass (Tare)"] = f"{beaker_m:.4f} g ({beaker_m*1000:.2f} mg)"
                data["Scale Screen (Beaker + Element)"] = f"{total_g:.{elem_decimals}f} g ({total_g*1000:.{mg_decimals}f} mg)"

        if acid != "No Acid":
            a_pct = float(ent_acid_percent.get() or 0.0)
            a_pur = float(ent_acid_purity.get() or 1.0)
            a_den = float(ent_acid_density.get() or 1.0)
            pure_a_g = final_vol * (a_pct / 100.0)
            stock_a_ml = (pure_a_g / (a_pur / 100.0)) / a_den
            solute_left = final_vol - stock_a_ml

            data["Matrix Acid"] = f"{acid} (Target Matrix: %{a_pct}, Stock Purity: %{a_pur})"
            data["Stock Acid Volume to Add"] = f"{stock_a_ml:.4f} mL"
            data[f"Needed {solute} Volume"] = f"Approx. {solute_left:.4f} mL"
        else:
            data["Matrix Acid"] = "No Acid"
            data[f"Needed {solute} Volume"] = f"Complete to {final_vol} mL"

        return data

    elif current_tab == 1:
        total_val = float(ent_alloy_total.get() or 0.0)
        total_unit = cb_alloy_unit.get()
        desired_mg = total_val * 1000.0 if total_unit == "g" else total_val

        data = {
            "Mode": "Alloy Preparation",
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Desired Total Alloy Weight": f"{desired_mg:.4f} mg ({desired_mg/1000.0:.6f} g)"
        }
        actual_masses = []
        for idx, r in enumerate(alloy_rows, 1):
            el = r["cb_elem"].get()
            pct = float(r["ent_pct"].get() or 0.0)
            theo_mg = desired_mg * (pct / 100.0)
            act_str = r["ent_act"].get().strip()
            act_u = r["cb_act_unit"].get()
            if act_str:
                act_val = float(act_str)
                m_mg = act_val * 1000.0 if act_u == "g" else act_val
                status = f"Actual: {m_mg:.4f} mg (Target: {pct:.2f}% / {theo_mg:.4f} mg)"
            else:
                m_mg = theo_mg
                status = f"Target: {theo_mg:.4f} mg ({pct:.2f}%)"
            data[f"Element #{idx} - {el}"] = status
            actual_masses.append((el, m_mg))

        tot_actual = sum(m for _, m in actual_masses)
        comp_parts = [f"{el}: {(m/tot_actual*100.0 if tot_actual else 0):.2f} wt%" for el, m in actual_masses]
        data["Resulting Total Mass"] = f"{tot_actual:.4f} mg ({tot_actual/1000.0:.6f} g)"
        data["Resulting Composition"] = " | ".join(comp_parts)
        return data

    else:
        no_elem = var_adj_no_elem.get()
        data = {
            "Mode": "Acid-Only Dilution & Adjustment" if no_elem else "Solution Dilution & Adjustment",
            "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Volume Handling Mode": "Adjust Final Volume" if var_adj_mode.get() == "adjust_vol" else "Keep Volume Constant",
        }
        if no_elem:
            data["Acid"] = cb_adj_acid.get()
            data["Initial Acid Solution"] = f"{ent_adj_v1.get()} mL @ {ent_adj_curr_acid.get()}%"
            data["Target Acid Percentage"] = f"{ent_adj_targ_acid.get()}%"
        else:
            data["Element"] = cb_adj_element.get()
            data["Initial Solution"] = f"{ent_adj_v1.get()} mL @ {ent_adj_c1.get()} {cb_adj_u1.get()}"
            data["Target Concentration"] = f"{ent_adj_c2.get()} {cb_adj_u2.get()}"

        if adj_withdraw_card.winfo_ismapped():
            data["Volume to Withdraw"] = lbl_adj_withdraw_big.cget("text").replace("\n", " | ")
        if adj_spike_card.winfo_ismapped():
            data["Element / Volume Step"] = lbl_adj_spike_big.cget("text").replace("\n", " | ")
        if adj_acid_card.winfo_ismapped():
            data["Stock Acid to Add"] = lbl_adj_acid_big.cget("text").replace("\n", " | ")
        data[f"Needed {cb_adj_solute.get()} Volume"] = lbl_adj_solute_big.cget("text").replace("\n", " | ")
        return data

def export_to_word():
    filepath = filedialog.asksaveasfilename(
        defaultextension=".docx",
        filetypes=[("Word Document", "*.docx")],
        initialfile="DiluCalc_Recipe.docx",
        title="Export Recipe to Word"
    )
    if not filepath:
        return

    data = get_current_recipe_data()
    doc = Document()
    doc.add_heading(f"🧪 DiluCalc v3 - {data.get('Mode', 'Preparation Recipe')}", level=1)
    doc.add_paragraph(f"Generated on: {data['Date']}")

    table = doc.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Parameter / Component"
    hdr_cells[1].text = "Required / Measured Value"

    for k, v in data.items():
        if k in ("Date", "Mode"):
            continue
        row_cells = table.add_row().cells
        row_cells[0].text = str(k)
        row_cells[1].text = str(v)

    doc.save(filepath)
    messagebox.showinfo("Export Successful", f"Recipe successfully exported to Word:\n{filepath}")

def export_to_excel():
    filepath = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile="DiluCalc_Recipe.xlsx",
        title="Export Recipe to Excel"
    )
    if not filepath:
        return

    data = get_current_recipe_data()
    wb = Workbook()
    ws = wb.active
    ws.title = "Preparation Recipe"

    ws["A1"] = f"DiluCalc v3 - {data.get('Mode', 'Laboratory Recipe')}"
    ws["A1"].font = Font(size=13, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws.merge_cells("A1:B1")

    ws.append(["Parameter / Component", "Required / Measured Value"])
    ws["A2"].font = Font(bold=True)
    ws["B2"].font = Font(bold=True)

    for k, v in data.items():
        ws.append([k, v])

    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 55

    wb.save(filepath)
    messagebox.showinfo("Export Successful", f"Recipe successfully exported to Excel:\n{filepath}")

def export_to_png():
    filepath = filedialog.asksaveasfilename(
        defaultextension=".png",
        filetypes=[("PNG Image", "*.png")],
        initialfile="DiluCalc_Recipe.png",
        title="Export Recipe to PNG"
    )
    if not filepath:
        return

    data = get_current_recipe_data()
    width = 780
    height = max(480, 140 + len(data) * 38)
    img = Image.new("RGB", (width, height), color="#ffffff")
    draw = ImageDraw.Draw(img)

    try:
        font_title = ImageFont.truetype("segoeui.ttf", 19)
        font_bold = ImageFont.truetype("segoeuib.ttf", 13)
        font_reg = ImageFont.truetype("segoeui.ttf", 13)
    except:
        font_title = font_bold = font_reg = ImageFont.load_default()

    draw.rectangle([(0, 0), (width, 60)], fill="#1e293b")
    draw.text((20, 18), f"DiluCalc v3 - {data.get('Mode', 'Recipe')}", fill="#ffffff", font=font_title)

    y = 80
    for k, v in data.items():
        if any(w in k for w in ("Needed Pure", "Scale Screen", "Stock Acid", "Resulting", "Withdraw", "Element / Volume")):
            draw.rectangle([(15, y - 4), (width - 15, y + 26)], fill="#eff6ff", outline="#bfdbfe")
        draw.text((25, y), f"{k}:", fill="#334155", font=font_bold)
        draw.text((290, y), f"{v}", fill="#0f172a", font=font_reg)
        y += 36

    img.save(filepath)
    messagebox.showinfo("Export Successful", f"Recipe card exported to PNG:\n{filepath}")

# MENU BAR SETUP
menubar = tk.Menu(root)

file_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10))
file_menu.add_command(label="🔄 Reset Active Tab", command=reset_parameters)

export_menu = tk.Menu(file_menu, tearoff=0, font=("Segoe UI", 10))
export_menu.add_command(label="📄 Export as Word (.docx)", command=export_to_word)
export_menu.add_command(label="📊 Export as Excel (.xlsx)", command=export_to_excel)
export_menu.add_command(label="🖼️ Export as Image (.png)", command=export_to_png)
file_menu.add_cascade(label="📤 Export Active Recipe As...", menu=export_menu)

file_menu.add_separator()
file_menu.add_command(label="🚪 Exit", command=root.quit)
menubar.add_cascade(label="File", menu=file_menu)

data_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10))
data_menu.add_command(label="📚 See Library", command=open_library_window)
menubar.add_cascade(label="Data", menu=data_menu)

help_menu = tk.Menu(menubar, tearoff=0, font=("Segoe UI", 10))
help_menu.add_command(label="📖 How to Use", command=open_help_window)
help_menu.add_command(label="ℹ️ About DiluCalc", command=open_about_window)
menubar.add_cascade(label="Help", menu=help_menu)

root.config(menu=menubar)

# BINDINGS & INITIALIZATION (TAB 1)
chk_tab1_no_elem.config(command=on_tab1_no_elem_toggle)
cb_element.bind("<<ComboboxSelected>>", calculate)
cb_unit.bind("<<ComboboxSelected>>", calculate)
cb_solute.bind("<<ComboboxSelected>>", calculate)
cb_acid.bind("<<ComboboxSelected>>", on_acid_change)
chk_beaker.config(command=on_beaker_toggle)

for entry_widget in [ent_conc, ent_vol, ent_acid_percent, ent_acid_purity, ent_acid_density, ent_beaker]:
    entry_widget.bind("<KeyRelease>", calculate)

# BINDINGS & INITIALIZATION (TAB 3)
chk_adj_no_elem.config(command=on_adj_no_elem_toggle)
rb_adj_vol.config(command=on_adj_mode_or_v2_toggle)
rb_keep_vol.config(command=on_adj_mode_or_v2_toggle)
chk_adj_custom_v2.config(command=on_adj_mode_or_v2_toggle)
ent_adj_v2.bind("<KeyRelease>", calculate_adjustment)
rb_spike_solid.config(command=on_spike_type_change)
rb_spike_stock.config(command=on_spike_type_change)
chk_adj_tare.config(command=on_adj_tare_toggle)

for cb_w in [cb_adj_element, cb_adj_solute, cb_adj_u1, cb_adj_u2, cb_spike_unit]:
    cb_w.bind("<<ComboboxSelected>>", calculate_adjustment)
cb_adj_acid.bind("<<ComboboxSelected>>", on_adj_acid_change)

for ent_w in [
    ent_adj_v1, ent_adj_c1, ent_adj_c2, ent_spike_stock, ent_spike_acid,
    ent_adj_curr_acid, ent_adj_targ_acid, ent_adj_acid_pur, ent_adj_acid_den, ent_adj_tare
]:
    ent_w.bind("<KeyRelease>", calculate_adjustment)

# Initialize 2 default element rows in the Alloy tab (40% and 60% Balance)
add_alloy_element_row(default_elem_idx=0, default_pct="40.0", is_balance=False)
add_alloy_element_row(default_elem_idx=1, default_pct="60.0", is_balance=True)

on_acid_change()
on_beaker_toggle()
on_alloy_tare_toggle()
on_adj_acid_change()
on_spike_type_change()
on_adj_tare_toggle()
calculate()
calculate_alloy()
calculate_adjustment()

root.mainloop()