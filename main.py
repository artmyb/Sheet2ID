import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def get_unicode_font():
    font_paths = [
        "C:/Windows/Fonts/arial.ttf", 
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/Library/Fonts/Arial.ttf"
    ]
    for path in font_paths:
        if os.path.exists(path):
            return path
    return None

class IDCardGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Sheet2ID")
        self.root.geometry("1400x900")

        icon_data = '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'
        self.icon_image = tk.PhotoImage(data=icon_data)
        self.root.wm_iconphoto(True, self.icon_image)
        self.df = None
        self.photo_dir = ""
        self.bg_path = ""
        self.elements = {} 
        self.photo_settings = {}
        self.current_index = 0
        self.font_path = get_unicode_font()
        if self.font_path:
            pdfmetrics.registerFont(TTFont('UnicodeFont', self.font_path))
            self.font_name = 'UnicodeFont'
        else:
            self.font_name = 'Helvetica'
        
        self.setup_ui()

    def setup_ui(self):
        left_panel = ttk.Frame(self.root, padding="10")
        left_panel.pack(side=tk.LEFT, fill=tk.Y)

        self.load_data_button = ttk.Button(left_panel, text="📁 Load Spreadsheet", command=self.load_data)
        self.load_data_button.pack(fill=tk.X, pady=2)
        self.set_photo_dir_button = ttk.Button(left_panel, text="📸 Select Photo Folder", command=self.set_photo_dir)
        self.set_photo_dir_button.pack(fill=tk.X, pady=2)
        self.select_bg_button = ttk.Button(left_panel, text="🖼️ Select Background", command=self.set_background)
        self.select_bg_button.pack(fill=tk.X, pady=2)

        dim_frame = ttk.LabelFrame(left_panel, text="Card & Photo Dimensions (mm)")
        dim_frame.pack(fill=tk.X, pady=5)
        
        self.card_w = tk.DoubleVar(value=90)
        self.card_h = tk.DoubleVar(value=120)
        self.img_w = tk.DoubleVar(value=25.0)
        self.img_h = tk.DoubleVar(value=30.0)

        params = [("Card W", self.card_w), ("Card H", self.card_h), ("Photo W", self.img_w), ("Photo H", self.img_h)]
        for i, (txt, var) in enumerate(params):
            ttk.Label(dim_frame, text=txt).grid(row=i//2, column=(i%2)*2, sticky="w", padx=2)
            ent = ttk.Entry(dim_frame, textvariable=var, width=7)
            ent.grid(row=i//2, column=(i%2)*2+1, pady=2)
            self.bind_scroll(ent, var)
            var.trace_add("write", lambda *a: self.update_preview())

        mapping_container = ttk.Frame(left_panel)
        mapping_container.pack(fill=tk.BOTH, expand=True)
        
        self.el_canvas = tk.Canvas(mapping_container, width=650)
        self.el_vars_container = ttk.Frame(self.el_canvas)
        scrollbar = ttk.Scrollbar(mapping_container, orient="vertical", command=self.el_canvas.yview)
        self.el_canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.el_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.el_canvas.create_window((0,0), window=self.el_vars_container, anchor="nw")

        preview_panel = ttk.Frame(self.root, padding="10")
        preview_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self.canvas_preview = tk.Canvas(preview_panel, bg="#333")
        self.canvas_preview.pack(fill=tk.BOTH, expand=True)

        nav_frame = ttk.Frame(preview_panel)
        nav_frame.pack(fill=tk.X, pady=5)
        ttk.Button(nav_frame, text="◀ Previous", command=lambda: self.navigate(-1)).pack(side=tk.LEFT)
        self.nav_label = ttk.Label(nav_frame, text="0 / 0", font=("Arial", 10, "bold"))
        self.nav_label.pack(side=tk.LEFT, padx=20)
        ttk.Button(nav_frame, text="Next ▶", command=lambda: self.navigate(1)).pack(side=tk.LEFT)
        ttk.Button(preview_panel, text="GENERATE PDF", command=self.export_pdf).pack(side=tk.RIGHT)

    def bind_scroll(self, widget, var):
        def on_mousewheel(event):
            step = 0.5 if (event.delta > 0 or event.num == 4) else -0.5
            try:
                var.set(round(float(var.get()) + step, 1))
            except: pass
        widget.bind("<MouseWheel>", on_mousewheel)
        widget.bind("<Button-4>", on_mousewheel)
        widget.bind("<Button-5>", on_mousewheel)

    def load_data(self):
        path = filedialog.askopenfilename(filetypes=[("Data", "*.xlsx *.csv")])
        if not path: return
        self.load_data_button.config(text = '📁 '+path.split('/')[-1])
        self.df = pd.read_csv(path) if path.endswith('.csv') else pd.read_excel(path)
        self.refresh_element_list()
        self.update_preview()

    def refresh_element_list(self):
        for widget in self.el_vars_container.winfo_children(): widget.destroy()
        
        p_fr = ttk.Frame(self.el_vars_container); p_fr.pack(fill=tk.X, pady=10)
        p_act, px, py = tk.BooleanVar(value=True), tk.DoubleVar(value=45), tk.DoubleVar(value=45)
        ax, ay = tk.StringVar(value="center"), tk.StringVar(value="top")

        ttk.Checkbutton(p_fr, variable=p_act).grid(row=0, column=0)
        ttk.Label(p_fr, text="[ID PHOTO]", font=("Arial", 9, "bold"), width=12).grid(row=0, column=1)
        ttk.Label(p_fr, text="X").grid(row=0, column=2); ex = ttk.Entry(p_fr, textvariable=px, width=5); ex.grid(row=0, column=3)
        ttk.Label(p_fr, text="Y").grid(row=0, column=4); ey = ttk.Entry(p_fr, textvariable=py, width=5); ey.grid(row=0, column=5)
        ttk.Combobox(p_fr, textvariable=ax, values=["left", "center", "right"], width=7).grid(row=0, column=6)
        ttk.Combobox(p_fr, textvariable=ay, values=["top", "center", "bottom"], width=7).grid(row=0, column=7)

        self.bind_scroll(ex, px); self.bind_scroll(ey, py)
        self.photo_settings = {"active": p_act, "x": px, "y": py, "anch_x": ax, "anch_y": ay}
        for v in [p_act, px, py, ax, ay]: v.trace_add("write", lambda *a: self.update_preview())

        ttk.Separator(self.el_vars_container, orient="horizontal").pack(fill=tk.X, pady=5)

        self.elements = {}
        for col in self.df.columns:
            fr = ttk.Frame(self.el_vars_container); fr.pack(fill=tk.X, pady=1)
            act, spec = tk.BooleanVar(), tk.BooleanVar()
            x, y, fz = tk.DoubleVar(value=45), tk.DoubleVar(value=10), tk.IntVar(value=12)
            anch = tk.StringVar(value="center")

            ttk.Checkbutton(fr, variable=act).grid(row=0, column=0)
            ttk.Label(fr, text=str(col)[:15], width=15).grid(row=0, column=1)
            ttk.Checkbutton(fr, text="Photo Specifier", variable=spec).grid(row=0, column=2)
            ttk.Label(fr, text="  X:").grid(row=0, column=3)
            ex = ttk.Entry(fr, textvariable=x, width=4); ex.grid(row=0, column=4)
            ttk.Label(fr, text="  Y:").grid(row=0, column=5)
            ey = ttk.Entry(fr, textvariable=y, width=4); ey.grid(row=0, column=6)
            ttk.Label(fr, text="  Size:").grid(row=0, column=7)
            ef = ttk.Entry(fr, textvariable=fz, width=3); ef.grid(row=0, column=8)
            ttk.Label(fr, text="  Anchor:").grid(row=0, column=9)
            cb = ttk.Combobox(fr, textvariable=anch, values=["left", "center", "right"], width=10); cb.grid(row=0, column=10)

            self.bind_scroll(ex, x); self.bind_scroll(ey, y); self.bind_scroll(ef, fz)
            for v in [act, spec, x, y, fz, anch]: v.trace_add("write", lambda *a: self.update_preview())
            self.elements[col] = {"active": act, "specifier": spec, "x": x, "y": y, "font_size": fz, "anchor": anch}

    def find_photo(self, idx):
        if not self.photo_dir: return None
        search_terms = [str(self.df.iloc[idx][c]).lower() for c, s in self.elements.items() if s["specifier"].get()]
        if not search_terms: return None
        try:
            for f in os.listdir(self.photo_dir):
                if all(t in f.lower() for t in search_terms): return os.path.join(self.photo_dir, f)
        except: pass
        return None

    def update_preview(self):
        if self.df is None: return
        sc = 4.0
        pw, ph = int(self.card_w.get()*sc), int(self.card_h.get()*sc)
        img = Image.new('RGB', (pw, ph), color='white')
        if self.bg_path and os.path.exists(self.bg_path):
            try: img.paste(Image.open(self.bg_path).resize((pw, ph)), (0,0))
            except: pass
            
        draw = ImageDraw.Draw(img)
        row_data = self.df.iloc[self.current_index]

        if self.photo_settings["active"].get():
            p_path = self.find_photo(self.current_index)
            if p_path:
                try:
                    iw, ih = int(self.img_w.get()*sc), int(self.img_h.get()*sc)
                    tx, ty = self.photo_settings["x"].get()*sc, self.photo_settings["y"].get()*sc
                    if self.photo_settings["anch_x"].get() == "center": tx -= iw/2
                    elif self.photo_settings["anch_x"].get() == "right": tx -= iw
                    if self.photo_settings["anch_y"].get() == "center": ty -= ih/2
                    elif self.photo_settings["anch_y"].get() == "bottom": ty -= ih
                    img.paste(Image.open(p_path).resize((iw, ih), Image.Resampling.LANCZOS), (int(tx), int(ty)))
                except: pass

        for col, s in self.elements.items():
            if s["active"].get():
                try: 
                    p_font = ImageFont.truetype(self.font_path, int(s["font_size"].get()*sc/2.5)) if self.font_path else ImageFont.load_default()
                except: p_font = ImageFont.load_default()
                
                txt = str(row_data[col])
                bbox = draw.textbbox((0,0), txt, font=p_font)
                tw = bbox[2] - bbox[0]
                tx, ty = s["x"].get()*sc, s["y"].get()*sc
                if s["anchor"].get() == "center": tx -= tw/2
                elif s["anchor"].get() == "right": tx -= tw
                draw.text((tx, ty), txt, fill="black", font=p_font)

        self.tk_img = ImageTk.PhotoImage(img)
        self.canvas_preview.delete("all")
        self.canvas_preview.create_image(self.canvas_preview.winfo_width()//2, self.canvas_preview.winfo_height()//2, image=self.tk_img)
        self.nav_label.config(text=f"Record {self.current_index + 1} of {len(self.df)}")

    def set_photo_dir(self):
        self.photo_dir = filedialog.askdirectory()
        if self.photo_dir:
            self.set_photo_dir_button.config(text = "📸 "+self.photo_dir.split('/')[-1])
        self.update_preview()
    def set_background(self):
        self.bg_path = filedialog.askopenfilename()
        if self.bg_path:
            self.select_bg_button.config(text = "🖼️ "+self.bg_path.split('/')[-1])
        self.update_preview()
    def navigate(self, step): 
        if self.df is not None: self.current_index = (self.current_index + step) % len(self.df); self.update_preview()

    def export_pdf(self):
        out = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not out: return
        c = canvas.Canvas(out, pagesize=A4)
        pw, ph = A4
        cw, ch = self.card_w.get()*mm, self.card_h.get()*mm
        curr_x, curr_y = 10*mm, ph - ch - 10*mm

        for i in range(len(self.df)):
            if self.bg_path: 
                try: c.drawImage(self.bg_path, curr_x, curr_y, cw, ch)
                except: pass
            c.setDash(1, 2); c.rect(curr_x, curr_y, cw, ch); c.setDash([])
            
            if self.photo_settings["active"].get():
                p = self.find_photo(i)
                if p:
                    try:
                        iw, ih = self.img_w.get()*mm, self.img_h.get()*mm
                        tx, ty = self.photo_settings["x"].get()*mm, self.photo_settings["y"].get()*mm
                        if self.photo_settings["anch_x"].get() == "center": tx -= iw/2
                        elif self.photo_settings["anch_x"].get() == "right": tx -= iw
                        if self.photo_settings["anch_y"].get() == "center": ty -= ih/2
                        elif self.photo_settings["anch_y"].get() == "bottom": ty -= ih
                        c.drawImage(p, curr_x + tx, curr_y + (ch - ty - ih), iw, ih)
                    except: pass

            for col, s in self.elements.items():
                if s["active"].get():
                    c.setFont(self.font_name, s["font_size"].get())
                    txt = str(self.df.iloc[i][col])
                    tw = c.stringWidth(txt, self.font_name, s["font_size"].get())
                    tx, ty = s["x"].get()*mm, s["y"].get()*mm
                    final_x = curr_x + tx
                    if s["anchor"].get() == "center": final_x -= tw/2
                    elif s["anchor"].get() == "right": final_x -= tw
                    c.drawString(final_x, curr_y + (ch - ty - s["font_size"].get()*0.8), txt)

            curr_x += cw + 5*mm
            if curr_x + cw > pw - 10*mm: curr_x = 10*mm; curr_y -= (ch + 5*mm)
            if curr_y < 10*mm: c.showPage(); curr_y = ph - ch - 10*mm
            
        c.save()
        messagebox.showinfo("Done", f"Exported {len(self.df)} cards.")

if __name__ == "__main__":
    root = tk.Tk(); app = IDCardGenerator(root); root.mainloop()