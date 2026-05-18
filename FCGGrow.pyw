"""
FCG Grow v1 - AI-Powered Business Operations Board
==================================================
Built by Fairy Circle Garden | fairycirclegarden.com
Free and open source for every entrepreneur.

Requirements:
    pip install anthropic openai pillow requests icalendar tkinterdnd2 pdfplumber python-docx

Supports: Anthropic, OpenAI, OpenRouter, Groq, Gemini, Grok, Cohere, Ollama, NVIDIA NIM, Custom
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import json, os, sys, threading, urllib.request, shutil, re
from datetime import datetime

# ── EXE-safe base directory ────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

WORKSPACES_DIR = os.path.join(BASE_DIR, "fcggrow_workspaces")
SESSION_FILE   = os.path.join(BASE_DIR, "fcggrow_session.json")
os.makedirs(WORKSPACES_DIR, exist_ok=True)

# ── Optional deps ──────────────────────────────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_ALL
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

try:
    from icalendar import Calendar as ICal
    ICAL_AVAILABLE = True
except ImportError:
    ICAL_AVAILABLE = False

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

try:
    from openai import OpenAI
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    from PIL import Image, ImageDraw, ImageTk, ImageFilter
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

try:
    import docx as python_docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# ── Dark ttk style (applied once at startup) ──────────────────────────────
def apply_dark_ttk_style(root):
    """Apply dark theme. Must be called after Tk root exists."""
    try:
        # Combobox dropdown list popup (tk.Listbox inside ttk) needs option_add
        root.option_add("*TCombobox*Listbox.background", "#0d2618")
        root.option_add("*TCombobox*Listbox.foreground", "#ccddcc")
        root.option_add("*TCombobox*Listbox.selectBackground", "#2d6a4f")
        root.option_add("*TCombobox*Listbox.selectForeground", "white")
        root.option_add("*TCombobox*Listbox.font", "Arial 9")
        s = ttk.Style(root)
        s.theme_use("clam")
        s.configure("Dark.Vertical.TScrollbar",
                    troughcolor="#071510", background="#2d6a4f",
                    arrowcolor="#52b788", bordercolor="#071510",
                    lightcolor="#2d6a4f", darkcolor="#0d2618",
                    gripcount=0, arrowsize=12, relief="flat", width=12)
        s.map("Dark.Vertical.TScrollbar",
              background=[("active", "#52b788"), ("pressed", "#52b788"),
                          ("!active", "#1a3a2a")])
        s.configure("TCombobox", fieldbackground="#0d2618",
                    background="#1a3a2a", foreground="#ccddcc",
                    arrowcolor="#52b788", selectbackground="#2d6a4f",
                    selectforeground="#ccddcc", bordercolor="#1a4a2a")
        s.map("TCombobox",
              fieldbackground=[("readonly", "#0d2618"), ("!readonly", "#0d2618")],
              foreground=[("readonly", "#ccddcc"), ("!readonly", "#ccddcc")],
              background=[("readonly", "#1a3a2a")])
    except Exception:
        pass


# ── Board constants ────────────────────────────────────────────────────────
COL_LABELS = ["To Do", "In Progress", "Review", "Done"]
COL_KEYS   = ["todo", "inprog", "review", "done"]
PRIORITIES = ["High", "Medium", "Low"]
CATEGORIES = ["Grant", "Operations", "Admin", "Outreach", "Marketing", "Finance", "Other"]

# Grant keywords for auto-categorization
GRANT_KEYWORDS = [
    "grant", "award", "foundation", "funding", "scholarship", "fellowship",
    "application", "deadline", "cycle", "prize", "competition", "stipend",
    "SBIR", "STTR", "USDA", "SBA", "NASE", "Breva", "Comcast RISE",
    "QuickBooks Hero", "Awesome Foundation", "StreetShares"
]
ADMIN_KEYWORDS  = ["meeting", "call", "zoom", "review", "training", "certification", "license"]
OUTREACH_KEYWORDS = ["email", "follow up", "contact", "intro", "pitch", "proposal", "client"]
FINANCE_KEYWORDS  = ["invoice", "payment", "pay", "billing", "expense", "budget", "tax", "insurance"]
MARKETING_KEYWORDS = ["social media", "blog", "post", "content", "brand", "marketing", "SEO", "website"]

def auto_categorize(title, notes=""):
    """Guess category from task title and notes."""
    text = (title + " " + notes).lower()
    for kw in GRANT_KEYWORDS:
        if kw.lower() in text:
            return "Grant"
    for kw in FINANCE_KEYWORDS:
        if kw.lower() in text:
            return "Finance"
    for kw in MARKETING_KEYWORDS:
        if kw.lower() in text:
            return "Marketing"
    for kw in OUTREACH_KEYWORDS:
        if kw.lower() in text:
            return "Outreach"
    for kw in ADMIN_KEYWORDS:
        if kw.lower() in text:
            return "Admin"
    return "Operations"

# ── Provider registry ──────────────────────────────────────────────────────
PROVIDERS = {
    "Anthropic (Claude)": {
        "base_url": "https://api.anthropic.com",
        "default_model": "claude-haiku-4-5-20251001",
        "type": "anthropic",
        "key_hint": "sk-ant-api03-...",
        "free": False,
        "models": ["claude-haiku-4-5-20251001", "claude-sonnet-4-6", "claude-opus-4-6"],
    },
    "OpenRouter": {
        "base_url": "https://openrouter.ai/api/v1",
        "default_model": "anthropic/claude-3-haiku",
        "type": "openai_compat",
        "key_hint": "sk-or-v1-...",
        "free": True,
        "models": [
            # Anthropic (paid but reliable)
            "anthropic/claude-3-haiku",
            "anthropic/claude-3.5-haiku",
            "anthropic/claude-3.5-sonnet",
            # Meta Llama free tier (confirmed working)
            "meta-llama/llama-3.3-70b-instruct:free",
            "meta-llama/llama-3.1-8b-instruct:free",
            # Mistral free tier (confirmed working)
            "mistralai/mistral-7b-instruct:free",
            "mistralai/mistral-small:free",
            # Google (gemma-2 stable, gemma-3-12b REMOVED - dead endpoint)
            "google/gemma-2-9b-it:free",
            # DeepSeek free (reliable)
            "deepseek/deepseek-r1:free",
            "deepseek/deepseek-chat:free",
            # OpenAI (paid)
            "openai/gpt-4o-mini",
            "openai/gpt-4o",
        ],
    },
    "Groq": {
        "base_url": "https://api.groq.com/openai/v1",
        "default_model": "llama-3.1-8b-instant",
        "type": "openai_compat",
        "key_hint": "gsk_...",
        "free": True,
        "models": ["llama-3.1-8b-instant", "llama-3.3-70b-versatile",
                   "gemma2-9b-it", "mixtral-8x7b-32768"],
    },
    "Google Gemini": {
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai",
        "default_model": "gemini-2.0-flash",
        "type": "openai_compat",
        "key_hint": "AIza...",
        "free": True,
        "models": ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"],
    },
    "NVIDIA NIM": {
        "base_url": "https://integrate.api.nvidia.com/v1",
        "default_model": "meta/llama-3.1-8b-instruct",
        "type": "openai_compat",
        "key_hint": "nvapi-...",
        "free": True,
        "models": ["meta/llama-3.1-8b-instruct", "meta/llama-3.3-70b-instruct",
                   "nvidia/llama-3.1-nemotron-70b-instruct",
                   "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning",
                   "mistralai/mistral-7b-instruct-v0.3"],
    },
    "xAI Grok": {
        "base_url": "https://api.x.ai/v1",
        "default_model": "grok-beta",
        "type": "openai_compat",
        "key_hint": "xai-...",
        "free": False,
        "models": ["grok-beta"],
    },
    "OpenAI (GPT)": {
        "base_url": "https://api.openai.com/v1",
        "default_model": "gpt-4o-mini",
        "type": "openai_compat",
        "key_hint": "sk-...",
        "free": False,
        "models": ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"],
    },
    "Cohere": {
        "base_url": "https://api.cohere.com/v2",
        "default_model": "command-r-plus",
        "type": "openai_compat",
        "key_hint": "cohere-key...",
        "free": False,
        "models": ["command-r-plus", "command-r", "command-light"],
    },
    "Ollama (Local)": {
        "base_url": "http://localhost:11434/v1",
        "default_model": "llama3",
        "type": "openai_compat",
        "key_hint": "ollama",
        "free": True,
        "models": ["llama3", "llama3.1", "mistral", "qwen2.5-coder:7b", "phi3"],
    },
    "Custom": {
        "base_url": "",
        "default_model": "",
        "type": "openai_compat",
        "key_hint": "your-api-key",
        "free": False,
        "models": [],
    },
}
PROVIDER_NAMES = list(PROVIDERS.keys())

# ── Colour palette ─────────────────────────────────────────────────────────
C = {
    "bg":        "#f5fbf7",
    "surface":   "#ffffff",
    "alt":       "#f0faf2",
    "dark":      "#0D3520",
    "dark2":     "#1a3a2a",
    "mid":       "#2d6a4f",
    "soft":      "#52b788",
    "pale":      "#d8f3dc",
    "text":      "#1a2e22",
    "muted":     "#8aab96",
    "border":    "#c8e6d0",
    "amber":     "#e9a319",
    "amber_bg":  "#fef3dc",
    "red":       "#c0392b",
    "red_bg":    "#fdecea",
    "blue":      "#1a6fa8",
    "ai_bg":     "#071510",
    "ai_surf":   "#0d2618",
    "ai_border": "#1a4a2a",
    "ai_accent": "#76b900",
    "ai_text":   "#ccddcc",
    "card_bg":   "#f8fdf9",
}
PRI_COLOR = {"High": C["red"], "Medium": C["amber"], "Low": C["soft"]}

# ── Gradient/geometry generators ───────────────────────────────────────────
def make_gradient(w, h, c1, c2, angle=135):
    if not PIL_AVAILABLE: return None
    img = Image.new("RGB", (w, h))
    draw = ImageDraw.Draw(img)
    r1,g1,b1 = int(c1[0:2],16),int(c1[2:4],16),int(c1[4:6],16)
    r2,g2,b2 = int(c2[0:2],16),int(c2[2:4],16),int(c2[4:6],16)
    for i in range(h):
        t = i/max(h-1,1)
        r=int(r1+(r2-r1)*t); g=int(g1+(g2-g1)*t); b=int(b1+(b2-b1)*t)
        draw.line([(0,i),(w,i)], fill=(r,g,b))
    return img

def add_hex_overlay(img, line_color="76B900", opacity=15):
    if not PIL_AVAILABLE or img is None: return img
    import math
    w, h = img.size
    overlay = Image.new("RGBA", (w,h), (0,0,0,0))
    draw = ImageDraw.Draw(overlay)
    lc = tuple(int(line_color[i:i+2],16) for i in (0,2,4))
    R = 30; dx = R*math.sqrt(3); dy = R*1.5
    row = -1
    while row*dy < h+R:
        col = -1
        while col*dx < w+R:
            cx = col*dx + (row%2)*dx/2; cy = row*dy
            pts = [(cx+R*math.cos(math.radians(60*i-30)),
                    cy+R*math.sin(math.radians(60*i-30))) for i in range(6)]
            draw.polygon(pts, outline=lc+(opacity,), fill=None)
            col += 1
        row += 1
    return Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")

def make_topbar_img(w, h):
    base = make_gradient(w, h, "071510", "1B4A2A", 145)
    return add_hex_overlay(base, "76B900", 18) if base else None

def make_ai_img(w, h):
    base = make_gradient(w, h, "050E0A", "0D2618", 160)
    return add_hex_overlay(base, "76B900", 12) if base else None

def make_ws_img(w, h):
    base = make_gradient(w, h, "0A2015", "1B3D2A", 90)
    return add_hex_overlay(base, "52b788", 8) if base else None

def make_stats_img(w, h):
    base = make_gradient(w, h, "071510", "0F2A1A", 160)
    return add_hex_overlay(base, "76B900", 10) if base else None

def pil_to_tk(img):
    if img is None or not PIL_AVAILABLE: return None
    return ImageTk.PhotoImage(img)

# ── Gradient image cache (generated once, reused everywhere) ───────────────
_GRAD_CACHE = {}

def get_cached_gradient(fn_name, w, h):
    """Return cached gradient or generate and cache it."""
    key = (fn_name, w, h)
    if key not in _GRAD_CACHE:
        fns = {"topbar": make_topbar_img, "ai": make_ai_img,
               "ws": make_ws_img, "stats": make_stats_img}
        fn = fns.get(fn_name)
        if fn:
            img = fn(w, h)
            _GRAD_CACHE[key] = pil_to_tk(img) if img else None
        else:
            _GRAD_CACHE[key] = None
    return _GRAD_CACHE[key]

# ── Tooltip ────────────────────────────────────────────────────────────────
class Tooltip:
    def __init__(self, widget, text, anchor=None):
        self.widget = widget; self.text = text; self.anchor = anchor
        self._win = None; self._id = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)
        widget.bind("<Button-1>", self._click)

    def _schedule(self, e):
        self._id = self.widget.after(600, self._show)

    def _show(self):
        if self._win: return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._win = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        outer = tk.Frame(tw, bg=C["ai_accent"], padx=1, pady=1)
        outer.pack()
        inner = tk.Frame(outer, bg=C["ai_bg"], padx=12, pady=10)
        inner.pack()
        tk.Label(inner, text=self.text, bg=C["ai_bg"], fg=C["ai_text"],
                 font=("Arial",9), wraplength=280, justify="left").pack()
        if self.anchor:
            tk.Label(inner, text="Click for more info  →",
                     bg=C["ai_bg"], fg=C["ai_accent"],
                     font=("Arial",8,"italic"), cursor="hand2").pack(anchor="e", pady=(4,0))

    def _hide(self, e=None):
        if self._id: self.widget.after_cancel(self._id); self._id = None
        if self._win: self._win.destroy(); self._win = None

    def _click(self, e):
        if self.anchor:
            _open_help_browser(self.anchor)

def _apply_dark_titlebar(win):
    """Apply dark titlebar instantly. Call after withdraw(), before deiconify()."""
    try:
        import ctypes
        # update() (not just update_idletasks()) ensures the OS window handle
        # exists before we query it — eliminates the white-flash on slower machines.
        win.update()
        hwnd = ctypes.windll.user32.GetParent(win.winfo_id())
        if not hwnd:
            # Fallback: use the Tk internal frame handle
            try:
                hwnd = int(win.frame(), 16)
            except Exception:
                hwnd = win.winfo_id()
        val = ctypes.c_int(1)
        sz  = ctypes.sizeof(val)
        # Attribute 20 = DWMWA_USE_IMMERSIVE_DARK_MODE (Win11 / Win10 build 19041+)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(val), sz)
        # Attribute 19 = older Win10 pre-release builds
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 19, ctypes.byref(val), sz)
        # Force a non-client area repaint so the change takes effect before deiconify
        SWP_NOMOVE = 0x0002; SWP_NOSIZE = 0x0001; SWP_NOZORDER = 0x0004
        SWP_FRAMECHANGED = 0x0020
        ctypes.windll.user32.SetWindowPos(
            hwnd, None, 0, 0, 0, 0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
    except Exception:
        pass

def _open_help_browser(anchor=None):
    import webbrowser
    hp = os.path.join(BASE_DIR, "fcggrow_help.html")
    if os.path.exists(hp):
        url = "file:///" + hp.replace(chr(92), "/")
        if anchor: url += "#" + anchor
        webbrowser.open(url)

# ── Copy/paste context menu ────────────────────────────────────────────────
def add_copy_paste(widget):
    menu = tk.Menu(widget, tearoff=0, bg=C["surface"], fg=C["text"],
                   activebackground=C["mid"], activeforeground="white",
                   font=("Arial",9))
    menu.add_command(label="Cut",   command=lambda: widget.event_generate("<<Cut>>"))
    menu.add_command(label="Copy",  command=lambda: widget.event_generate("<<Copy>>"))
    menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
    menu.add_separator()
    menu.add_command(label="Select All", command=lambda: widget.event_generate("<<SelectAll>>"))
    def show(e):
        try: menu.tk_popup(e.x_root, e.y_root)
        finally: menu.grab_release()
    widget.bind("<Button-3>", show)
    widget.bind("<Control-a>", lambda e: widget.event_generate("<<SelectAll>>"))

# ── Workspace helpers ──────────────────────────────────────────────────────
def list_workspaces():
    if not os.path.exists(WORKSPACES_DIR): return []
    return sorted([d for d in os.listdir(WORKSPACES_DIR)
                   if os.path.isdir(os.path.join(WORKSPACES_DIR, d))])

def ws_path(name):
    safe = name.replace(" ","_").replace("/","_").replace(chr(92),"_")
    return os.path.join(WORKSPACES_DIR, safe)

def load_ws_config(name):
    p = os.path.join(ws_path(name), "config.json")
    if os.path.exists(p):
        with open(p) as f: return json.load(f)
    return {"org_name": name}

def save_ws_config(name, cfg):
    d = ws_path(name); os.makedirs(d, exist_ok=True)
    with open(os.path.join(d,"config.json"),"w") as f: json.dump(cfg,f,indent=2)

DEFAULT_TASKS = [
    {"id":1,"title":"Set up your first project","status":"todo","priority":"High",
     "category":"Operations","due":"","notes":"Welcome to FCG Grow!","done":False},
    {"id":2,"title":"Add your team members","status":"todo","priority":"Medium",
     "category":"Admin","due":"","notes":"Track who is responsible for what.","done":False},
    {"id":3,"title":"Review your pipeline","status":"inprog","priority":"High",
     "category":"Operations","due":"","notes":"Ask the AI to summarize your board.","done":False},
    {"id":4,"title":"Try the AI assistant","status":"todo","priority":"Medium",
     "category":"Other","due":"","notes":"Ask: what should I prioritize today?","done":False},
]

def load_ws_tasks(name):
    p = os.path.join(ws_path(name), "tasks.json")
    if os.path.exists(p):
        with open(p) as f: return json.load(f)
    tasks = [t.copy() for t in DEFAULT_TASKS]
    save_ws_tasks(name, tasks); return tasks

def save_ws_tasks(name, tasks):
    d = ws_path(name); os.makedirs(d, exist_ok=True)
    with open(os.path.join(d,"tasks.json"),"w") as f: json.dump(tasks,f,indent=2)

def load_session():
    if os.path.exists(SESSION_FILE):
        with open(SESSION_FILE) as f: return json.load(f)
    return {}

def save_session(data):
    with open(SESSION_FILE,"w") as f: json.dump(data,f,indent=2)

def next_id(tasks):
    return max((t["id"] for t in tasks), default=0) + 1

# ── ICS parser ─────────────────────────────────────────────────────────────
def parse_ics(data: bytes):
    if not ICAL_AVAILABLE: return []
    events = []
    try:
        cal = ICal.from_ical(data)
        for comp in cal.walk():
            if comp.name == "VEVENT":
                title = str(comp.get("SUMMARY","Calendar event"))
                dtstart = comp.get("DTSTART")
                due = ""
                if dtstart:
                    val = dtstart.dt
                    due = val.date().isoformat() if hasattr(val,"date") else val.isoformat()
                notes = str(comp.get("DESCRIPTION",""))
                events.append({"title":title,"due":due,"notes":notes})
    except Exception as e:
        print(f"ICS error: {e}")
    return events

# ── Document text extractor ────────────────────────────────────────────────
def extract_doc_text(filepath):
    """Extract text from PDF, DOCX, or TXT files."""
    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".pdf":
        if not PDF_AVAILABLE:
            return None, "pdfplumber not installed. Run: pip install pdfplumber"
        try:
            import pdfplumber
            with pdfplumber.open(filepath) as pdf:
                text = "\n".join(p.extract_text() or "" for p in pdf.pages)
            return text[:8000], None  # Limit to 8000 chars
        except Exception as e:
            return None, str(e)
    elif ext == ".docx":
        if not DOCX_AVAILABLE:
            return None, "python-docx not installed. Run: pip install python-docx"
        try:
            doc = python_docx.Document(filepath)
            text = "\n".join(p.text for p in doc.paragraphs)
            return text[:8000], None
        except Exception as e:
            return None, str(e)
    elif ext == ".txt":
        try:
            with open(filepath, encoding="utf-8", errors="ignore") as f:
                return f.read()[:8000], None
        except Exception as e:
            return None, str(e)
    else:
        return None, f"Unsupported file type: {ext}"

# ── NAICS codes ────────────────────────────────────────────────────────────
NAICS_CODES = [
    "111110 - Soybean Farming","111219 - Other Vegetable Farming",
    "112111 - Beef Cattle Ranching","112210 - Hog and Pig Farming",
    "115110 - Support for Crop Production","115210 - Support for Animal Production",
    "236110 - Residential Building Construction",
    "236220 - Commercial and Institutional Building Construction",
    "237310 - Highway, Street, and Bridge Construction",
    "238110 - Poured Concrete Contractors","238120 - Structural Steel Contractors",
    "238160 - Roofing Contractors","238210 - Electrical Contractors",
    "238220 - Plumbing and HVAC Contractors","238310 - Drywall and Insulation Contractors",
    "238320 - Painting and Wall Covering","238330 - Flooring Contractors",
    "238910 - Site Preparation Contractors","238990 - Other Specialty Trade Contractors",
    "311811 - Retail Bakeries","311812 - Commercial Bakeries",
    "312111 - Soft Drink Manufacturing","312120 - Breweries","312130 - Wineries",
    "423110 - Auto and Motor Vehicle Wholesalers",
    "441110 - New Car Dealers","441120 - Used Car Dealers",
    "444110 - Home Centers","444130 - Hardware Stores",
    "444220 - Nursery, Garden Center, Farm Supply",
    "444240 - Nursery and Garden Centers",
    "445110 - Supermarkets and Grocery","445291 - Health Food Stores",
    "446110 - Pharmacies and Drug Stores",
    "448110 - Men's Clothing","448120 - Women's Clothing",
    "448140 - Family Clothing","448310 - Jewelry Stores",
    "451110 - Sporting Goods Stores","451120 - Hobby, Toy, and Game Stores",
    "452210 - Department Stores","453310 - Used Merchandise Stores",
    "453910 - Pet and Pet Supplies","453998 - All Other Miscellaneous Retailers",
    "454110 - Electronic Shopping and Mail-Order",
    "481111 - Scheduled Passenger Air","484110 - General Freight Trucking Local",
    "484121 - General Freight Trucking Long-Distance Truckload",
    "485310 - Taxi and Ridesharing","485320 - Limousine Service",
    "488410 - Motor Vehicle Towing",
    "492110 - Couriers and Express Delivery","492210 - Local Messengers and Delivery",
    "493110 - General Warehousing and Storage",
    "511110 - Newspaper Publishers","511120 - Periodical Publishers",
    "511130 - Book Publishers","511210 - Software Publishers",
    "512110 - Motion Picture and Video Production",
    "515112 - Radio Stations","515120 - Television Broadcasting",
    "517111 - Wired Telecommunications","517112 - Wireless Telecommunications",
    "518210 - Computing Infrastructure and Web Hosting",
    "519130 - Internet Publishing and Web Search Portals",
    "519190 - All Other Information Services",
    "522110 - Commercial Banking","522120 - Savings Institutions",
    "522130 - Credit Unions","522291 - Consumer Lending",
    "523110 - Investment Banking","523120 - Securities Brokerage",
    "523920 - Portfolio Management","523930 - Investment Advice",
    "524113 - Direct Life Insurance","524114 - Direct Health Insurance",
    "524126 - Property and Casualty Insurance",
    "524210 - Insurance Agencies and Brokerages",
    "531110 - Residential Property Lessors",
    "531210 - Real Estate Agents and Brokers",
    "531311 - Residential Property Managers",
    "532111 - Passenger Car Rental","532120 - Truck and RV Rental",
    "541110 - Offices of Lawyers","541211 - Certified Public Accountants",
    "541213 - Tax Preparation Services","541214 - Payroll Services",
    "541310 - Architectural Services","541320 - Landscape Architectural Services",
    "541330 - Engineering Services","541380 - Testing Laboratories",
    "541410 - Interior Design","541430 - Graphic Design",
    "541490 - Other Specialized Design",
    "541511 - Custom Computer Programming",
    "541512 - Computer Systems Design",
    "541519 - Other Computer Related Services",
    "541611 - Administrative Management Consulting",
    "541613 - Marketing Consulting",
    "541618 - Other Management Consulting",
    "541810 - Advertising Agencies","541820 - Public Relations",
    "541921 - Photography Studios Portrait","541922 - Commercial Photography",
    "541930 - Translation and Interpretation",
    "541940 - Veterinary Services",
    "541990 - All Other Professional Services",
    "561110 - Office Administrative Services",
    "561310 - Employment Placement Agencies",
    "561320 - Temporary Help Services",
    "561410 - Document Preparation Services",
    "561420 - Telephone Answering Services",
    "561510 - Travel Agencies",
    "561611 - Investigation Services",
    "561612 - Security Guards and Patrol",
    "561710 - Exterminating and Pest Control",
    "561720 - Janitorial Services",
    "561730 - Landscaping Services",
    "561740 - Carpet and Upholstery Cleaning",
    "561790 - Other Services to Buildings",
    "561920 - Convention and Trade Show Organizers",
    "562111 - Solid Waste Collection",
    "611110 - Elementary and Secondary Schools",
    "611310 - Colleges and Universities",
    "611420 - Computer Training",
    "611430 - Professional Development Training",
    "611511 - Cosmetology and Barber Schools",
    "611610 - Fine Arts Schools",
    "611620 - Sports and Recreation Instruction",
    "611691 - Exam Preparation and Tutoring",
    "621111 - Offices of Physicians",
    "621210 - Offices of Dentists",
    "621310 - Offices of Chiropractors",
    "621330 - Mental Health Practitioners",
    "621610 - Home Health Care Services",
    "621910 - Ambulance Services",
    "622110 - General Medical and Surgical Hospitals",
    "623110 - Nursing Care Facilities",
    "624110 - Child and Youth Services",
    "624120 - Services for Elderly and Disabled",
    "624410 - Child Day Care Services",
    "711110 - Theater Companies","711130 - Musical Groups and Artists",
    "711211 - Sports Teams and Clubs",
    "713910 - Golf Courses and Country Clubs",
    "713940 - Fitness and Recreational Sports Centers",
    "721110 - Hotels and Motels","721191 - Bed and Breakfast Inns",
    "722310 - Food Service Contractors","722320 - Caterers",
    "722330 - Mobile Food Services",
    "722511 - Full-Service Restaurants",
    "722513 - Limited-Service Restaurants",
    "722515 - Snack and Nonalcoholic Beverage Bars",
    "811111 - General Automotive Repair",
    "811192 - Car Washes",
    "811212 - Computer and Office Machine Repair",
    "811411 - Home and Garden Equipment Repair",
    "811412 - Appliance Repair",
    "812111 - Barber Shops","812112 - Beauty Salons","812113 - Nail Salons",
    "812310 - Coin-Operated Laundries",
    "812320 - Drycleaning and Laundry Services",
    "812910 - Pet Care Services",
    "812990 - All Other Personal Services",
    "813110 - Religious Organizations",
    "813211 - Grantmaking Foundations",
    "813212 - Voluntary Health Organizations",
    "813410 - Civic and Social Organizations",
    "813910 - Business Associations",
    "813920 - Professional Organizations",
    "814110 - Private Households",
]

# ── NAICS intelligence ─────────────────────────────────────────────────────
NAICS_INTEL = {
    "561730": """INDUSTRY: Landscaping Services (NAICS 561730)
Revenue peaks spring/summer -- plan for winter cash flow gaps.
Key costs: labor 35-45%, fuel, equipment maintenance.
Standard crew: 2-3 residential, 4-6 commercial.
Insurance minimums for government: $1M/$2M GL, workers comp, commercial auto.
SDVOSB set-asides available -- VA and municipal contracts are primary targets.
Miller Act: bonds required on federal contracts over $35K.
Pricing: residential $40-80/visit; commercial $500-5000/month.
Growth: HOA contracts, property management firms, government grounds maintenance.
Certifications: NALP Landscape Professional, ISA Arborist, irrigation certs.""",

    "238": """INDUSTRY: Construction Trades (NAICS 238xxx)
Surety bonds required on most public works projects.
Davis-Bacon Act applies to federally funded construction.
OSHA 10/30 certifications expected by most GCs.
Cash flow: milestone billing, watch retainage (5-10%).
Certifications: MBE, WBE, DBE, SDVOSB open set-aside opportunities.""",

    "722": """INDUSTRY: Food Service (NAICS 722xxx)
Food cost target: 28-35% of revenue.
Labor cost target: 25-35%.
Health inspection logs must be current.
Third-party delivery takes 15-30% -- own channel preferred.
Respond to every review within 24 hours.""",

    "541": """INDUSTRY: Professional Services (NAICS 541xxx)
Document everything in SOWs -- scope creep is margin death.
No single client should exceed 30% of revenue.
GSA Schedule is the primary contracting vehicle for consulting.""",

    "484": """INDUSTRY: Trucking (NAICS 484xxx)
DOT compliance -- CSA scores affect insurance rates.
Fuel is the largest variable cost.
DAT and Truckstop load boards for spot market.""",

    "541511": """INDUSTRY: Software Development (NAICS 541511)
Two-week sprints with clear deliverables.
SaaS metrics: MRR, churn, CAC, LTV.
FedRAMP opens the federal market.
SOC 2 Type II for enterprise clients.""",

    "621": """INDUSTRY: Healthcare (NAICS 621xxx)
HIPAA compliance required for patient data.
Credentialing takes 90-120 days -- plan ahead.
Telehealth reduces overhead and expands reach.""",

    "813": """INDUSTRY: Nonprofits (NAICS 813xxx)
Form 990 is public -- donors review it.
Grant compliance: separate cost centers per grant.
Target 75%+ of expenses going to programs.""",

    "812": """INDUSTRY: Personal Services (NAICS 812xxx)
Target 60%+ repeat client rate.
Online scheduling reduces no-shows 20-30%.
Referral programs are most cost-effective acquisition.""",

    "519130": """INDUSTRY: Online Content/Blog (NAICS 519130)
Content consistency drives organic growth.
SEO and email list are the most durable channels.
Monetization: ads, affiliate, sponsored content, digital products.""",
}

def get_naics_intel(code):
    if not code: return ""
    if code in NAICS_INTEL: return NAICS_INTEL[code]
    for k in NAICS_INTEL:
        if code.startswith(k): return NAICS_INTEL[k]
    return ""

# ── AI prompts ─────────────────────────────────────────────────────────────
BASE_PROMPT = """You are FCG Grow, an AI operations assistant embedded in a Kanban task board.
You serve small business owners across all industries.
You are a capable, knowledgeable assistant. You can help with ANY topic the user asks about,
including coding, technical questions, writing, strategy, analysis, and more.
Do NOT refuse requests by saying you "cannot" do something -- just do it.

KNOWLEDGE: Cash flow, grant writing, government contracting (SAM.gov, SDVOSB, set-asides),
hiring, sales pipeline, marketing, operations efficiency, vendor negotiation, SBA programs,
strategic planning, coding and technical assistance, common small business pitfalls.

BOARD ACTIONS -- include JSON at END of response (never in the middle):
Add task:
```json
{"action":"add_task","title":"title","priority":"High|Medium|Low","category":"Grant|Operations|Admin|Outreach|Marketing|Finance|Other","due":"YYYY-MM-DD or empty","notes":"notes","status":"todo|inprog|review|done"}
```
Update task (change category, priority, title, due date, notes):
```json
{"action":"update_task","task_id":123,"category":"Grant","priority":"High","title":"new title","due":"2026-06-30","notes":"updated notes"}
```
Move task:
```json
{"action":"move_task","task_id":123,"status":"todo|inprog|review|done"}
```
Delete task:
```json
{"action":"delete_task","task_id":123}
```
Batch -- use array for multiple actions:
```json
[{"action":"move_task","task_id":1,"status":"review"},{"action":"update_task","task_id":2,"category":"Grant"}]
```

CRITICAL RULES:
- NEVER print the board state in your response -- just act and give a brief confirmation
- NEVER show JSON blocks to the user -- they execute silently
- NEVER refer to tasks by ID in your responses -- always use the task title/name
- For bulk operations use a JSON array with ALL actions at once -- do not do them one at a time
- When categorizing tasks: scan title and notes for keywords to pick the right category
- Grant keywords: grant, award, foundation, funding, deadline, cycle, NASE, SBA, USDA, Breva, QuickBooks Hero, scholarship, fellowship
- Admin keywords: meeting, call, zoom, training, certification, license, review
- Finance keywords: invoice, payment, insurance, tax, budget, expense
- Marketing keywords: social media, blog, post, content, brand, SEO, website
- Outreach keywords: email, follow up, contact, intro, pitch, proposal, client
- Before bulk deleting >5 tasks, ask for confirmation first
- Lead with the answer, be direct, no fluff
- When adding multiple tasks at once, batch them all in ONE JSON array -- never multiple separate blocks

RESPONSE STYLE: Direct, practical, concise. Bullet points for lists. No board state dumps.
Answer ANY question the user asks. If asked to write code, write it. If asked about a topic outside business, help anyway."""

def build_system_prompt(cfg):
    org      = cfg.get("org_name","your organization")
    naics    = cfg.get("naics","")
    biz_type = cfg.get("biz_type","small business")
    goals    = cfg.get("goals","")
    naics_line = f" (NAICS {naics})" if naics else ""
    intel = get_naics_intel(naics)
    prompt = BASE_PROMPT
    prompt += f"\n\nWORKSPACE: {org} -- {biz_type}{naics_line}"
    if goals: prompt += f"\nGOALS: {goals}"
    if intel: prompt += f"\n\n{intel}"
    return prompt

def build_board_context(tasks, max_per_col=50):
    """Build board context. Caps per-column to keep context window lean."""
    today = datetime.now().date().isoformat()
    lines = [f"\nBOARD ({datetime.now().strftime('%b %d %Y %I:%M %p')}):"]
    for key, label in zip(COL_KEYS, COL_LABELS):
        col = [t for t in tasks if t["status"]==key and not t.get("done")]
        lines.append(f"\n{label.upper()} ({len(col)} tasks):")
        if not col:
            lines.append("  (empty)")
        else:
            for t in col[:max_per_col]:
                ov = " [OVERDUE]" if t.get("due") and t["due"]<today else ""
                due = f" due {t['due']}" if t.get("due") else ""
                lines.append(f"  [ID:{t['id']}] {t['title']} | {t['priority']} | {t['category']}{due}{ov}")
            if len(col) > max_per_col:
                lines.append(f"  ... +{len(col)-max_per_col} more (ask to see more)")
    lines.append(f"\nDONE: {len([t for t in tasks if t.get('done')])} tasks")
    return "\n".join(lines)

# ── AI caller ──────────────────────────────────────────────────────────────
def fetch_models(provider_name, api_key, base_url):
    try:
        if not OPENAI_AVAILABLE: return []
        kw = {"api_key": api_key or "ollama"}
        if base_url: kw["base_url"] = base_url
        client = OpenAI(**kw)
        return sorted([m.id for m in client.models.list().data])
    except Exception:
        return []

def call_ai(providers, system_prompt, messages):
    errors = []
    # Thinking/reasoning models burn tokens on internal CoT before responding
    THINKING_PATTERNS = ("thinking", ":r1", "/r1", "qwq", "acree", "o1-",
                         "o3-", "reasoning", "deepseek-r", "claude-3-7",
                         "sonnet-3-7", "preview")

    for p in providers:
        name  = p.get("name","")
        key   = p.get("api_key","").strip()
        model = p.get("model","").strip()
        is_local = "ollama" in name.lower() or "local" in name.lower()
        if not model: continue
        if not key and not is_local: continue
        try:
            ptype      = p.get("type","openai_compat")
            base       = p.get("base_url","").strip()
            model_low  = model.lower()
            is_thinking = any(pat in model_low for pat in THINKING_PATTERNS)
            # Thinking models need headroom for internal reasoning tokens
            max_tok = 16000 if is_thinking else 4096

            if ptype == "anthropic":
                if not ANTHROPIC_AVAILABLE: raise Exception("pip install anthropic")
                client = anthropic.Anthropic(api_key=key)
                resp = client.messages.create(model=model, max_tokens=max_tok,
                                              system=system_prompt, messages=messages)
                return resp.content[0].text, name
            else:
                if not OPENAI_AVAILABLE: raise Exception("pip install openai")
                kw = {"api_key": key or "ollama"}
                if base: kw["base_url"] = base
                client = OpenAI(**kw)
                all_msgs = [{"role":"system","content":system_prompt}] + messages
                extra = {}
                # Some OpenRouter thinking models need reasoning suppressed in output
                if is_thinking and "openrouter" in base.lower():
                    extra["extra_body"] = {"include_reasoning": False}
                resp = client.chat.completions.create(
                    model=model, max_tokens=max_tok,
                    messages=all_msgs, **extra)
                return resp.choices[0].message.content, name
        except Exception as e:
            errors.append(f"{name}: {e}")
    raise Exception("All providers failed:\n" + "\n".join(errors) if errors else "No providers")

# ── Setup dialog ───────────────────────────────────────────────────────────
class SetupDialog:
    def __init__(self, parent, cfg=None, title="FCG Grow Setup"):
        self.result = None; self.cfg = cfg or {}; self._parent = parent
        self.win = tk.Toplevel(parent)
        self.win.withdraw()  # Hide until dark titlebar is applied
        self.win.title(title)
        self.win.geometry("560x760"); self.win.configure(bg=C["dark2"])
        self.win.resizable(True, True); self.win.minsize(480, 640)
        _apply_dark_titlebar(self.win)  # Apply before first paint
        self.win.deiconify()            # Now show with dark bar already set
        self.win.grab_set()
        self.win.protocol("WM_DELETE_WINDOW", self._cancel)
        icon_path = os.path.join(BASE_DIR, "fcggrow.ico")
        if os.path.exists(icon_path):
            try: self.win.iconbitmap(icon_path)
            except Exception: pass
        self._build()
        parent.wait_window(self.win)

    def _build(self):
        # Gradient header
        hdr = tk.Frame(self.win, bg=C["dark"], height=72)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        if PIL_AVAILABLE:
            try:
                self.win.update_idletasks()
                photo = get_cached_gradient("topbar", 560, 72)
                if photo:
                    lbl = tk.Label(hdr, image=photo, bd=0)
                    lbl.place(x=0, y=0, relwidth=1, relheight=1)
                    lbl._photo = photo
            except Exception: pass
        tk.Label(hdr, text="FCG Grow", bg="#071510", fg="white",
                 font=("Arial",18,"bold"), pady=8).pack()
        tk.Label(hdr, text="Built by Fairy Circle Garden  |  fairycirclegarden.com",
                 bg="#071510", fg=C["soft"], font=("Arial",9)).pack(pady=(0,8))

        # Scrollable body — dark background
        outer = tk.Frame(self.win, bg=C["dark2"])
        outer.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(outer, bg=C["dark2"], highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical",
                            command=self._canvas.yview,
                            style="Dark.Vertical.TScrollbar")
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)
        self.body = tk.Frame(self._canvas, bg=C["dark2"])
        wid = self._canvas.create_window((0,0), window=self.body, anchor="nw")
        self.body.bind("<Configure>", lambda e: self._canvas.configure(
            scrollregion=self._canvas.bbox("all")))
        self._canvas.bind("<Configure>", lambda e: self._canvas.itemconfig(wid, width=e.width))
        self._canvas.bind("<Enter>", lambda e: self._canvas.bind_all(
            "<MouseWheel>", lambda ev: self._canvas.yview_scroll(-1*(ev.delta//120),"units")))
        self._canvas.bind("<Leave>", lambda e: self._canvas.unbind_all("<MouseWheel>"))

        self._build_org()
        self._build_providers()

        bf = tk.Frame(self.win, bg=C["dark2"])
        bf.pack(fill="x", padx=20, pady=10)
        tk.Button(bf, text="Cancel", command=self._cancel, relief="flat",
                  bg=C["ai_surf"], fg=C["ai_text"], font=("Arial",10),
                  padx=14, pady=7, cursor="hand2").pack(side="left", padx=4)
        tk.Button(bf, text="Save and Launch", command=self._save, relief="flat",
                  bg=C["dark"], fg="white", font=("Arial",10,"bold"),
                  padx=14, pady=7, cursor="hand2").pack(side="right", padx=4)

    def _sec(self, label):
        f = tk.Frame(self.body, bg=C["dark2"])
        f.pack(fill="x", padx=20, pady=(12,4))
        tk.Label(f, text=label, bg=C["dark2"], fg=C["soft"],
                 font=("Arial",8,"bold")).pack(anchor="w")
        tk.Frame(f, bg=C["ai_border"], height=1).pack(fill="x", pady=(3,8))
        return f

    def _lbl_row(self, parent, text, hint_text=None, anchor_key=None):
        row = tk.Frame(parent, bg=C["dark2"])
        row.pack(fill="x")
        tk.Label(row, text=text, bg=C["dark2"], fg=C["ai_text"],
                 font=("Arial",9,"bold")).pack(side="left")
        if hint_text:
            q = tk.Label(row, text=" ? ", bg=C["mid"], fg="white",
                         font=("Arial",8,"bold"), cursor="hand2", padx=4, pady=1)
            q.pack(side="left", padx=6)
            Tooltip(q, hint_text, anchor=anchor_key)

    def _entry(self, parent, default="", secret=False):
        kw = {"show":"*"} if secret else {}
        e = tk.Entry(parent, font=("Arial",10), relief="flat",
                     bg=C["ai_surf"], fg=C["ai_text"],
                     insertbackground=C["ai_text"],
                     highlightbackground=C["ai_border"], highlightthickness=1, **kw)
        e.pack(fill="x", pady=(2,4), ipady=6)
        if default: e.insert(0, default)
        add_copy_paste(e)
        return e

    def _hint_lbl(self, parent, text):
        tk.Label(parent, text=text, bg=C["dark2"], fg=C["muted"],
                 font=("Arial",8), wraplength=460).pack(anchor="w", pady=(0,6))

    def _build_org(self):
        sec = self._sec("WORKSPACE")
        tk.Label(sec, text="Organization Name *", bg=C["dark2"], fg=C["ai_text"],
                 font=("Arial",9,"bold")).pack(anchor="w")
        self.org_e = self._entry(sec, self.cfg.get("org_name",""))

        self._lbl_row(sec, "NAICS Code",
            "NAICS stands for North American Industry Classification System.\n"
            "It is a 6-digit code that tells the AI your industry.\n"
            "FCG Grow uses it for industry-specific advice and grant guidance.\n"
            "Completely optional -- skip if unsure.",
            anchor_key="naics")
        naics_var = tk.StringVar(value=self.cfg.get("naics",""))

        # Fully dark custom NAICS picker: Entry + dropdown Listbox popup
        naics_frame = tk.Frame(sec, bg=C["ai_surf"],
                               highlightbackground=C["ai_border"], highlightthickness=1)
        naics_frame.pack(fill="x", pady=(2,4))
        naics_entry = tk.Entry(naics_frame, textvariable=naics_var,
                               font=("Arial",10), relief="flat",
                               bg=C["ai_surf"], fg=C["ai_text"],
                               insertbackground=C["ai_text"],
                               highlightthickness=0)
        naics_entry.pack(fill="x", ipady=6, padx=4)
        add_copy_paste(naics_entry)

        # Popup listbox for filtered results
        _naics_popup = [None]
        def _close_popup(e=None):
            if _naics_popup[0]:
                try: _naics_popup[0].destroy()
                except Exception: pass
                _naics_popup[0] = None

        def _pick(val):
            naics_var.set(val)
            _close_popup()
            naics_entry.focus_set()

        def _show_popup(filtered):
            _close_popup()
            if not filtered: return
            # Position below the entry
            naics_frame.update_idletasks()
            x = naics_frame.winfo_rootx()
            y = naics_frame.winfo_rooty() + naics_frame.winfo_height()
            w = naics_frame.winfo_width()
            pop = tk.Toplevel(self.win)
            pop.wm_overrideredirect(True)
            pop.wm_geometry(f"{w}x{min(180, len(filtered)*22+4)}+{x}+{y}")
            pop.configure(bg=C["dark2"])
            _naics_popup[0] = pop
            lb = tk.Listbox(pop, font=("Arial",9), relief="flat",
                            bg=C["dark2"], fg=C["ai_text"],
                            selectbackground=C["mid"], selectforeground="white",
                            highlightthickness=0, bd=0,
                            activestyle="none")
            sb = ttk.Scrollbar(pop, orient="vertical", command=lb.yview,
                               style="Dark.Vertical.TScrollbar")
            lb.config(yscrollcommand=sb.set)
            sb.pack(side="right", fill="y")
            lb.pack(side="left", fill="both", expand=True)
            for item in filtered:
                lb.insert("end", item)
            lb.bind("<ButtonRelease-1>", lambda e: _pick(lb.get(lb.curselection()[0])) if lb.curselection() else None)
            lb.bind("<Return>", lambda e: _pick(lb.get(lb.curselection()[0])) if lb.curselection() else None)
            # Close popup if window loses focus
            pop.bind("<FocusOut>", lambda e: self.win.after(100, _close_popup))
            self.win.bind("<ButtonPress>", lambda e: self.win.after(50, _close_popup), add="+")

        def filter_naics(e=None):
            typed = naics_var.get().lower().strip()
            if not typed:
                _close_popup(); return
            filtered = [n for n in NAICS_CODES if typed in n.lower()][:50]
            if filtered:
                _show_popup(filtered)
            else:
                _close_popup()

        naics_entry.bind("<KeyRelease>", filter_naics)
        naics_entry.bind("<FocusOut>", lambda e: self.win.after(150, _close_popup))

        # Expose a .get() compatible interface
        class _NaicsWidget:
            def get(self): return naics_var.get().strip()
        self.naics_e = _NaicsWidget()
        self._hint_lbl(sec, "Type a number or keyword to filter — e.g. '561730' or 'landscaping'")

        self._lbl_row(sec, "Business Type",
            "Describe what your business does in plain English.\n"
            "The AI uses this to tailor advice.\n"
            "Examples: Veteran-owned landscaping, Online reseller, Food truck.\n"
            "Leave blank if unsure.",
            anchor_key="first-time")
        self.biz_e = self._entry(sec, self.cfg.get("biz_type",""))

        self._lbl_row(sec, "Top Goals",
            "What are you working toward this year?\n"
            "The AI reads this with every message.\n"
            "Examples: Win a government contract, grow to 10 clients.",
            anchor_key="first-time")
        self.goals_t = tk.Text(sec, height=2, font=("Arial",10), relief="flat",
                                bg=C["ai_surf"], fg=C["ai_text"],
                                insertbackground=C["ai_text"],
                                highlightbackground=C["ai_border"], highlightthickness=1)
        self.goals_t.pack(fill="x", pady=(2,4))
        if self.cfg.get("goals"): self.goals_t.insert("1.0", self.cfg["goals"])

    def _build_providers(self):
        sec = self._sec("AI PROVIDERS")
        tk.Label(sec,
                 text="Add one or more AI providers. FCG Grow tries them in order and "
                      "falls through if one fails. All providers work equally well -- "
                      "free options work just as well as paid.",
                 bg=C["dark2"], fg=C["muted"], font=("Arial",8),
                 wraplength=460, justify="left").pack(anchor="w", pady=(0,8))
        self.pf = tk.Frame(sec, bg=C["dark2"])
        self.pf.pack(fill="x")
        tk.Button(sec, text="+ Add Provider", relief="flat", cursor="hand2",
                  bg=C["mid"], fg="white", font=("Arial",9,"bold"),
                  padx=10, pady=4, bd=0, command=self._add_prov).pack(anchor="w", pady=6)
        raw = self.cfg.get("providers", [{
            "name":"OpenRouter","api_key":"",
            "model":"meta-llama/llama-3.3-70b-instruct:free",
            "base_url":"https://openrouter.ai/api/v1","type":"openai_compat"
        }])
        self.prov_data = [p.copy() for p in raw]
        self._rebuild_provs()

    def _rebuild_provs(self):
        for w in self.pf.winfo_children(): w.destroy()
        self._pw = []
        for i, p in enumerate(self.prov_data):
            self._prov_card(i, p)

    def _prov_card(self, i, p):
        card = tk.Frame(self.pf, bg=C["ai_surf"],
                        highlightbackground=C["ai_border"], highlightthickness=1)
        card.pack(fill="x", pady=3)
        hdr = tk.Frame(card, bg=C["ai_surf"])
        hdr.pack(fill="x", padx=8, pady=6)
        num_colors = [C["ai_accent"], C["mid"], C["muted"]]
        tk.Label(hdr, text=f" {i+1} ", bg=num_colors[min(i,2)], fg=C["dark"],
                 font=("Arial",9,"bold")).pack(side="left")
        name_lbl = tk.Label(hdr, text=p.get("name","Provider"),
                             bg=C["ai_surf"], fg="white", font=("Arial",10,"bold"))
        name_lbl.pack(side="left", padx=6)
        preset = PROVIDERS.get(p.get("name",""),{})
        if preset.get("free"):
            tk.Label(hdr, text="FREE", bg=C["ai_accent"], fg=C["dark"],
                     font=("Arial",7,"bold"), padx=5, pady=1).pack(side="left")
        elif p.get("name","") not in ("","Custom","Ollama (Local)"):
            tk.Label(hdr, text="PAID", bg=C["mid"], fg="white",
                     font=("Arial",7,"bold"), padx=5, pady=1).pack(side="left")
        if i > 0:
            tk.Button(hdr, text="Remove", relief="flat", cursor="hand2",
                      bg=C["ai_surf"], fg=C["red"], font=("Arial",8), bd=0,
                      command=lambda idx=i: self._rm_prov(idx)).pack(side="right")

        body = tk.Frame(card, bg=C["ai_surf"])
        body.pack(fill="x", padx=8, pady=(0,8))
        body.columnconfigure(1, weight=1)
        row = 0

        # Provider dropdown
        tk.Label(body, text="Provider", bg=C["ai_surf"], fg=C["muted"],
                 font=("Arial",8,"bold")).grid(row=row, column=0, sticky="w", pady=2)
        pv = tk.StringVar(value=p.get("name",PROVIDER_NAMES[0]))
        pdd_frame = tk.Frame(body, bg=C["ai_surf"],
                             highlightbackground=C["ai_border"], highlightthickness=1)
        pdd_frame.grid(row=row, column=1, sticky="ew", padx=(8,0), pady=2)
        pdd = tk.OptionMenu(pdd_frame, pv, *PROVIDER_NAMES)
        pdd.config(relief="flat", bg=C["ai_surf"], fg=C["ai_text"],
                   activebackground=C["mid"], activeforeground="white",
                   highlightthickness=0, bd=0, font=("Arial",9),
                   indicatoron=True, anchor="w")
        pdd["menu"].config(bg=C["dark2"], fg=C["ai_text"],
                           activebackground=C["mid"], activeforeground="white",
                           font=("Arial",9))
        pdd.pack(fill="x")
        # Wire change event same as before
        pv.trace_add("write", lambda *a, pv=pv, mv=mv if False else None: None)
        row += 1

        # API Key with ? tooltip
        key_row = tk.Frame(body, bg=C["ai_surf"])
        key_row.grid(row=row, column=0, sticky="w", pady=2)
        tk.Label(key_row, text="API Key", bg=C["ai_surf"], fg=C["muted"],
                 font=("Arial",8,"bold")).pack(side="left")
        q_key = tk.Label(key_row, text=" ? ", bg=C["mid"], fg="white",
                         font=("Arial",7,"bold"), cursor="hand2", padx=3)
        q_key.pack(side="left", padx=2)
        Tooltip(q_key,
                "This is your AI service password.\n"
                "Each provider gives you a unique key from their website.\n"
                "Click for a step-by-step guide to getting one free.",
                anchor="providers")

        existing_key = p.get("api_key","")
        kv = tk.StringVar(value=existing_key)
        ke = tk.Entry(body, textvariable=kv, show="*", font=("Arial",9),
                      relief="flat", bg=C["ai_surf"], fg=C["ai_text"],
                      insertbackground=C["ai_text"],
                      highlightbackground=C["ai_border"], highlightthickness=1, width=34)
        ke.grid(row=row, column=1, sticky="ew", padx=(8,0), pady=2)
        add_copy_paste(ke)
        hint_txt = preset.get("key_hint","")
        if not existing_key and hint_txt:
            ke.insert(0, hint_txt); ke.config(fg=C["muted"])
            def _fi(e, entry=ke, h=hint_txt):
                if entry.get()==h: entry.delete(0,"end"); entry.config(fg=C["ai_text"])
            ke.bind("<FocusIn>", _fi)
        row += 1

        # Free key help link
        FREE_LINKS = {
            "OpenRouter": ("providers","Get a free key at openrouter.ai"),
            "Groq":       ("providers","Get a free key at console.groq.com"),
            "NVIDIA NIM": ("providers","Get $50 free at build.nvidia.com"),
            "Ollama (Local)": ("providers","No key needed -- download at ollama.com"),
            "Google Gemini":  ("providers","Get a free key at aistudio.google.com"),
        }
        if p.get("name","") in FREE_LINKS:
            anchor_k, link_txt = FREE_LINKS[p["name"]]
            hl = tk.Label(body, text=link_txt + "  (click for guide)",
                          bg=C["ai_surf"], fg=C["soft"],
                          font=("Arial",8,"italic"), cursor="hand2")
            hl.grid(row=row, column=0, columnspan=2, sticky="w", pady=(0,4))
            hl.bind("<Button-1>", lambda e, a=anchor_k: _open_help_browser(a))
            row += 1

        # Model
        tk.Label(body, text="Model", bg=C["ai_surf"], fg=C["muted"],
                 font=("Arial",8,"bold")).grid(row=row, column=0, sticky="w", pady=2)
        mv = tk.StringVar(value=p.get("model", preset.get("default_model","")))
        models_list = preset.get("models",[]) or [""]
        if mv.get() not in models_list and mv.get():
            models_list = [mv.get()] + models_list
        mc_frame = tk.Frame(body, bg=C["ai_surf"],
                            highlightbackground=C["ai_border"], highlightthickness=1)
        mc_frame.grid(row=row, column=1, sticky="ew", padx=(8,0), pady=2)
        mc = tk.OptionMenu(mc_frame, mv, *models_list)
        mc.config(relief="flat", bg=C["ai_surf"], fg=C["ai_text"],
                  activebackground=C["mid"], activeforeground="white",
                  highlightthickness=0, bd=0, font=("Arial",9),
                  indicatoron=True, anchor="w")
        mc["menu"].config(bg=C["dark2"], fg=C["ai_text"],
                          activebackground=C["mid"], activeforeground="white",
                          font=("Arial",9))
        mc.pack(fill="x")
        row += 1

        # Base URL for custom/ollama
        uv = tk.StringVar(value=p.get("base_url", preset.get("base_url","")))
        if p.get("name","") in ("Custom","Ollama (Local)"):
            tk.Label(body, text="Base URL", bg=C["ai_surf"], fg=C["muted"],
                     font=("Arial",8,"bold")).grid(row=row, column=0, sticky="w", pady=2)
            ue = tk.Entry(body, textvariable=uv, font=("Arial",9), relief="flat",
                          bg=C["ai_surf"], fg=C["ai_text"],
                          insertbackground=C["ai_text"],
                          highlightbackground=C["ai_border"], highlightthickness=1, width=34)
            ue.grid(row=row, column=1, sticky="ew", padx=(8,0), pady=2)
            add_copy_paste(ue)
            row += 1

        # Refresh models button
        def _refresh(pv=pv, kv=kv, uv=uv, mc=mc, mv=mv, ke=ke):
            n = pv.get(); ps = PROVIDERS.get(n,{})
            k = kv.get().strip()
            if k == ps.get("key_hint",""): k = ""
            url = uv.get().strip() or ps.get("base_url","")
            ids = fetch_models(n, k, url)
            if ids:
                mc["menu"].delete(0,"end")
                for m in ids:
                    mc["menu"].add_command(label=m, command=tk._setit(mv, m))
                if not mv.get() or mv.get() not in ids: mv.set(ids[0])
                messagebox.showinfo("Models updated", f"Found {len(ids)} models.", parent=self.win)
            else:
                messagebox.showwarning("Could not fetch",
                    "Could not fetch live models.\nCheck your API key.", parent=self.win)

        tk.Button(body, text="Refresh model list", relief="flat", cursor="hand2",
                  bg=C["mid"], fg="white", font=("Arial",8),
                  padx=8, pady=2, bd=0, command=_refresh
                  ).grid(row=row, column=1, sticky="e", pady=(4,0))

        # Wire provider change
        def _on_change(*args, pv=pv, mv=mv, uv=uv, kv=kv, ke=ke, mc=mc, nl=name_lbl):
            nn = pv.get(); np = PROVIDERS.get(nn,{})
            mv.set(np.get("default_model",""))
            uv.set(np.get("base_url",""))
            # Rebuild OptionMenu choices
            new_models = np.get("models",[]) or [""]
            mc["menu"].delete(0,"end")
            for m in new_models:
                mc["menu"].add_command(label=m, command=tk._setit(mv, m))
            kv.set(""); ke.delete(0,"end")
            nh = np.get("key_hint","")
            if nh: ke.insert(0,nh); ke.config(fg=C["muted"])
            nl.config(text=nn)
        pv.trace_add("write", lambda *a: _on_change())

        self._pw.append({"pv":pv,"kv":kv,"mv":mv,"uv":uv,"ke":ke,"idx":i})

    def _add_prov(self):
        self.prov_data.append({"name":"OpenRouter","api_key":"",
            "model":"meta-llama/llama-3.3-70b-instruct:free",
            "base_url":"https://openrouter.ai/api/v1","type":"openai_compat"})
        self._rebuild_provs()

    def _rm_prov(self, idx):
        if len(self.prov_data) <= 1:
            messagebox.showwarning("Cannot remove","Need at least one provider.",parent=self.win)
            return
        self.prov_data.pop(idx); self._rebuild_provs()

    def _collect_providers(self):
        providers = []
        for w in self._pw:
            n = w["pv"].get(); ps = PROVIDERS.get(n,{})
            k = w["kv"].get().strip()
            if k == ps.get("key_hint",""): k = ""
            providers.append({"name":n,"api_key":k,"model":w["mv"].get().strip(),
                "base_url":w["uv"].get().strip() or ps.get("base_url",""),
                "type":ps.get("type","openai_compat")})
        return providers

    def _cancel(self): self.win.destroy()

    def _save(self):
        org = self.org_e.get().strip()
        if not org:
            messagebox.showwarning("Required","Please enter your organization name.",parent=self.win)
            return
        self.result = {"org_name":org,"naics":self.naics_e.get().strip(),
            "biz_type":self.biz_e.get().strip(),
            "goals":self.goals_t.get("1.0","end").strip(),
            "providers":self._collect_providers()}
        self.win.destroy()

# ── Main Application ───────────────────────────────────────────────────────
class FCGGrowApp:
    def __init__(self):
        self.drag_data = {}; self.filter_cat = "All"
        self.chat_history = []; self.ai_busy = False
        self.active_provider = ""; self._mb_drag = {}
        self._pending_bulk_actions = None

        Root = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk
        self.root = Root()
        self.root.withdraw()
        apply_dark_ttk_style(self.root)  # Apply once globally before any widgets
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.render_all())

        session = load_session()
        last_ws = session.get("last_workspace")
        workspaces = list_workspaces()

        if not workspaces:
            cfg = self._run_setup()
            if not cfg: self.root.destroy(); return
            ws_name = cfg["org_name"]
            save_ws_config(ws_name, cfg)
            self.current_ws = ws_name
        else:
            self.current_ws = last_ws if last_ws in workspaces else workspaces[0]

        self._load_ws(self.current_ws)
        self.root.title(f"FCG Grow  --  {self.current_ws.replace('_',' ')}")
        self.root.geometry("1280x800")
        self.root.minsize(900, 600)
        self.root.configure(bg=C["dark"])
        # Style the menu bar row to match dark theme
        try:
            self.root.option_add("*Menu.Background", C["dark2"])
            self.root.option_add("*Menu.Foreground", C["ai_text"])
            self.root.option_add("*Menu.activeBackground", C["mid"])
            self.root.option_add("*Menu.activeForeground", "white")
            self.root.option_add("*Menu.relief", "flat")
        except Exception: pass
        icon_path = os.path.join(BASE_DIR, "fcggrow.ico")
        if os.path.exists(icon_path):
            try: self.root.iconbitmap(icon_path)
            except Exception: pass
        # Dark title bar on Windows — apply before deiconify so it's never white
        _apply_dark_titlebar(self.root)
        self.root.deiconify()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        # Reduce resize glitch -- defer redraws during resize
        self.root.bind("<Configure>", self._on_resize)
        self._resize_after = None
        self._build_ui()
        self.render_all()
        self._ai_welcome()
        self.root.mainloop()

    def _load_ws(self, name):
        self.current_ws = name
        self.cfg = load_ws_config(name)
        self.tasks = load_ws_tasks(name)
        self.chat_history = []

    def _save_ws(self):
        save_ws_tasks(self.current_ws, self.tasks)
        save_session({"last_workspace": self.current_ws})

    def _run_setup(self, cfg=None):
        dlg = SetupDialog(self.root, cfg or {})
        return dlg.result

    def _on_resize(self, e):
        """Debounce resize — suppress gradient redraws during drag, apply once when done."""
        if e.widget is not self.root:
            return  # Ignore child widget configure events
        if self._resize_after:
            self.root.after_cancel(self._resize_after)
        self._resize_after = self.root.after(200, self._finish_resize)

    def _finish_resize(self):
        """Called once after resize stops — refresh topbar gradient at new width."""
        self._resize_after = None
        try:
            sw = self.root.winfo_width()
            if sw > 10 and PIL_AVAILABLE:
                # Invalidate topbar cache entry and regenerate
                for key in list(_GRAD_CACHE.keys()):
                    if key[0] == "topbar":
                        del _GRAD_CACHE[key]
                photo = get_cached_gradient("topbar", sw, 80)
                if photo and hasattr(self, "_topbar_bg_lbl"):
                    self._topbar_bg_lbl.config(image=photo)
                    self._topbar_bg_lbl._photo = photo
        except Exception:
            pass

    def _on_close(self):
        self._save_ws(); self.root.destroy()

    # ── UI ─────────────────────────────────────────────────────────────────
    def _build_ui(self):
        self._build_topbar()
        self._build_stats()
        self._build_toolbar()
        self.main_frame = tk.Frame(self.root, bg=C["dark"])
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0)
        self._build_board()
        self._build_ai_panel()
        self._build_statusbar()

    def _apply_bg_image(self, frame, img_fn, w, h, store_attr):
        if not PIL_AVAILABLE: return
        try:
            img = img_fn(w, h)
            if img:
                photo = pil_to_tk(img)
                setattr(self, store_attr, photo)
                lbl = tk.Label(frame, image=photo, bd=0)
                lbl.place(x=0, y=0, relwidth=1, relheight=1)
        except Exception: pass

    def _build_menubar(self):
        menubar = tk.Menu(self.root, bg=C["dark2"], fg=C["ai_text"],
                          activebackground=C["mid"], activeforeground="white",
                          relief="flat", bd=0)
        # Custom menu -- not applied to root (would show white bar)
        # self.root.config(menu=menubar)  # disabled

        file_menu = tk.Menu(menubar, tearoff=0, bg=C["dark2"], fg=C["ai_text"],
                            activebackground=C["mid"], activeforeground="white")
        file_menu.add_command(label="Import Calendar", command=self.import_calendar)
        file_menu.add_command(label="Import Document", command=self.import_document)
        file_menu.add_command(label="Export Tasks", command=self.export_tasks)
        file_menu.add_separator()
        file_menu.add_command(label="New Workspace", command=self._new_ws)
        file_menu.add_command(label="Delete Workspace", command=self._del_ws)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        view_menu = tk.Menu(menubar, tearoff=0, bg=C["dark2"], fg=C["ai_text"],
                            activebackground=C["mid"], activeforeground="white")
        view_menu.add_command(label="Toggle AI Panel", command=self._toggle_ai)
        view_menu.add_separator()
        for cat in ["All"] + CATEGORIES:
            view_menu.add_command(label=f"Filter: {cat}",
                                  command=lambda c=cat: self.set_filter(c))
        menubar.add_cascade(label="View", menu=view_menu)

        tools_menu = tk.Menu(menubar, tearoff=0, bg=C["dark2"], fg=C["ai_text"],
                             activebackground=C["mid"], activeforeground="white")
        tools_menu.add_command(label="Settings", command=self._open_settings)
        tools_menu.add_command(label="? Help", command=self._open_help)
        tools_menu.add_separator()
        tools_menu.add_command(label="About FCG Grow",
                               command=lambda: self._open_help("about"))
        tools_menu.add_command(label="♥ Support FCG Grow",
                               command=lambda: _open_help_browser("about"))
        menubar.add_cascade(label="Tools", menu=tools_menu)

    def _build_topbar(self):
        # Single seamless gradient bar — topbar + workspace tabs merged
        bar = tk.Frame(self.root, bg=C["dark"], height=80, bd=0, highlightthickness=0)
        bar.pack(fill="x"); bar.pack_propagate(False)
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        if PIL_AVAILABLE:
            try:
                photo = get_cached_gradient("topbar", sw, 80)
                if photo:
                    lbl = tk.Label(bar, image=photo, bd=0)
                    lbl.place(x=0, y=0, relwidth=1, relheight=1)
                    lbl._photo = photo
                    self._topbar_bg_lbl = lbl
            except Exception:
                pass

        TOP_BG = "#0c2316"   # gradient top — matches logo area at y≈0
        BOT_BG = "#143921"
        BTN_BG = "#143921"   # mid-gradient colour for button row at y≈6

        # Logo — top-left, row 1
        lf = tk.Frame(bar, bg=TOP_BG, bd=0, highlightthickness=0)
        lf.place(x=14, y=4)
        tk.Label(lf, text="FCG Grow", bg=TOP_BG, fg="white",
                 font=("Arial",13,"bold")).pack(anchor="w")
        tk.Label(lf, text="by FCG · 2026", bg=TOP_BG, fg=C["soft"],
                 font=("Arial",7)).pack(anchor="w")

        # Action buttons — top-right, row 1
        bf = tk.Frame(bar, bg=BTN_BG, bd=0, highlightthickness=0)
        bf.place(relx=1.0, x=-12, y=6, anchor="ne")
        tk.Button(bf, text="⚙", relief="flat", cursor="hand2",
                  bg=BTN_BG, fg=C["soft"],
                  activebackground=C["mid"], activeforeground="white",
                  font=("Arial",13), padx=8, pady=2, bd=0,
                  command=self._open_settings).pack(side="left", padx=2)
        tk.Button(bf, text="AI", relief="flat", cursor="hand2",
                  bg=C["ai_accent"], fg=C["dark"],
                  font=("Arial",10,"bold"), padx=10, pady=4, bd=0,
                  command=self._toggle_ai).pack(side="left", padx=4)
        self._tb(bf, "+ Add task", lambda: self.open_task_dialog(), primary=True)

        # Menu buttons — center, row 1
        mf = tk.Frame(bar, bg=BTN_BG, bd=0, highlightthickness=0)
        mf.place(relx=0.5, y=14, anchor="n")
        for mlabel, mcmd in [
            ("File ▾",  self._show_file_menu),
            ("View ▾",  self._show_view_menu),
            ("Tools ▾", self._show_tools_menu),
        ]:
            tk.Button(mf, text=mlabel, relief="flat", cursor="hand2",
                      bg=BTN_BG, fg="#ccddcc",
                      activebackground=C["mid"], activeforeground="white",
                      font=("Arial",9), padx=10, pady=3, bd=0,
                      command=mcmd).pack(side="left", padx=2)

        # Workspace tab row — sits at bottom-left of topbar, content-width only
        WS_ROW_BG = "#1B4A2A"
        self._ws_row_bg = WS_ROW_BG
        self.ws_bar = tk.Frame(bar, bg=WS_ROW_BG, bd=0, highlightthickness=0)
        self.ws_bar.place(x=8, y=52)   # fixed x,y — no relwidth, so it's content-wide only
        self._rebuild_ws_tabs()

    def _tb(self, parent, text, cmd, primary=False):
        tk.Button(parent, text=text, command=cmd, relief="flat", cursor="hand2",
                  bg=C["mid"] if primary else "#1B4A2A",
                  fg="white" if primary else "#ccddcc",
                  activebackground=C["dark2"], activeforeground="white",
                  font=("Arial",10,"bold" if primary else "normal"),
                  padx=10, pady=4, bd=0).pack(side="left", padx=3)

    def _build_ws_bar(self):
        pass  # Workspace bar now embedded in topbar

    def _rebuild_ws_tabs(self):
        for w in self.ws_bar.winfo_children(): w.destroy()
        WS_ACTIVE   = "#2d6a4f"   # bright green pill for active tab
        # Inactive tabs and + / ✕ buttons use the same bg as the ws_bar row
        # so they look like they float on the gradient with no colour break.
        WS_INACTIVE = getattr(self, "_ws_row_bg", "#1B4A2A")
        WS_ROW_BG   = WS_INACTIVE
        for ws in list_workspaces():
            active = ws == self.current_ws
            tab = tk.Button(self.ws_bar, text=ws.replace("_"," "),
                      relief="flat", cursor="hand2",
                      bg=WS_ACTIVE if active else WS_INACTIVE,
                      fg="white" if active else C["muted"],
                      font=("Arial",8,"bold" if active else "normal"),
                      padx=10, pady=2, bd=0,
                      command=lambda w=ws: self._switch_ws(w))
            tab.pack(side="left", padx=2, pady=4)
        # Compact "+" button instead of full "New Workspace" bar
        tk.Button(self.ws_bar, text=" + ", relief="flat", cursor="hand2",
                  bg=WS_ROW_BG, fg=C["ai_accent"],
                  font=("Arial",10,"bold"), padx=6, pady=1, bd=0,
                  command=self._new_ws).pack(side="left", padx=2, pady=4)
        if len(list_workspaces()) > 1:
            tk.Button(self.ws_bar, text="✕", relief="flat", cursor="hand2",
                      bg=WS_ROW_BG, fg=C["muted"], font=("Arial",8), padx=5, pady=1, bd=0,
                      command=self._del_ws).pack(side="left", padx=0, pady=4)

    def _switch_ws(self, name):
        self._save_ws(); self._load_ws(name)
        self.root.title(f"FCG Grow  --  {name.replace('_',' ')}")
        self._rebuild_ws_tabs(); self.render_all()
        # Update AI panel header
        if hasattr(self, 'prov_lbl'):
            org = self.cfg.get("org_name","Your Business")
            self.ai_org_lbl.config(text=org)
        self.chat_display.config(state="normal")
        self.chat_display.delete("1.0","end")
        self.chat_display.config(state="disabled")
        self._ai_welcome()
        self._toast(f"Switched to {name.replace('_',' ')}")

    def _new_ws(self):
        cfg = self._run_setup()
        if not cfg: return
        ws_name = cfg["org_name"]
        if ws_name.replace(" ","_") in list_workspaces():
            messagebox.showwarning("Exists",f"Workspace '{ws_name}' already exists.",parent=self.root)
            return
        save_ws_config(ws_name, cfg); self._switch_ws(ws_name)

    def _del_ws(self):
        if not messagebox.askyesno("Delete workspace",
            f"Delete '{self.current_ws.replace('_',' ')}' and all its tasks?\nThis cannot be undone.",
            parent=self.root): return
        shutil.rmtree(ws_path(self.current_ws), ignore_errors=True)
        remaining = list_workspaces()
        if remaining: self._switch_ws(remaining[0])
        else: self._new_ws()

    def _open_settings(self):
        dlg = SetupDialog(self.root, self.cfg,
                          title=f"Settings -- {self.current_ws.replace('_',' ')}")
        if dlg.result:
            self.cfg.update(dlg.result)
            save_ws_config(self.current_ws, self.cfg)
            self._toast("Settings saved")

    def _open_help(self, anchor=None):
        _open_help_browser(anchor)

    def _build_stats(self):
        # Slim stat strip — compact, blends into dark chrome
        self.stats_bar = tk.Frame(self.root, bg=C["dark"], height=48, bd=0, highlightthickness=0)
        self.stats_bar.pack(fill="x"); self.stats_bar.pack_propagate(False)
        self.stat_vars = {}
        inner = tk.Frame(self.stats_bar, bg=C["dark"], bd=0, highlightthickness=0)
        inner.pack(anchor="w", padx=14, pady=6)
        for label, key in [("Total","total"),("To Do","todo"),
                            ("In Progress","inprog"),("Active","active"),("Done","done")]:
            f = tk.Frame(inner, bg=C["dark2"],
                         highlightbackground=C["ai_border"], highlightthickness=1)
            f.pack(side="left", padx=3, ipadx=8, ipady=2)
            tk.Label(f, text=label, bg=C["dark2"],
                     fg=C["ai_accent"] if key=="total" else C["muted"],
                     font=("Arial",7)).pack()
            v = tk.StringVar(value="0"); self.stat_vars[key] = v
            tk.Label(f, textvariable=v, bg=C["dark2"],
                     fg="white" if key=="total" else C["ai_text"],
                     font=("Arial",13,"bold")).pack()

    def _build_toolbar(self):
        tb = tk.Frame(self.root, bg=C["dark"], bd=0, highlightthickness=0)
        tb.pack(fill="x", padx=0, pady=0)
        inner = tk.Frame(tb, bg=C["dark"]); inner.pack(anchor="w", padx=12, pady=(4,6))
        tk.Label(inner, text="Filter:", bg=C["dark"], fg=C["muted"],
                 font=("Arial",9)).pack(side="left", padx=(0,4))
        self.filter_buttons = {}
        # Active filter  = lime #76b900 bg + dark text (high contrast, stands out)
        # Inactive filter = dark surface bg + lime text (readable, clearly clickable)
        ACTIVE_BG   = C["ai_accent"]   # #76b900 lime bg
        ACTIVE_FG   = C["dark"]        # dark text on lime
        INACTIVE_BG = C["ai_surf"]     # #0d2618 dark bg
        INACTIVE_FG = C["ai_accent"]   # lime text on dark — clearly lime but not active
        for cat in ["All"] + CATEGORIES:
            active = (cat == "All")
            btn = tk.Button(inner, text=cat, relief="flat", cursor="hand2",
                            bg=ACTIVE_BG if active else INACTIVE_BG,
                            fg=ACTIVE_FG if active else INACTIVE_FG,
                            activebackground=C["ai_accent"], activeforeground=C["dark"],
                            font=("Arial",9,"bold" if active else "normal"),
                            padx=8, pady=2, bd=0,
                            command=lambda c=cat: self.set_filter(c))
            btn.pack(side="left", padx=2)
            self.filter_buttons[cat] = btn
            btn._active_bg = ACTIVE_BG; btn._active_fg = ACTIVE_FG
            btn._inactive_bg = INACTIVE_BG; btn._inactive_fg = INACTIVE_FG
        tk.Label(inner, text="Search:", bg=C["dark"], fg=C["muted"],
                 font=("Arial",9)).pack(side="left", padx=(12,4))
        se = tk.Entry(inner, textvariable=self.search_var, width=20,
                      relief="flat", font=("Arial",9),
                      bg=C["ai_surf"], fg=C["ai_text"],
                      insertbackground=C["ai_text"],
                      highlightbackground=C["ai_border"], highlightthickness=1)
        se.pack(side="left")

    def _build_board(self):
        self.board_container = tk.Frame(self.main_frame, bg=C["dark"])
        self.board_container.pack(side="left", fill="both", expand=True)
        canvas = tk.Canvas(self.board_container, bg=C["dark"], highlightthickness=0)
        # Dark scrollbar
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Dark.Vertical.TScrollbar",
                        background=C["dark2"], troughcolor=C["dark"],
                        arrowcolor=C["soft"], bordercolor=C["dark"],
                        lightcolor=C["dark2"], darkcolor=C["dark2"])
        style.map("Dark.Vertical.TScrollbar",
                  background=[("active", C["mid"]), ("disabled", C["dark"])])
        vsb = ttk.Scrollbar(self.board_container, orient="vertical",
                            command=canvas.yview, style="Dark.Vertical.TScrollbar")
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y"); canvas.pack(side="left", fill="both", expand=True)
        self.board_frame = tk.Frame(canvas, bg=C["dark"])
        wid = canvas.create_window((0,0), window=self.board_frame, anchor="nw")
        self.board_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(wid, width=e.width))
        self._board_canvas = canvas
        canvas.bind("<Enter>", lambda e: canvas.bind_all(
            "<MouseWheel>", lambda ev: canvas.yview_scroll(-1*(ev.delta//120),"units")))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        self.col_inner = {}
        for i, (label, key) in enumerate(zip(COL_LABELS, COL_KEYS)):
            self.board_frame.columnconfigure(i, weight=1, uniform="col")
            col = tk.Frame(self.board_frame, bg=C["dark2"],
                           highlightbackground=C["ai_border"], highlightthickness=1)
            col.grid(row=0, column=i, sticky="nsew", padx=4, pady=4)
            hf = tk.Frame(col, bg=C["dark2"]); hf.pack(fill="x", padx=10, pady=(10,6))
            dot_c = [C["muted"],C["blue"],C["amber"],C["soft"]][i]
            tk.Label(hf, text="●", bg=C["dark2"], fg=dot_c, font=("Arial",8)).pack(side="left")
            tk.Label(hf, text=f"  {label}", bg=C["dark2"], fg=C["ai_text"],
                     font=("Arial",11,"bold")).pack(side="left")
            cv = tk.StringVar(value="0"); self.stat_vars[f"col_{key}"] = cv
            tk.Label(hf, textvariable=cv, bg=C["ai_surf"], fg=C["muted"],
                     font=("Arial",9), padx=6, pady=1,
                     highlightbackground=C["ai_border"], highlightthickness=1).pack(side="right")
            tk.Frame(col, bg=C["ai_border"], height=1).pack(fill="x", padx=8)
            inner = tk.Frame(col, bg=C["dark2"])
            inner.pack(fill="both", expand=True, padx=8, pady=8)
            if DND_AVAILABLE:
                inner.drop_target_register(DND_ALL)
                inner.dnd_bind("<<Drop>>", lambda e, k=key: self._on_drop(e, k))
            self.col_inner[key] = inner
            tk.Button(col, text="+ Add task", relief="flat", cursor="hand2",
                      bg=C["dark2"], fg=C["muted"], font=("Arial",9), pady=5, bd=0,
                      command=lambda k=key: self.open_task_dialog(default_status=k)
                      ).pack(fill="x", padx=8, pady=(0,8))

    def _build_ai_panel(self):
        self.ai_frame = tk.Frame(self.main_frame, bg=C["ai_bg"], width=310)
        self.ai_frame.pack(side="right", fill="y", padx=0)
        self.ai_frame.pack_propagate(False)

        # Input anchored to bottom FIRST
        bottom = tk.Frame(self.ai_frame, bg=C["ai_bg"])
        bottom.pack(side="bottom", fill="x")
        tk.Label(bottom,
                 text="FCG Grow  |  Built by Fairy Circle Garden  |  fairycirclegarden.com",
                 bg=C["ai_bg"], fg="#1e3828", font=("Arial",7)).pack(pady=(0,2))
        # Donate link
        don_lbl = tk.Label(bottom, text="♥ Support FCG Grow",
                           bg=C["ai_bg"], fg=C["ai_accent"],
                           font=("Arial",7,"italic"), cursor="hand2")
        don_lbl.pack(pady=(0,2))
        don_lbl.bind("<Button-1>", lambda e: _open_help_browser("about"))
        tk.Frame(bottom, bg=C["ai_border"], height=1).pack(fill="x", padx=8)
        inp = tk.Frame(bottom, bg=C["ai_bg"])
        inp.pack(fill="x", padx=8, pady=6)
        br = tk.Frame(inp, bg=C["ai_bg"]); br.pack(fill="x", pady=(0,4))
        self.send_btn = tk.Button(br, text="Send", relief="flat", cursor="hand2",
                                   bg=C["ai_accent"], fg=C["dark"],
                                   font=("Arial",10,"bold"), padx=12, pady=5, bd=0,
                                   command=self._send)
        self.send_btn.pack(side="right")
        tk.Button(br, text="Clear", relief="flat", cursor="hand2",
                  bg=C["ai_bg"], fg=C["muted"], font=("Arial",9),
                  padx=8, pady=5, bd=0, command=self._clear_chat).pack(side="right", padx=4)
        self.ai_input = tk.Text(inp, height=3, font=("Arial",10), relief="flat",
                                 bg=C["ai_surf"], fg=C["ai_text"],
                                 insertbackground=C["ai_text"],
                                 highlightbackground=C["ai_border"],
                                 highlightthickness=1, wrap="word")
        self.ai_input.pack(fill="x")
        self.ai_input.bind("<Return>", self._on_enter)
        add_copy_paste(self.ai_input)
        tk.Frame(bottom, bg=C["ai_border"], height=1).pack(fill="x", padx=8)

        # Gradient header
        ai_hdr_frame = tk.Frame(self.ai_frame, bg=C["ai_bg"], height=48)
        ai_hdr_frame.pack(fill="x"); ai_hdr_frame.pack_propagate(False)
        if PIL_AVAILABLE:
            try:
                img = make_ai_img(320, 48)
                if img:
                    self._ai_hdr_photo = pil_to_tk(img)
                    tk.Label(ai_hdr_frame, image=self._ai_hdr_photo, bd=0).place(
                        x=0, y=0, relwidth=1, relheight=1)
            except Exception: pass
        hdr = tk.Frame(ai_hdr_frame, bg="#050E0A")
        hdr.pack(fill="x", padx=10, pady=(10,4))
        org_name = self.cfg.get("org_name","Your Business")
        self.ai_org_lbl = tk.Label(hdr, text=org_name, bg="#050E0A", fg="white",
                                    font=("Arial",11,"bold"), wraplength=190)
        self.ai_org_lbl.pack(side="left")
        self.prov_lbl = tk.Label(hdr, text="", bg="#050E0A",
                                  fg=C["ai_accent"], font=("Arial",8))
        self.prov_lbl.pack(side="right")
        tk.Frame(self.ai_frame, bg=C["ai_border"], height=1).pack(fill="x", padx=8)

        # Quick actions -- collapsible
        self._qa_expanded = tk.BooleanVar(value=False)
        qa_hdr = tk.Frame(self.ai_frame, bg=C["ai_bg"])
        qa_hdr.pack(fill="x", padx=8, pady=(4,0))
        self._qa_btn = tk.Button(qa_hdr, text="Quick actions  ▶",
                                  relief="flat", cursor="hand2",
                                  bg=C["ai_bg"], fg=C["muted"], font=("Arial",8),
                                  bd=0, anchor="w", command=self._toggle_qa)
        self._qa_btn.pack(fill="x")
        self._qa_frame = tk.Frame(self.ai_frame, bg=C["ai_bg"])
        for label, prompt in [
            ("Prioritize today",  "What should I prioritize today?"),
            ("Show overdue",      "What tasks are overdue?"),
            ("Board summary",     "Give me a summary of my board"),
            ("Next steps",        "What are my most important next steps?"),
        ]:
            tk.Button(self._qa_frame, text=label, relief="flat", cursor="hand2",
                      bg=C["ai_surf"], fg=C["soft"], font=("Arial",9),
                      pady=3, padx=8, bd=0, anchor="w",
                      highlightbackground=C["ai_border"], highlightthickness=1,
                      command=lambda p=prompt: self._quick_ask(p)
                      ).pack(fill="x", pady=1, padx=8)
        tk.Frame(self.ai_frame, bg=C["ai_border"], height=1).pack(fill="x", padx=8, pady=4)

        # Chat display — manual Text + dark scrollbar (ScrolledText forces white scrollbar on Windows)
        chat_outer = tk.Frame(self.ai_frame, bg=C["ai_bg"])
        chat_outer.pack(fill="both", expand=True, padx=4)
        dark_vsb = ttk.Scrollbar(chat_outer, orient="vertical",
                                  style="Dark.Vertical.TScrollbar")
        dark_vsb.pack(side="right", fill="y")
        self.chat_display = tk.Text(
            chat_outer, bg=C["ai_bg"], fg=C["ai_text"],
            font=("Consolas", 9), relief="flat", wrap="word",
            state="disabled", padx=10, pady=8, highlightthickness=0,
            yscrollcommand=dark_vsb.set)
        self.chat_display.pack(side="left", fill="both", expand=True)
        dark_vsb.config(command=self.chat_display.yview)
        self.chat_display.tag_config("ul", foreground=C["ai_accent"], font=("Consolas",8,"bold"))
        self.chat_display.tag_config("ut", foreground="#e8f5e0", font=("Consolas",9))
        self.chat_display.tag_config("al", foreground=C["soft"], font=("Consolas",8,"bold"))
        self.chat_display.tag_config("at", foreground=C["ai_text"], font=("Consolas",9))
        self.chat_display.tag_config("ac", foreground=C["ai_accent"], font=("Consolas",8,"italic"))
        self.chat_display.tag_config("er", foreground=C["red"], font=("Consolas",9,"italic"))

    def _toggle_qa(self):
        if self._qa_expanded.get():
            self._qa_frame.pack_forget()
            self._qa_btn.config(text="Quick actions  ▶")
            self._qa_expanded.set(False)
        else:
            self._qa_frame.pack(fill="x", pady=(0,4))
            self._qa_btn.config(text="Quick actions  ▼")
            self._qa_expanded.set(True)

    def _show_file_menu(self):
        m = tk.Menu(self.root, tearoff=0, bg=C["dark2"], fg=C["ai_text"],
                    activebackground=C["mid"], activeforeground="white",
                    relief="flat", bd=1)
        m.add_command(label="Import Calendar",  command=self.import_calendar)
        m.add_command(label="Import Document",  command=self.import_document)
        m.add_command(label="Export Tasks",     command=self.export_tasks)
        m.add_separator()
        m.add_command(label="New Workspace",    command=self._new_ws)
        m.add_command(label="Delete Workspace", command=self._del_ws)
        m.add_separator()
        m.add_command(label="Exit",             command=self._on_close)
        try:
            m.tk_popup(self.root.winfo_rootx()+400, self.root.winfo_rooty()+54)
        finally:
            m.grab_release()

    def _show_view_menu(self):
        m = tk.Menu(self.root, tearoff=0, bg=C["dark2"], fg=C["ai_text"],
                    activebackground=C["mid"], activeforeground="white",
                    relief="flat", bd=1)
        m.add_command(label="Toggle AI Panel", command=self._toggle_ai)
        m.add_separator()
        for cat in ["All"] + CATEGORIES:
            m.add_command(label=f"Filter: {cat}", command=lambda c=cat: self.set_filter(c))
        try:
            m.tk_popup(self.root.winfo_rootx()+500, self.root.winfo_rooty()+54)
        finally:
            m.grab_release()

    def _show_tools_menu(self):
        m = tk.Menu(self.root, tearoff=0, bg=C["dark2"], fg=C["ai_text"],
                    activebackground=C["mid"], activeforeground="white",
                    relief="flat", bd=1)
        m.add_command(label="Settings",      command=self._open_settings)
        m.add_command(label="? Help",        command=self._open_help)
        m.add_separator()
        m.add_command(label="About FCG Grow", command=lambda: self._open_help("about"))
        m.add_command(label="♥ Support",     command=lambda: _open_help_browser("about"))
        try:
            m.tk_popup(self.root.winfo_rootx()+590, self.root.winfo_rooty()+54)
        finally:
            m.grab_release()

    def _toggle_ai(self):
        if self.ai_frame.winfo_ismapped():
            self.ai_frame.pack_forget()
        else:
            self.ai_frame.pack(side="right", fill="y", padx=0)

    def _build_statusbar(self):
        self.status_var = tk.StringVar(value="Ready  |  FCG Grow by Fairy Circle Garden")
        tk.Label(self.root, textvariable=self.status_var,
                 bg=C["dark"], fg=C["soft"], font=("Arial",9),
                 anchor="w", padx=10).pack(fill="x", side="bottom")

    # ── Render ─────────────────────────────────────────────────────────────
    def render_cols(self, *keys):
        """Render only the specified column keys. Fast path for drag/move ops."""
        q = self.search_var.get().lower()
        # Suspend layout updates until all columns are built
        self._board_canvas.update_idletasks()
        for key in keys:
            inner = self.col_inner.get(key)
            if inner is None: continue
            # Freeze the inner frame during rebuild to prevent flicker
            inner.config(bg=inner.cget("bg"))
            for w in inner.winfo_children(): w.destroy()
            col_tasks = [t for t in self.tasks if t["status"]==key]
            if self.filter_cat != "All":
                col_tasks = [t for t in col_tasks if t["category"]==self.filter_cat]
            if q:
                col_tasks = [t for t in col_tasks if q in t["title"].lower()
                             or q in t.get("notes","").lower()
                             or q in t.get("category","").lower()]
            for t in col_tasks: self._make_card(inner, t)
            real = len([t for t in self.tasks if t["status"]==key])
            if f"col_{key}" in self.stat_vars:
                self.stat_vars[f"col_{key}"].set(str(real))
        self._update_stats()

    def render_all(self):
        """Full board render. Use render_cols() for targeted updates."""
        self.render_cols(*COL_KEYS)

    def _make_card(self, parent, task):
        CARD_S = "#0f2a1a"
        CARD_B = "#1a3a28"
        card = tk.Frame(parent, bg=CARD_S,
                        highlightbackground=C["ai_border"], highlightthickness=1,
                        cursor="hand2")
        card.pack(fill="x", pady=3)
        pri_color = PRI_COLOR.get(task["priority"], C["muted"])
        tk.Frame(card, bg=pri_color, width=4).pack(side="left", fill="y")
        body = tk.Frame(card, bg=CARD_S)
        body.pack(side="left", fill="both", expand=True, padx=8, pady=7)

        tr = tk.Frame(body, bg=CARD_S); tr.pack(fill="x")
        dv = tk.BooleanVar(value=task.get("done",False))
        tk.Checkbutton(tr, variable=dv, bg=CARD_S,
                       activebackground=C["card_bg"], cursor="hand2",
                       command=lambda t=task, v=dv: self._toggle_done(t,v)).pack(side="left")
        tl = tk.Label(tr, text=task["title"], bg=CARD_S,
                      fg=C["muted"] if task.get("done") else C["ai_text"],
                      font=("Arial",10,"bold" if not task.get("done") else "normal"),
                      wraplength=185, justify="left", anchor="w")
        tl.pack(side="left", fill="x", expand=True)

        ft = tk.Frame(body, bg=CARD_S); ft.pack(fill="x", pady=(3,0))
        cat_colors = {"Grant":("#dbeafe","#1a6fa8"),"Operations":(C["pale"],C["mid"]),
                      "Admin":(C["amber_bg"],"#a06b10"),"Finance":(C["red_bg"],C["red"]),
                      "Marketing":("#f0ebfb","#6b4fa8"),"Outreach":("#e8f3fb",C["blue"]),
                      "Other":(C["pale"],C["muted"])}
        cat = task.get("category","Other")
        cbg, cfg2 = cat_colors.get(cat,(C["pale"],C["muted"]))
        cat_lbl = tk.Label(ft, text=cat,
                 bg=cbg, fg=cfg2, font=("Arial",8), padx=5, pady=1,
                 cursor="hand2")
        cat_lbl.pack(side="left")
        # Click badge → popup to reassign category in place
        def _show_cat_picker(e, t=task, lbl=cat_lbl):
            pop = tk.Toplevel(self.root)
            pop.wm_overrideredirect(True)
            pop.configure(bg=C["dark2"])
            # Position just below the badge
            lbl.update_idletasks()
            x = lbl.winfo_rootx()
            y = lbl.winfo_rooty() + lbl.winfo_height() + 2
            pop.wm_geometry(f"+{x}+{y}")
            # Header
            tk.Label(pop, text="Move to category:", bg=C["dark2"], fg=C["muted"],
                     font=("Arial",8), padx=8, pady=4).pack(fill="x")
            tk.Frame(pop, bg=C["ai_border"], height=1).pack(fill="x")
            def _pick(new_cat, p=pop):
                old_cat = t.get("category","Other")
                if new_cat != old_cat:
                    t["category"] = new_cat
                    save_ws_tasks(self.current_ws, self.tasks)
                    self.render_cols(t["status"])
                p.destroy()
            for c in CATEGORIES:
                cc = cat_colors.get(c,(C["pale"],C["muted"]))
                row = tk.Frame(pop, bg=C["dark2"], cursor="hand2")
                row.pack(fill="x")
                dot = tk.Label(row, text="●", bg=C["dark2"], fg=cc[1],
                               font=("Arial",9), padx=6, pady=3)
                dot.pack(side="left")
                lbl2 = tk.Label(row, text=c, bg=C["dark2"],
                                fg="white" if c != t.get("category") else C["ai_accent"],
                                font=("Arial",9,"bold" if c==t.get("category") else "normal"),
                                padx=4, pady=3, anchor="w")
                lbl2.pack(side="left", fill="x", expand=True)
                for w in (row, dot, lbl2):
                    w.bind("<Button-1>", lambda e, nc=c: _pick(nc))
                    w.bind("<Enter>", lambda e, r=row: r.config(bg=C["ai_surf"]))
                    w.bind("<Leave>", lambda e, r=row: r.config(bg=C["dark2"]))
            # Close on click outside
            pop.bind("<FocusOut>", lambda e: pop.after(100, lambda: pop.destroy() if pop.winfo_exists() else None))
            pop.focus_set()
        cat_lbl.bind("<Button-1>", _show_cat_picker)
        cat_lbl.bind("<Enter>", lambda e, l=cat_lbl: l.config(relief="solid", bd=1))
        cat_lbl.bind("<Leave>", lambda e, l=cat_lbl: l.config(relief="flat", bd=0))
        if task.get("due"):
            today = datetime.now().date().isoformat()
            urgent = task["due"] <= today and not task.get("done")
            tk.Label(ft, text=task["due"], bg=CARD_S,
                     fg=C["red"] if urgent else C["muted"],
                     font=("Arial",8,"bold" if urgent else "normal")).pack(side="right")

        # Notes preview
        if task.get("notes"):
            preview = task["notes"][:60] + ("..." if len(task["notes"])>60 else "")
            tk.Label(body, text=preview, bg=CARD_S, fg=C["muted"],
                     font=("Arial",8), wraplength=185, justify="left",
                     anchor="w").pack(fill="x", pady=(2,0))

        mr = tk.Frame(body, bg=CARD_S); mr.pack(fill="x", pady=(5,0))
        mv = tk.StringVar(value="Move to...")
        opts = [l for l in COL_LABELS if COL_KEYS[COL_LABELS.index(l)]!=task["status"]]
        mc = tk.OptionMenu(mr, mv, *opts)
        mc.config(relief="flat", bg=C["mid"], fg="white",
                  activebackground=C["dark2"], activeforeground="white",
                  highlightthickness=0, bd=0, font=("Arial",8,"bold"),
                  padx=4, pady=1)
        mc["menu"].config(bg=C["dark2"], fg=C["ai_text"],
                          activebackground=C["mid"], activeforeground="white",
                          font=("Arial",8))
        mc.pack(side="left")
        mv.trace_add("write", lambda *a, t=task, v=mv: self._move_task(t,v))
        tk.Button(mr, text="Edit", relief="flat", cursor="hand2",
                  bg=C["ai_surf"], fg=C["soft"], font=("Arial",8), padx=5, pady=1, bd=0,
                  command=lambda t=task: self.open_task_dialog(task=t)).pack(side="left", padx=3)
        tk.Button(mr, text="Del", relief="flat", cursor="hand2",
                  bg=C["ai_surf"], fg=C["red"], font=("Arial",8), padx=5, pady=1, bd=0,
                  command=lambda t=task: self._delete_task(t)).pack(side="left")

        # Mousewheel scoped to board
        def _ce(e): self._board_canvas.bind_all("<MouseWheel>",
            lambda ev: self._board_canvas.yview_scroll(-1*(ev.delta//120),"units"))
        def _cl(e): self._board_canvas.unbind_all("<MouseWheel>")
        for w in [card, body, tl]:
            w.bind("<Enter>", _ce); w.bind("<Leave>", _cl)

        # Mouse drag — press only; motion+release handled at root level in _mb_press
        for w in [card, body, tl]:
            w.bind("<ButtonPress-1>",   lambda e,t=task: self._mb_press(e,t))

        # DnD
        if DND_AVAILABLE:
            for w in [card, body, tl]:
                w.drag_source_register(DND_ALL)
                w.dnd_bind("<<DragInitCmd>>",
                           lambda e,tid=task["id"]: self._drag_start(e,tid))

    # ── Card drag ──────────────────────────────────────────────────────────
    def _drag_start(self, e, tid):
        self._mb_drag["id"] = tid
        return (DND_ALL, "copy", str(tid))

    def _on_drop(self, e, col):
        tid = self._mb_drag.get("id")
        if not tid: return
        task = next((t for t in self.tasks if t["id"]==tid), None)
        if task:
            old_col = task["status"]
            task["status"] = col
            if col=="done": task["done"] = True
            save_ws_tasks(self.current_ws, self.tasks)
            self.render_cols(old_col, col)
        self._mb_drag.clear()

    def _mb_press(self, e, task):
        self._mb_drag = {"task":task,"sx":e.x_root,"sy":e.y_root,"dragging":False,"ghost":None}
        # Bind motion and release to root so they fire even when cursor leaves the card
        self.root.bind("<B1-Motion>",       lambda ev,t=task: self._mb_motion(ev,t), add=False)
        self.root.bind("<ButtonRelease-1>", lambda ev,t=task: self._mb_release(ev,t), add=False)

    def _mb_motion(self, e, task):
        d = self._mb_drag
        if not d: return
        if abs(e.x_root-d["sx"])+abs(e.y_root-d["sy"]) > 3:
            d["dragging"] = True
            if not d.get("ghost"):
                g = tk.Label(self.root, text=task["title"][:28],
                             bg=C["dark2"], fg=C["ai_text"],
                             font=("Arial",9), padx=8, pady=4, relief="flat",
                             highlightbackground=C["ai_accent"], highlightthickness=1)
                g.place(x=e.x_root-self.root.winfo_rootx()+10,
                        y=e.y_root-self.root.winfo_rooty()+10)
                d["ghost"] = g
            else:
                d["ghost"].place(x=e.x_root-self.root.winfo_rootx()+10,
                                 y=e.y_root-self.root.winfo_rooty()+10)

    def _mb_release(self, e, task):
        # Unbind root-level handlers immediately
        try:
            self.root.unbind("<B1-Motion>")
            self.root.unbind("<ButtonRelease-1>")
        except Exception:
            pass
        d = self._mb_drag
        if not d: return
        if d.get("ghost"): d["ghost"].destroy()
        if d.get("dragging"):
            for key, inner in self.col_inner.items():
                cf = inner.master
                if (cf.winfo_rootx() <= e.x_root <= cf.winfo_rootx()+cf.winfo_width() and
                    cf.winfo_rooty() <= e.y_root <= cf.winfo_rooty()+cf.winfo_height()):
                    if key != task["status"]:
                        old_key = task["status"]
                        task["status"] = key
                        if key=="done": task["done"] = True
                        save_ws_tasks(self.current_ws, self.tasks)
                        # Only redraw the two affected columns — not the whole board
                        self.render_cols(old_key, key)
                        self._toast(f"Moved to {COL_LABELS[COL_KEYS.index(key)]}")
                    break
        elif not d.get("dragging"):
            self.open_task_dialog(task=task)
        self._mb_drag = {}

    # ── Actions ────────────────────────────────────────────────────────────
    def _toggle_done(self, task, var):
        old_key = task["status"]
        task["done"] = var.get()
        if task["done"]: task["status"] = "done"
        save_ws_tasks(self.current_ws, self.tasks)
        if old_key != task["status"]:
            self.render_cols(old_key, "done")
        else:
            self.render_cols(old_key)

    def _move_task(self, task, var):
        label = var.get()
        if label not in COL_LABELS: return
        old_key = task["status"]
        new_key = COL_KEYS[COL_LABELS.index(label)]
        if new_key == old_key: return
        task["status"] = new_key
        if new_key=="done": task["done"] = True
        var.set("Move to...")
        save_ws_tasks(self.current_ws, self.tasks)
        self.render_cols(old_key, new_key)
        self._toast(f"Moved to {label}")

    def _delete_task(self, task):
        if messagebox.askyesno("Delete", f"Delete '{task['title']}'?"):
            col = task["status"]
            self.tasks = [t for t in self.tasks if t["id"]!=task["id"]]
            save_ws_tasks(self.current_ws, self.tasks)
            self.render_cols(col)

    def set_filter(self, cat):
        self.filter_cat = cat
        for c, btn in self.filter_buttons.items():
            if c == cat:
                btn.config(bg=C["ai_accent"], fg=C["dark"],
                           font=("Arial",9,"bold"), relief="flat")
            else:
                btn.config(bg=C["ai_surf"], fg=C["ai_accent"],
                           font=("Arial",9,"normal"), relief="flat")
        self.render_all()

    def _update_stats(self):
        self.stat_vars["total"].set(str(len(self.tasks)))
        self.stat_vars["todo"].set(str(len([t for t in self.tasks if t["status"]=="todo"])))
        self.stat_vars["inprog"].set(str(len([t for t in self.tasks if t["status"]=="inprog"])))
        self.stat_vars["active"].set(str(len([t for t in self.tasks if not t.get("done")])))
        self.stat_vars["done"].set(str(len([t for t in self.tasks if t.get("done")])))

    # ── Task dialog ────────────────────────────────────────────────────────
    def open_task_dialog(self, task=None, default_status="todo"):
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.title("Edit task" if task else "Add task")
        win.geometry("500x560"); win.configure(bg=C["dark2"])
        win.resizable(False, False)
        _apply_dark_titlebar(win)
        win.deiconify()
        win.grab_set()
        icon_path = os.path.join(BASE_DIR, "fcggrow.ico")
        if os.path.exists(icon_path):
            try: win.iconbitmap(icon_path)
            except Exception: pass

        # Gradient header
        hdr = tk.Frame(win, bg=C["dark"], height=50)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        if PIL_AVAILABLE:
            try:
                img = make_topbar_img(500, 50)
                if img:
                    ph = pil_to_tk(img); win._hdr_photo = ph
                    tk.Label(hdr, image=ph, bd=0).place(x=0,y=0,relwidth=1,relheight=1)
            except Exception: pass
        tk.Label(hdr, text="Edit task" if task else "New task",
                 bg=C["dark"], fg="white",
                 font=("Arial",13,"bold"), pady=14).pack()

        fields = {}
        def row(label, wfn, default=""):
            f = tk.Frame(win, bg=C["dark2"]); f.pack(fill="x", padx=22, pady=5)
            tk.Label(f, text=label, bg=C["dark2"], fg=C["muted"],
                     font=("Arial",9,"bold"), width=12, anchor="w").pack(side="left")
            w = wfn(f); w.pack(side="left", fill="x", expand=True)
            if default and hasattr(w,"insert"): w.insert(0,default)
            elif default and hasattr(w,"set"): w.set(default)
            return w

        def entry(p):
            e = tk.Entry(p, font=("Arial",10), relief="flat",
                         bg=C["ai_surf"], fg=C["ai_text"],
                         insertbackground=C["ai_text"],
                         highlightbackground=C["ai_border"], highlightthickness=1)
            add_copy_paste(e); return e
        def combo(opts):
            def _make(p):
                v = tk.StringVar()
                om = tk.OptionMenu(p, v, *opts)
                om.config(relief="flat", bg=C["ai_surf"], fg=C["ai_text"],
                          activebackground=C["mid"], activeforeground="white",
                          highlightthickness=1, highlightbackground=C["ai_border"],
                          bd=0, font=("Arial",10), anchor="w")
                om["menu"].config(bg=C["dark2"], fg=C["ai_text"],
                                  activebackground=C["mid"], activeforeground="white",
                                  font=("Arial",10))
                # give it a .get() and .set() so row() default works
                om.get = v.get; om.set = v.set
                return om
            return _make

        fields["title"]    = row("Title",    entry, task["title"] if task else "")
        fields["status"]   = row("Status",   combo(COL_LABELS),
                                  COL_LABELS[COL_KEYS.index(task["status"])] if task
                                  else COL_LABELS[COL_KEYS.index(default_status)])
        fields["priority"] = row("Priority", combo(PRIORITIES),
                                  task["priority"] if task else "Medium")
        fields["category"] = row("Category", combo(CATEGORIES),
                                  task["category"] if task else "Operations")
        fields["due"]      = row("Due date", entry, task["due"] if task else "")

        nf = tk.Frame(win, bg=C["dark2"]); nf.pack(fill="x", padx=22, pady=5)
        tk.Label(nf, text="Notes", bg=C["dark2"], fg=C["muted"],
                 font=("Arial",9,"bold"), width=12, anchor="w").pack(side="left", anchor="n")
        nw = tk.Text(nf, height=5, font=("Arial",10), relief="flat",
                     bg=C["ai_surf"], fg=C["ai_text"],
                     insertbackground=C["ai_text"],
                     highlightbackground=C["ai_border"], highlightthickness=1)
        nw.pack(side="left", fill="x", expand=True)
        add_copy_paste(nw)
        if task and task.get("notes"): nw.insert("1.0", task["notes"])
        fields["notes"] = nw

        tk.Frame(win, bg=C["ai_border"], height=1).pack(fill="x", padx=22, pady=6)

        def do_save():
            title = fields["title"].get().strip()
            if not title:
                messagebox.showwarning("Required","Please enter a title.",parent=win); return
            sl = fields["status"].get()
            sk = COL_KEYS[COL_LABELS.index(sl)] if sl in COL_LABELS else "todo"
            cat = fields["category"].get() or "Operations"
            data = {"title":title,"status":sk,
                    "priority":fields["priority"].get() or "Medium",
                    "category":cat,"due":fields["due"].get().strip(),
                    "notes":fields["notes"].get("1.0","end").strip(),
                    "done":sk=="done"}
            if task: task.update(data)
            else: data["id"] = next_id(self.tasks); self.tasks.append(data)
            save_ws_tasks(self.current_ws, self.tasks); self.render_all()
            win.destroy(); self._toast("Task saved")

        bf = tk.Frame(win, bg=C["dark2"]); bf.pack(pady=4)
        tk.Button(bf, text="Cancel", command=win.destroy, relief="flat",
                  bg=C["ai_surf"], fg=C["ai_text"], font=("Arial",10),
                  padx=14, pady=6, cursor="hand2").pack(side="left", padx=5)
        tk.Button(bf, text="Save task", command=do_save, relief="flat",
                  bg=C["dark"], fg="white", font=("Arial",10,"bold"),
                  padx=14, pady=6, cursor="hand2").pack(side="left", padx=5)

    # ── Import calendar ────────────────────────────────────────────────────
    def import_calendar(self):
        win = tk.Toplevel(self.root)
        win.title("Import calendar"); win.geometry("520x300")
        win.configure(bg=C["dark2"]); win.grab_set(); win.resizable(False,False)
        win.after(50, lambda: _apply_dark_titlebar(win))
        icon_path = os.path.join(BASE_DIR, "fcggrow.ico")
        if os.path.exists(icon_path):
            try: win.iconbitmap(icon_path)
            except Exception: pass

        tk.Label(win, text="Import calendar events as tasks",
                 bg=C["dark2"], fg=C["ai_text"], font=("Arial",13,"bold")).pack(pady=(18,4))
        tk.Frame(win, bg=C["ai_border"], height=1).pack(fill="x", padx=22, pady=4)
        tk.Label(win, text="Paste a public .ics calendar URL:",
                 bg=C["dark2"], fg=C["muted"], font=("Arial",9)).pack(anchor="w", padx=22)
        url_var = tk.StringVar()
        url_e = tk.Entry(win, textvariable=url_var, font=("Arial",10), relief="flat",
                         bg=C["ai_surf"], fg=C["ai_text"],
                         insertbackground=C["ai_text"],
                         highlightbackground=C["ai_border"], highlightthickness=1)
        url_e.pack(fill="x", padx=22, pady=6, ipady=6)
        add_copy_paste(url_e)
        sv = tk.StringVar()
        tk.Label(win, textvariable=sv, bg=C["dark2"], fg=C["soft"], font=("Arial",9)).pack()

        def from_file():
            path = filedialog.askopenfilename(title="Open .ics file",
                filetypes=[("iCalendar","*.ics"),("All","*.*")])
            if path:
                with open(path,"rb") as f: _do_import(f.read())

        def from_url():
            url = url_var.get().strip()
            if not url:
                messagebox.showwarning("URL needed","Enter a calendar URL.",parent=win); return
            # Fix Google Calendar embed URLs
            if "calendar/embed" in url:
                url = url.replace("calendar/embed?src=","calendar/ical/")
                for tz in ["&ctz=America%2FChicago","&ctz=America/Chicago","&ctz=UTC"]:
                    url = url.replace(tz,"/public/basic.ics")
                if not url.endswith(".ics"):
                    url += "/public/basic.ics"
            sv.set("Fetching..."); win.update()
            def fetch():
                try:
                    req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0"})
                    with urllib.request.urlopen(req, timeout=15) as r: data = r.read()
                    win.after(0, lambda: _do_import(data))
                except Exception as e:
                    win.after(0, lambda: sv.set(f"Error: {e}"))
            threading.Thread(target=fetch, daemon=True).start()

        def _do_import(data):
            if not ICAL_AVAILABLE:
                messagebox.showerror("Missing","pip install icalendar",parent=win); return
            events = parse_ics(data)
            if not events: sv.set("No events found."); return
            added = 0
            for ev in events:
                if any(t["title"]==ev["title"] for t in self.tasks): continue
                cat = auto_categorize(ev["title"], ev.get("notes",""))
                self.tasks.append({"id":next_id(self.tasks),"title":ev["title"],
                    "status":"todo","priority":"Medium","category":cat,
                    "due":ev.get("due",""),"notes":ev.get("notes","Imported from calendar"),
                    "done":False})
                added += 1
            save_ws_tasks(self.current_ws, self.tasks); self.render_all()
            sv.set(f"Done. Added {added} task(s) with auto-categorization.")
            self._toast(f"Imported {added} calendar event(s)")

        bf = tk.Frame(win, bg=C["dark2"]); bf.pack(pady=10)
        tk.Button(bf, text="Load .ics file", command=from_file, relief="flat",
                  bg=C["ai_surf"], fg=C["ai_text"], font=("Arial",10),
                  padx=12, pady=6, cursor="hand2").pack(side="left", padx=5)
        tk.Button(bf, text="Import from URL", command=from_url, relief="flat",
                  bg=C["dark"], fg="white", font=("Arial",10,"bold"),
                  padx=12, pady=6, cursor="hand2").pack(side="left", padx=5)

    # ── Import document ────────────────────────────────────────────────────
    def import_document(self):
        path = filedialog.askopenfilename(
            title="Import document -- AI will extract tasks",
            filetypes=[("Documents","*.pdf *.docx *.txt"),
                       ("PDF","*.pdf"),("Word","*.docx"),
                       ("Text","*.txt"),("All","*.*")])
        if not path: return
        text, err = extract_doc_text(path)
        if err:
            messagebox.showerror("Import error", err, parent=self.root); return
        if not text or not text.strip():
            messagebox.showwarning("Empty document",
                "No text could be extracted from this document.", parent=self.root); return
        fname = os.path.basename(path)
        prompt = (f"I've imported a document called '{fname}'. "
                  f"Please read it and create tasks on my board for every action item, "
                  f"requirement, or deadline you find. "
                  f"Auto-categorize each task appropriately. "
                  f"Here is the document text:\n\n{text}")
        self.ai_input.delete("1.0","end")
        self.ai_input.insert("1.0", prompt)
        self._send()
        self._toast(f"Processing document: {fname}")

    # ── Export ─────────────────────────────────────────────────────────────
    def export_tasks(self):
        path = filedialog.asksaveasfilename(
            title="Export tasks", defaultextension=".json",
            initialfile=f"FCGGrow_{self.current_ws}.json",
            filetypes=[("JSON","*.json"),("All","*.*")])
        if not path: return
        with open(path,"w") as f: json.dump(self.tasks, f, indent=2)
        self._toast(f"Exported to {os.path.basename(path)}")

    # ── AI ──────────────────────────────────────────────────────────────────
    def _ai_welcome(self):
        org   = self.cfg.get("org_name","your organization")
        naics = self.cfg.get("naics","")
        biz   = self.cfg.get("biz_type","")
        lines = [f"Ready -- {org}"]
        if biz:   lines.append(f"Type: {biz}")
        if naics: lines.append(f"NAICS: {naics}")
        if get_naics_intel(naics): lines.append("Industry intelligence: loaded")
        lines += ["","Try asking:",
                  "  What should I prioritize today?",
                  "  Find all grant tasks and recategorize them",
                  "  Add a task: [description] by [date]",
                  "  What is overdue?",
                  "  Move [task] to In Progress",
                  "  How do I win a government contract?"]
        self._append("FCG Grow", "\n".join(lines), "ai")

    def _quick_ask(self, prompt):
        self.ai_input.delete("1.0","end")
        self.ai_input.insert("1.0", prompt)
        self._send()

    def _on_enter(self, e):
        if not (e.state & 0x1): self._send(); return "break"

    def _send(self):
        if self.ai_busy: return
        msg = self.ai_input.get("1.0","end").strip()
        if not msg: return
        providers = self.cfg.get("providers",[])
        if not providers:
            self._append("System","No AI providers configured. Open Settings.","er"); return

        # Check for pending bulk confirmation
        if self._pending_bulk_actions and msg.lower().strip() in ("confirm","yes","y"):
            actions = self._pending_bulk_actions
            self._pending_bulk_actions = None
            self.ai_input.delete("1.0","end")
            self._append("You","confirm","user")
            results = []
            for tool in actions:
                r = self._execute_tool(tool)
                if r: results.append(r)
            if results:
                action_str = f"{len(results)} actions completed"
                self._append("Action", action_str, "ac")
                save_ws_tasks(self.current_ws, self.tasks); self.render_all()
            self._ai_done(); return

        self.ai_input.delete("1.0","end")
        self._append("You", msg, "user")
        self.ai_busy = True
        self.send_btn.config(state="disabled", text="...")
        self._toast("AI thinking...")

        system_prompt = build_system_prompt(self.cfg)
        board_context = build_board_context(self.tasks)
        full_msg = f"{board_context}\n\nUser: {msg}"
        messages = []
        for h in self.chat_history[-20:]:
            messages.append({"role":h["role"],"content":h["content"]})
        messages.append({"role":"user","content":full_msg})

        def worker():
            try:
                reply, used = call_ai(providers, system_prompt, messages)
                self.chat_history.append({"role":"user","content":msg})
                self.chat_history.append({"role":"assistant","content":reply})
                self.root.after(0, lambda r=reply,u=used: self._handle_reply(r,u))
            except Exception as e:
                self.root.after(0, lambda err=str(e): self._append("Error",err,"er"))
                self.root.after(0, self._ai_done)

        threading.Thread(target=worker, daemon=True).start()

    def _handle_reply(self, reply, used):
        self.active_provider = used
        # Show provider + active model in header
        providers = self.cfg.get("providers", [])
        model_name = ""
        for p in providers:
            if p.get("name","") == used.split("(")[0].strip():
                m = p.get("model","")
                model_name = m.split("/")[-1] if "/" in m else m
                break
        display = f"{used.split('(')[0].strip()}"
        if model_name:
            display = model_name[:28]
        self.prov_lbl.config(text=display)

        # Strip all JSON blocks from display
        display = re.sub(r"```json.*?```", "", reply, flags=re.DOTALL).strip()
        display = re.sub(r"```\s*$", "", display, flags=re.MULTILINE).strip()

        # Extract all JSON blocks
        json_blocks = re.findall(r"```json\s*(.*?)```", reply, flags=re.DOTALL)
        all_tools = []
        for block in json_blocks:
            block = block.strip()
            try:
                raw = json.loads(block)
                tools = raw if isinstance(raw, list) else [raw]
                for t in tools:
                    if isinstance(t,dict) and "action" in t:
                        all_tools.append(t)
            except Exception:
                # Try finding individual objects
                for obj in re.findall(r"\{[^{}]+\}", block):
                    try:
                        t = json.loads(obj)
                        if isinstance(t,dict) and "action" in t:
                            all_tools.append(t)
                    except Exception: pass

        # Fallback -- scan entire reply
        if not all_tools:
            for obj in re.findall(r"\{[^{}]+\}", reply):
                try:
                    t = json.loads(obj)
                    if isinstance(t,dict) and "action" in t:
                        all_tools.append(t)
                except Exception: pass

        # Bulk delete safety check
        delete_actions = [t for t in all_tools if t.get("action")=="delete_task"]
        if len(delete_actions) > 5:
            self._pending_bulk_actions = all_tools
            self._append("FCG Grow", display, "ai")
            self._append("FCG Grow",
                f"This will {len(delete_actions)} delete tasks. Reply 'confirm' to proceed.",
                "ac")
            self._ai_done(); return

        self._append("FCG Grow", display, "ai")

        if all_tools:
            results = []
            for tool in all_tools:
                r = self._execute_tool(tool)
                if r: results.append(r)
            if results:
                if len(results) == 1: action_str = results[0]
                elif len(results) <= 3: action_str = f"{len(results)} actions: " + ", ".join(results)
                else: action_str = f"{len(results)} actions: {results[0]} ... +{len(results)-1} more"
                self._append("Action", action_str, "ac")
                save_ws_tasks(self.current_ws, self.tasks); self.render_all()  # AI may touch any col

        self._ai_done()

    def _execute_tool(self, tool):
        act = tool.get("action")
        if act == "add_task":
            title = tool.get("title","New task")
            sk    = tool.get("status","todo")
            if sk not in COL_KEYS: sk = "todo"
            cat   = tool.get("category","")
            if not cat or cat not in CATEGORIES:
                cat = auto_categorize(title, tool.get("notes",""))
            self.tasks.append({"id":next_id(self.tasks),"title":title,"status":sk,
                "priority":tool.get("priority","Medium"),"category":cat,
                "due":tool.get("due",""),"notes":tool.get("notes","Added by AI"),
                "done":sk=="done"})
            return f"Added: '{title}' [{cat}]"

        elif act == "update_task":
            tid  = tool.get("task_id")
            task = next((t for t in self.tasks if t["id"]==tid), None)
            if task:
                changed = []
                for field in ["category","priority","title","due","notes","status"]:
                    if field in tool and tool[field]:
                        if field == "category" and tool[field] not in CATEGORIES:
                            continue
                        task[field] = tool[field]
                        changed.append(field)
                return f"Updated '{task['title']}': {', '.join(changed)}"
            return f"Task {tid} not found"

        elif act == "move_task":
            tid = tool.get("task_id"); ns = tool.get("status","todo")
            if ns not in COL_KEYS: return f"Unknown status: {ns}"
            task = next((t for t in self.tasks if t["id"]==tid), None)
            if task:
                old = COL_LABELS[COL_KEYS.index(task["status"])]
                task["status"] = ns
                if ns=="done": task["done"] = True
                return f"Moved '{task['title']}': {old} to {COL_LABELS[COL_KEYS.index(ns)]}"
            return f"Task {tid} not found"

        elif act == "delete_task":
            tid  = tool.get("task_id")
            task = next((t for t in self.tasks if t["id"]==tid), None)
            if task:
                title = task["title"]
                self.tasks = [t for t in self.tasks if t["id"]!=tid]
                return f"Deleted '{title}'"
            return f"Task {tid} not found"

        return None

    def _ai_done(self):
        self.ai_busy = False
        self.send_btn.config(state="normal", text="Send")
        self._toast("Ready  |  FCG Grow by Fairy Circle Garden")

    def _append(self, speaker, text, kind="ai"):
        self.chat_display.config(state="normal")
        self.chat_display.insert("end","\n")
        if kind=="user":
            self.chat_display.insert("end","[ You ]\n","ul")
            self.chat_display.insert("end",f"{text}\n","ut")
        elif kind in ("ac","action"):
            self.chat_display.insert("end",f"  {text}\n","ac")
        elif kind in ("er","error"):
            self.chat_display.insert("end",f"  {text}\n","er")
        else:
            self.chat_display.insert("end","[ FCG Grow ]\n","al")
            self.chat_display.insert("end",f"{text}\n","at")
        self.chat_display.config(state="disabled")
        self.chat_display.see("end")

    def _clear_chat(self):
        self.chat_history = []
        self.chat_display.config(state="normal")
        self.chat_display.delete("1.0","end")
        self.chat_display.config(state="disabled")
        self._ai_welcome()

    def _toast(self, msg):
        self.status_var.set(f"  {msg}")
        if not self.ai_busy:
            self.root.after(3000, lambda: self.status_var.set(
                "  Ready  |  FCG Grow by Fairy Circle Garden"))

# ── Entry point ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    FCGGrowApp()
