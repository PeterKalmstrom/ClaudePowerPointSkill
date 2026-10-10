"""Bundled line icons (Lucide, ISC licence - assets/icons/LICENSE) drawn as editable PowerPoint vector shapes.

Each icon becomes ONE native shape with a custom geometry (a:custGeom): its SVG strokes become path outlines in the
theme accent, so the icon is crisp at any size, recolours with the theme, can be edited with Edit Points, and renders
the same in PowerPoint and LibreOffice. Why not a picture: a PNG needs a rasteriser (none in the standard toolchain)
and blurs when scaled; an SVG picture shows only in PowerPoint 2016+ (needs a PNG fallback); EMF renders unevenly in
LibreOffice. Plain stroked custGeom is the one form both renderers draw identically.

    kIconLibrary.Resolve("warning")        -> "triangle-alert"   (names, aliases)
    kIconLibrary.Guess("Report phishing")  -> "mail"             (keyword choice)
    kIconShape.Draw(Shapes, "shield", Left, Top, Size, MSO_THEME_COLOR.ACCENT_1, "Icon1")
"""
import math
import os
import re

from lxml import etree
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu

from kShared import kS

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ICON_DIR = os.path.join(ROOT, "assets", "icons")
SVG = "{http://www.w3.org/2000/svg}"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
DECORATIVE_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}"
DECORATIVE_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"
UNITS = 1000  # path units per SVG unit: the 24-unit Lucide grid becomes a 24000-unit path
VIEW = 24.0
STROKE = 2.0  # Lucide's stroke width on its 24-unit grid
KAPPA = 0.5522847498  # cubic Bezier handle length for a quarter circle

ALIASES = {
    "money": "banknote", "cash": "banknote", "cost": "coins", "budget": "wallet", "savings": "piggy-bank",
    "warning": "triangle-alert", "alert": "triangle-alert", "risk": "triangle-alert", "error": "circle-alert",
    "people": "users", "team": "users", "person": "user", "email": "mail", "e-mail": "mail",
    "chart": "chart-column", "bar-chart": "chart-column", "line-chart": "chart-line", "pie": "chart-pie",
    "growth": "trending-up", "decline": "trending-down", "time": "clock", "deadline": "timer",
    "security": "shield", "secure": "shield-check", "password": "key-round", "key": "key-round",
    "unlock": "lock-open", "ok": "circle-check", "done": "circle-check", "tick": "check", "cross": "x",
    "close": "x", "date": "calendar", "call": "phone", "goal": "goal", "idea": "lightbulb", "launch": "rocket",
    "world": "globe", "office": "building-2", "company": "building-2", "work": "briefcase",
    "training": "graduation-cap", "learn": "book-open", "document": "file-text", "file": "file-text",
    "data": "database", "settings": "settings", "tool": "wrench", "fast": "zap", "like": "thumbs-up",
    "chat": "message-square", "announce": "megaphone", "notify": "bell", "cart": "shopping-cart",
    "delivery": "truck", "location": "map-pin", "partner": "handshake", "legal": "scale", "law": "gavel",
    "checklist": "list-checks", "award": "award", "win": "trophy", "help": "circle-help", "question": "circle-help",
    "stop": "ban", "process": "workflow", "support": "headset", "ai": "bot", "robot": "bot", "mobile": "smartphone",
}

# Keyword choice, most specific first: (icon, word stems). A word matches a stem when it starts with it (stems of
# four or more letters) or equals it (shorter stems), so "phishing" matches "phish" but "ai" never matches "aim".
KEYWORDS = [
    ("mail", ("phish", "email", "e-mail", "mail", "inbox", "newsletter")),
    ("key-round", ("password", "passkey", "credential", "mfa", "2fa")),
    ("fingerprint", ("biometric", "fingerprint", "identity")),
    ("lock", ("lock", "encrypt", "privacy", "private", "confidential")),
    ("shield-check", ("compliance", "compliant", "certif", "audit")),
    ("shield", ("secur", "protect", "defen", "threat", "attack", "malware", "ransom", "breach")),
    ("bug", ("bug", "defect", "vulnerab")),
    ("triangle-alert", ("risk", "warning", "danger", "hazard", "incident")),
    ("bot", ("ai", "automat", "robot", "chatbot", "copilot", "agent")),
    ("brain", ("think", "brain", "cognit", "mindset")),
    ("graduation-cap", ("train", "course", "learn", "student", "teach", "lesson", "workshop")),
    ("book-open", ("read", "book", "guide", "manual", "handbook")),
    ("piggy-bank", ("saving", "save")),
    ("banknote", ("revenue", "income", "profit", "sales", "cash", "money", "price", "pricing", "payment", "pay")),
    ("coins", ("cost", "spend", "expense", "budget", "fund", "invest", "capex", "opex")),
    ("percent", ("percent", "margin", "discount", "rate")),
    ("trending-up", ("growth", "grow", "increase", "rise", "improv", "gain", "up")),
    ("trending-down", ("decline", "decrease", "drop", "fall", "reduc", "churn", "down")),
    ("chart-column", ("metric", "kpi", "measur", "report", "dashboard", "analytic", "data", "statistic")),
    ("users", ("team", "people", "staff", "employee", "customer", "client", "user", "member", "stakeholder",
               "hiring", "hire", "recruit", "talent", "hr", "headcount", "fte")),
    ("handshake", ("partner", "deal", "agree", "contract", "negotiat", "vendor", "supplier")),
    ("headset", ("support", "helpdesk", "service", "ticket")),
    ("message-square", ("feedback", "chat", "message", "comment", "talk", "discuss", "communicat")),
    ("megaphone", ("announc", "campaign", "marketing", "launch", "promot", "brand")),
    ("phone", ("phone", "call", "vish", "smish", "sms")),
    ("calendar", ("calendar", "schedul", "date", "week", "month", "quarter", "year", "annual", "meeting", "plan")),
    ("clock", ("time", "hour", "minute", "day", "daily", "late", "delay", "wait", "speed", "fast", "quick")),
    ("target", ("target", "aim", "focus", "objective", "okr")),
    ("goal", ("goal", "mission", "vision", "strategy")),
    ("rocket", ("pilot", "start", "kickoff", "rollout", "deploy", "ship", "release", "scale")),
    ("lightbulb", ("idea", "insight", "innovat", "tip", "suggest", "lesson")),
    ("circle-check", ("approv", "done", "complet", "success", "pass", "verify", "confirm", "check", "accept")),
    ("ban", ("block", "ban", "stop", "reject", "deny", "never")),
    ("settings", ("config", "setting", "setup", "process", "operat", "system")),
    ("wrench", ("fix", "repair", "maintain", "maintenance", "tool")),
    ("cloud", ("cloud", "saas", "azure", "aws", "online")),
    ("database", ("database", "storage", "backup", "record", "archive")),
    ("server", ("server", "infrastructure", "network", "hosting")),
    ("laptop", ("laptop", "computer", "device", "desktop", "endpoint", "software", "app")),
    ("globe", ("global", "world", "international", "market", "region", "web", "internet")),
    ("building-2", ("office", "company", "organi", "corporate", "headquarter", "site", "facility")),
    ("factory", ("factory", "manufactur", "production", "plant")),
    ("store", ("store", "shop", "retail")),
    ("truck", ("deliver", "logistic", "shipping", "transport", "supply")),
    ("scale", ("legal", "law", "regulat", "gdpr", "policy", "fair", "ethic")),
    ("file-text", ("document", "file", "doc", "paper", "form", "invoice", "proposal")),
    ("trophy", ("win", "award", "best", "champion", "reward", "recogni")),
    ("star", ("quality", "rating", "star", "excellen", "premium")),
    ("heart", ("health", "wellbeing", "care", "love", "patient")),
    ("leaf", ("sustainab", "green", "climate", "environment", "carbon", "eco")),
    ("refresh-cw", ("renew", "repeat", "refresh", "update", "iterat", "cycle", "review")),
    ("search", ("search", "find", "investigat", "research", "discover", "explor")),
    ("eye", ("visib", "monitor", "watch", "observ", "spot", "notice")),
    ("bell", ("alert", "notif", "remind", "alarm")),
    ("link", ("link", "url", "integrat", "connect")),
    ("layers", ("layer", "stack", "platform", "architect")),
    ("puzzle", ("puzzle", "fit", "piece", "module", "component")),
    ("flag", ("milestone", "flag", "finish", "deadline", "phase")),
    ("map-pin", ("location", "place", "local", "where")),
    ("user", ("owner", "person", "individual", "manager", "lead", "ceo", "cfo", "cto", "ciso")),
]


class kIconLibrary:
    """The bundled icon set: names, aliases and the keyword choice."""

    _names = None

    @staticmethod
    def Names():
        """Every bundled icon name, sorted (natural order)."""
        if kS.ErrorMode:
            return []
        try:
            if kIconLibrary._names is None:
                Found = [F[:-4] for F in os.listdir(ICON_DIR) if F.endswith(".svg")] if os.path.isdir(ICON_DIR) else []
                kIconLibrary._names = sorted(Found)
            return list(kIconLibrary._names)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconLibrary.Names")
            return []

    @staticmethod
    def Resolve(Name):
        """The bundled icon a spec name means (an icon name or an alias, any case); None when there is none."""
        if kS.ErrorMode:
            return None
        try:
            Key = str(Name or "").strip().lower().replace("_", "-").replace(" ", "-")
            Key = ALIASES.get(Key, Key)
            return Key if Key in kIconLibrary.Names() else None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconLibrary.Resolve")
            return None

    @staticmethod
    def Off(Name):
        """True when a spec icon value switches the icon off for that item ("none", "", false)."""
        if kS.ErrorMode:
            return False
        try:
            return Name is False or str(Name).strip().lower() in ("", "none", "off", "false")
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconLibrary.Off")
            return False

    @staticmethod
    def WordMatches(Word, Stem):
        """True when a word carries a keyword stem (prefix for stems of 4+ letters, whole word otherwise)."""
        if kS.ErrorMode:
            return False
        try:
            return Word.startswith(Stem) if len(Stem) >= 4 else Word == Stem
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconLibrary.WordMatches")
            return False

    @staticmethod
    def Guess(Text):
        """The icon the words of Text suggest (first match in KEYWORDS order), or None."""
        if kS.ErrorMode:
            return None
        try:
            Words = re.findall(r"[a-z0-9][a-z0-9\-]*", str(Text or "").lower())
            for Icon, Stems in KEYWORDS:
                for Word in Words:
                    for Stem in Stems:
                        if kIconLibrary.WordMatches(Word, Stem):
                            return Icon
            return None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconLibrary.Guess")
            return None

    @staticmethod
    def AutoSet(Texts):
        """Icons for a set of items chosen from their words - only when EVERY item gets one and no two share
        one (a half-iconed or repetitive row reads worse than none); otherwise []."""
        if kS.ErrorMode:
            return []
        try:
            Icons = [kIconLibrary.Guess(T) for T in Texts]
            if not Icons or None in Icons or len(set(Icons)) != len(Icons):
                return []
            return Icons
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconLibrary.AutoSet")
            return []


class kSvgIcon:
    """One bundled SVG read into absolute subpaths: lists of ("M", x, y), ("L", x, y), ("C", x1, y1, x2, y2, x, y)
    and ("Z",) on the 24-unit grid. Handles path (M L H V C S Q T A Z, absolute and relative), circle, ellipse,
    rect (with rx), line, polyline and polygon - everything Lucide uses."""

    _cache = {}

    @staticmethod
    def Load(Name):
        """The subpaths of icon Name (cached); [] when it cannot be read."""
        if kS.ErrorMode:
            return []
        try:
            if Name not in kSvgIcon._cache:
                Root = etree.parse(os.path.join(ICON_DIR, f"{Name}.svg"), etree.XMLParser(resolve_entities=False,
                                                                                         no_network=True)).getroot()
                Paths = []
                for El in Root.iter():
                    if isinstance(El.tag, str):
                        Paths.extend(kSvgIcon.Element(El))
                kSvgIcon._cache[Name] = Paths
            return kSvgIcon._cache[Name]
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kSvgIcon.Load(name={Name})")
            return []

    @staticmethod
    def Num(El, Key, Default=0.0):
        """A numeric attribute."""
        if kS.ErrorMode:
            return Default
        try:
            Val = El.get(Key)
            return float(Val) if Val not in (None, "") else Default
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Num")
            return Default

    @staticmethod
    def Element(El):
        """The subpaths one SVG element draws."""
        if kS.ErrorMode:
            return []
        try:
            Tag, N = El.tag.replace(SVG, ""), kSvgIcon.Num
            if Tag == "path":
                return kSvgIcon.PathData(El.get("d", ""))
            if Tag == "circle":
                return [kSvgIcon.Ellipse(N(El, "cx"), N(El, "cy"), N(El, "r"), N(El, "r"))]
            if Tag == "ellipse":
                return [kSvgIcon.Ellipse(N(El, "cx"), N(El, "cy"), N(El, "rx"), N(El, "ry"))]
            if Tag == "line":
                return [[("M", N(El, "x1"), N(El, "y1")), ("L", N(El, "x2"), N(El, "y2"))]]
            if Tag in ("polyline", "polygon"):
                V = [float(X) for X in re.findall(r"-?(?:\d+\.?\d*|\.\d+)(?:e-?\d+)?", El.get("points", ""))]
                Pts = list(zip(V[0::2], V[1::2]))
                if not Pts:
                    return []
                Sub = [("M",) + Pts[0]] + [("L",) + P for P in Pts[1:]]
                return [Sub + ([("Z",)] if Tag == "polygon" else [])]
            if Tag == "rect":
                return [kSvgIcon.Rect(N(El, "x"), N(El, "y"), N(El, "width"), N(El, "height"),
                                      N(El, "rx", N(El, "ry")), N(El, "ry", N(El, "rx")))]
            return []
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Element")
            return []

    @staticmethod
    def Ellipse(Cx, Cy, Rx, Ry):
        """A closed ellipse as four cubic quarter arcs."""
        if kS.ErrorMode:
            return []
        try:
            Kx, Ky = Rx * KAPPA, Ry * KAPPA
            return [("M", Cx + Rx, Cy),
                    ("C", Cx + Rx, Cy + Ky, Cx + Kx, Cy + Ry, Cx, Cy + Ry),
                    ("C", Cx - Kx, Cy + Ry, Cx - Rx, Cy + Ky, Cx - Rx, Cy),
                    ("C", Cx - Rx, Cy - Ky, Cx - Kx, Cy - Ry, Cx, Cy - Ry),
                    ("C", Cx + Kx, Cy - Ry, Cx + Rx, Cy - Ky, Cx + Rx, Cy), ("Z",)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Ellipse")
            return []

    @staticmethod
    def Rect(X, Y, Wd, Ht, Rx, Ry):
        """A closed rectangle, its corners rounded by Rx/Ry."""
        if kS.ErrorMode:
            return []
        try:
            Rx, Ry = min(Rx, Wd / 2), min(Ry, Ht / 2)
            if Rx <= 0 or Ry <= 0:
                return [("M", X, Y), ("L", X + Wd, Y), ("L", X + Wd, Y + Ht), ("L", X, Y + Ht), ("Z",)]
            Kx, Ky = Rx * (1 - KAPPA), Ry * (1 - KAPPA)
            R, B = X + Wd, Y + Ht
            return [("M", X + Rx, Y), ("L", R - Rx, Y), ("C", R - Kx, Y, R, Y + Ky, R, Y + Ry),
                    ("L", R, B - Ry), ("C", R, B - Ky, R - Kx, B, R - Rx, B),
                    ("L", X + Rx, B), ("C", X + Kx, B, X, B - Ky, X, B - Ry),
                    ("L", X, Y + Ry), ("C", X, Y + Ky, X + Kx, Y, X + Rx, Y), ("Z",)]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Rect")
            return []

    @staticmethod
    def Tokens(D):
        """SVG path data as a list of command letters and number strings (arc flags may run together)."""
        if kS.ErrorMode:
            return []
        try:
            return re.findall(r"[MmLlHhVvCcSsQqTtAaZz]|-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?", D)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Tokens")
            return []

    @staticmethod
    def ArcArgs(Toks, I):
        """Seven arc arguments from position I (flags may be written '01' with no separator); (args, next I)."""
        if kS.ErrorMode:
            return [], len(Toks)
        try:
            Args = []
            while len(Args) < 7 and I < len(Toks):
                T = Toks[I]
                if len(Args) in (3, 4) and len(T) > 1 and T[0] in "01" and not T.startswith(("0.", "1.")):
                    Args.append(float(T[0]))
                    Toks[I] = T[1:]
                    continue
                Args.append(float(T))
                I += 1
            return Args, I
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.ArcArgs")
            return [], len(Toks)

    @staticmethod
    def PathData(D):
        """The subpaths of an SVG path's d attribute, in absolute coordinates."""
        if kS.ErrorMode:
            return []
        try:
            Toks, I, Cmd = kSvgIcon.Tokens(D), 0, None
            St = {"x": 0.0, "y": 0.0, "sx": 0.0, "sy": 0.0, "cx": None, "cy": None, "q": None, "paths": [], "cur": None}
            while I < len(Toks):
                if Toks[I].isalpha():
                    Cmd, I = Toks[I], I + 1
                    if Cmd in "Zz":
                        kSvgIcon.Step(St, "Z", [])
                        continue
                Need = {"M": 2, "L": 2, "T": 2, "H": 1, "V": 1, "C": 6, "S": 4, "Q": 4, "A": 7}[Cmd.upper()]
                if Cmd.upper() == "A":
                    Args, I = kSvgIcon.ArcArgs(Toks, I)
                else:
                    Args, I = [float(X) for X in Toks[I:I + Need]], I + Need
                if len(Args) < Need:
                    break
                kSvgIcon.Step(St, Cmd, Args)
                if Cmd in "Mm":
                    Cmd = "l" if Cmd == "m" else "L"  # extra pairs after a moveto are linetos
            return [P for P in St["paths"] if len(P) > 1]
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.PathData")
            return []

    @staticmethod
    def Step(St, Cmd, Args):
        """Apply one path command to the parser state St (current point, last control points, subpaths)."""
        if kS.ErrorMode:
            return
        try:
            Rel, Up = Cmd.islower(), Cmd.upper()
            X0, Y0 = St["x"], St["y"]
            Dx, Dy = (X0, Y0) if Rel else (0.0, 0.0)
            Ctrl, Quad = None, None
            if Up == "Z":
                if St["cur"] is not None:
                    St["cur"].append(("Z",))
                St["x"], St["y"] = St["sx"], St["sy"]
                St["cur"] = None
            elif Up == "M":
                St["x"], St["y"] = Args[0] + Dx, Args[1] + Dy
                St["sx"], St["sy"] = St["x"], St["y"]
                St["cur"] = [("M", St["x"], St["y"])]
                St["paths"].append(St["cur"])
            else:
                if St["cur"] is None:  # drawing after a close: a new subpath from the start point
                    St["cur"] = [("M", X0, Y0)]
                    St["paths"].append(St["cur"])
                Ctrl, Quad = kSvgIcon.Draw(St, Up, Args, X0, Y0, Dx, Dy)
            St["cx"], St["cy"] = (Ctrl if Ctrl else (None, None))
            St["q"] = Quad
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Step")
            return

    @staticmethod
    def Draw(St, Up, Args, X0, Y0, Dx, Dy):
        """A drawing command (L H V C S Q T A) onto the current subpath; returns (last cubic control, last
        quadratic control) for the S and T commands that follow."""
        if kS.ErrorMode:
            return None, None
        try:
            Cur = St["cur"]
            if Up in ("L", "H", "V"):
                X = Args[0] + Dx if Up in ("L", "H") else X0
                Y = (Args[1] if Up == "L" else Args[0]) + Dy if Up in ("L", "V") else Y0
                Cur.append(("L", X, Y))
                St["x"], St["y"] = X, Y
                return None, None
            if Up in ("C", "S"):
                if Up == "C":
                    X1, Y1, A = Args[0] + Dx, Args[1] + Dy, Args[2:]
                else:
                    X1, Y1 = (2 * X0 - St["cx"], 2 * Y0 - St["cy"]) if St["cx"] is not None else (X0, Y0)
                    A = Args
                X2, Y2, X, Y = A[0] + Dx, A[1] + Dy, A[2] + Dx, A[3] + Dy
                Cur.append(("C", X1, Y1, X2, Y2, X, Y))
                St["x"], St["y"] = X, Y
                return (X2, Y2), None
            if Up in ("Q", "T"):
                if Up == "Q":
                    Qx, Qy, X, Y = Args[0] + Dx, Args[1] + Dy, Args[2] + Dx, Args[3] + Dy
                else:
                    Qx, Qy = (2 * X0 - St["q"][0], 2 * Y0 - St["q"][1]) if St["q"] else (X0, Y0)
                    X, Y = Args[0] + Dx, Args[1] + Dy
                Cur.append(("C", X0 + 2 / 3 * (Qx - X0), Y0 + 2 / 3 * (Qy - Y0), X + 2 / 3 * (Qx - X),
                            Y + 2 / 3 * (Qy - Y), X, Y))
                St["x"], St["y"] = X, Y
                return None, (Qx, Qy)
            X, Y = Args[5] + Dx, Args[6] + Dy  # A
            Cur.extend(kSvgIcon.Arc(X0, Y0, Args[0], Args[1], Args[2], Args[3], Args[4], X, Y))
            St["x"], St["y"] = X, Y
            return None, None
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Draw")
            return None, None

    @staticmethod
    def Arc(X1, Y1, Rx, Ry, Phi, Large, Sweep, X2, Y2):
        """An SVG elliptical arc as cubic segments (each at most a quarter turn), per the SVG implementation notes."""
        if kS.ErrorMode:
            return []
        try:
            if Rx == 0 or Ry == 0 or (X1 == X2 and Y1 == Y2):
                return [("L", X2, Y2)]
            Rx, Ry, Ph = abs(Rx), abs(Ry), math.radians(Phi)
            Co, Si = math.cos(Ph), math.sin(Ph)
            Xp = Co * (X1 - X2) / 2 + Si * (Y1 - Y2) / 2
            Yp = -Si * (X1 - X2) / 2 + Co * (Y1 - Y2) / 2
            L = Xp * Xp / (Rx * Rx) + Yp * Yp / (Ry * Ry)
            if L > 1:
                Rx, Ry = Rx * math.sqrt(L), Ry * math.sqrt(L)
            Num = Rx * Rx * Ry * Ry - Rx * Rx * Yp * Yp - Ry * Ry * Xp * Xp
            Den = Rx * Rx * Yp * Yp + Ry * Ry * Xp * Xp
            K = math.sqrt(max(0.0, Num / Den)) * (-1 if Large == Sweep else 1)
            Cxp, Cyp = K * Rx * Yp / Ry, -K * Ry * Xp / Rx
            Cx = Co * Cxp - Si * Cyp + (X1 + X2) / 2
            Cy = Si * Cxp + Co * Cyp + (Y1 + Y2) / 2
            T1 = math.atan2((Yp - Cyp) / Ry, (Xp - Cxp) / Rx)
            T2 = math.atan2((-Yp - Cyp) / Ry, (-Xp - Cxp) / Rx)
            Dt = T2 - T1
            if Sweep and Dt < 0:
                Dt += 2 * math.pi
            elif not Sweep and Dt > 0:
                Dt -= 2 * math.pi
            return kSvgIcon.ArcCubics(Cx, Cy, Rx, Ry, Co, Si, T1, Dt)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.Arc")
            return []

    @staticmethod
    def ArcCubics(Cx, Cy, Rx, Ry, Co, Si, T1, Dt):
        """Cubic segments approximating the arc from angle T1 through Dt on a rotated ellipse."""
        if kS.ErrorMode:
            return []
        try:
            N = max(1, int(math.ceil(abs(Dt) / (math.pi / 2) - 1e-9)))
            D = Dt / N
            Al = 4 / 3 * math.tan(D / 4)
            Out = []
            for I in range(N):
                A, B = T1 + I * D, T1 + (I + 1) * D
                P = [(math.cos(A) - Al * math.sin(A), math.sin(A) + Al * math.cos(A)),
                     (math.cos(B) + Al * math.sin(B), math.sin(B) - Al * math.cos(B)),
                     (math.cos(B), math.sin(B))]
                Seg = []
                for Ux, Uy in P:
                    Seg.extend([Cx + Co * Rx * Ux - Si * Ry * Uy, Cy + Si * Rx * Ux + Co * Ry * Uy])
                Out.append(("C",) + tuple(Seg))
            return Out
        except Exception as e:
            kS.GlobalErrorHandler(e, "kSvgIcon.ArcCubics")
            return []


class kIconShape:
    """Draws a bundled icon as one native, editable vector shape (custom geometry, stroked in a theme colour)."""

    @staticmethod
    def Draw(Shapes, Name, Left, Top, Size, Colour, ShapeName, Alt=None):
        """Icon Name as a shape Size EMU square at (Left, Top) EMU, stroked in theme colour Colour. Marked
        decorative unless Alt is given (an icon beside its own text adds no information). None when the icon is
        unknown or has no geometry."""
        if kS.ErrorMode:
            return None
        try:
            Icon = kIconLibrary.Resolve(Name)
            Paths = kSvgIcon.Load(Icon) if Icon else []
            if not Paths:
                return None
            Sh = Shapes.add_shape(MSO_SHAPE.RECTANGLE, int(Left), int(Top), int(Size), int(Size))
            Sh.name = ShapeName
            Geom = Sh._element.spPr.find(f"{{{A_NS}}}prstGeom")
            Geom.addprevious(kIconShape.CustGeom(Paths))
            Geom.getparent().remove(Geom)
            Sh.fill.background()
            Sh.line.color.theme_color = Colour
            Sh.line.width = Emu(max(1, int(Size * STROKE / VIEW)))
            kIconShape.RoundStroke(Sh)
            Sh.shadow.inherit = False
            Ref = Sh._element.find(f".//{{{A_NS}}}effectRef")
            if Ref is not None:
                Ref.set("idx", "0")
            kIconShape.Describe(Sh, Alt, Icon)
            return Sh
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kIconShape.Draw(icon={Name})")
            return None

    @staticmethod
    def CustGeom(Paths):
        """The a:custGeom element for a list of subpaths (each its own unfilled a:path on the 24000 grid)."""
        if kS.ErrorMode:
            return None
        try:
            Q = "{%s}" % A_NS
            Geom = etree.Element(Q + "custGeom")
            for Tag in ("avLst", "gdLst", "ahLst", "cxnLst"):
                etree.SubElement(Geom, Q + Tag)
            etree.SubElement(Geom, Q + "rect", l="0", t="0", r="r", b="b")
            Lst = etree.SubElement(Geom, Q + "pathLst")
            Side = str(int(VIEW * UNITS))
            for Sub in Paths:
                Path = etree.SubElement(Lst, Q + "path", w=Side, h=Side, fill="none")
                for Seg in Sub:
                    kIconShape.Segment(Path, Seg)
            return Geom
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconShape.CustGeom")
            return None

    @staticmethod
    def Segment(Path, Seg):
        """One subpath segment as DrawingML (moveTo / lnTo / cubicBezTo / close)."""
        if kS.ErrorMode:
            return
        try:
            Q = "{%s}" % A_NS
            Tag = {"M": "moveTo", "L": "lnTo", "C": "cubicBezTo", "Z": "close"}[Seg[0]]
            El = etree.SubElement(Path, Q + Tag)
            Vals = Seg[1:]
            for I in range(0, len(Vals), 2):
                etree.SubElement(El, Q + "pt", x=str(int(round(Vals[I] * UNITS))), y=str(int(round(Vals[I + 1] * UNITS))))
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconShape.Segment")
            return

    @staticmethod
    def RoundStroke(Sh):
        """Round caps and joins, as the icon set is drawn."""
        if kS.ErrorMode:
            return
        try:
            Ln = Sh.line._get_or_add_ln()
            Ln.set("cap", "rnd")
            for Old in Ln.findall(f"{{{A_NS}}}round") + Ln.findall(f"{{{A_NS}}}miter") + Ln.findall(f"{{{A_NS}}}bevel"):
                Ln.remove(Old)
            Join = etree.Element(f"{{{A_NS}}}round")
            Fill = Ln.find(f"{{{A_NS}}}solidFill")
            Dash = Ln.find(f"{{{A_NS}}}prstDash")
            Anchor = Dash if Dash is not None else Fill
            if Anchor is not None:
                Anchor.addnext(Join)
            else:
                Ln.insert(0, Join)
        except Exception as e:
            kS.GlobalErrorHandler(e, "kIconShape.RoundStroke")
            return

    @staticmethod
    def Describe(Sh, Alt, Icon):
        """Alt text when given; otherwise the PowerPoint 'Mark as decorative' flag (screen readers skip it)."""
        if kS.ErrorMode:
            return
        try:
            Nv = Sh._element.nvSpPr.cNvPr
            if Alt:
                Nv.set("descr", str(Alt))
                return
            Nv.set("descr", "")
            Ext = etree.SubElement(etree.SubElement(Nv, f"{{{A_NS}}}extLst"), f"{{{A_NS}}}ext", uri=DECORATIVE_URI)
            etree.SubElement(Ext, f"{{{DECORATIVE_NS}}}decorative", nsmap={"adec": DECORATIVE_NS}, val="1")
        except Exception as e:
            kS.GlobalErrorHandler(e, f"kIconShape.Describe(icon={Icon})")
            return
