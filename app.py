# -*- coding: utf-8 -*-
"""
DRACORE - Dashboard Radiology Assurance, Compliance & Operational Reliability Excellence
Mandaya Royal Hospital Puri  |  Streamlit app (app.py)

Menu besar:  DRACORE (Struktur Menu, Dashboard, Rekap)  |  PMI  |  PME
Jalankan  :  streamlit run app.py
Syarat    :  streamlit>=1.36  pandas  openpyxl
Data input tersimpan di  data/entries.csv  (lampiran di data/uploads/).
Sidebar "Tanggal simulasi" + "Tampilkan data dummy" dipakai untuk preview;
matikan data dummy untuk melihat dashboard hanya dari data yang benar-benar diinput.
"""
import calendar
import datetime as dt
import io
import json
import os
import random
import zlib

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="DRACORE - Mandaya Royal Hospital Puri", page_icon="🏥", layout="wide")

DATA_DIR = "data"
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
ENTRY_FILE = os.path.join(DATA_DIR, "entries.csv")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# =====================================================================
# 1. STYLE
# =====================================================================
st.markdown("""
<style>
.stApp{background:#f1f4f9}
.block-container{padding-top:1.4rem;max-width:100%}
[data-testid=stSidebar]{background:#ffffff}
button[kind="primary"]{background:#235b9c;border-color:#235b9c}
.dc-bc{background:#f3f6fa;border-radius:6px;padding:7px 14px;font-size:12px;color:#667;margin-bottom:10px}
.dc-h{font-size:22px;font-weight:700;color:#1f2937;margin-bottom:2px}
.dc-sub{font-size:13px;color:#667;margin-bottom:14px}
.dc-tag{display:inline-block;padding:2px 10px;border-radius:10px;font-size:11px;font-weight:700;color:#fff;margin-left:10px;vertical-align:middle}
.dc-hint{font-size:12px;margin:-8px 0 8px 2px}.dc-hint.ok{color:#2e9e5b}.dc-hint.bad{color:#d64545}
.dc-pn{background:#f6f8fb;border:1px solid #e1e7ef;border-radius:8px;padding:12px 14px;font-size:12.5px;line-height:1.75;margin-bottom:12px}
.dc-pn b{display:block;margin-bottom:3px}.dc-pn.gr{background:#f1faf4;border-color:#bfe3cb}
.dc-pn.rd{background:#fff4f4;border-color:#f2c0c0}.dc-pn.am{background:#fff9e8;border-color:#efd68f}
.dc-cat{background:#cfe3f7;font-weight:700;text-align:center;border-radius:4px;padding:6px;margin:10px 0 6px;font-size:13px}
.dc-ban{background:#fff8e1;border:1px dashed #e0b84a;color:#7a5b00;border-radius:6px;padding:7px 12px;font-size:12px;margin-bottom:12px}
</style>
""", unsafe_allow_html=True)

DASH_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
html,body{overflow-x:hidden;background:transparent}
body{font-family:'DejaVu Sans',Arial,sans-serif;color:#1f2937;font-size:12px}
.wrap{padding:2px 2px 6px}
.card{background:#fff;border-radius:8px;padding:14px 14px;margin-bottom:12px;box-shadow:0 1px 2px #0001}
.info{display:flex}.info div{flex:1;font-size:11px;color:#555}.info b{display:block;font-size:13px;color:#111;margin-top:3px}
.g{background:#2e9e5b}.y{background:#f0a81c}.r{background:#d64545}.b{background:#2a7ca8}.n{background:#e5e9ef}
.kpis{display:flex;gap:12px;margin-bottom:12px}.kpi{flex:1;background:#fff;border-radius:8px;padding:12px 14px;border-left:5px solid #2a7ca8;box-shadow:0 1px 2px #0001}
.kpi .t{font-size:11px;color:#666}.kpi .v{font-size:24px;font-weight:bold;margin:2px 0}.kpi .s{font-size:10.5px;color:#666}
.bar{height:6px;border-radius:3px;background:#e5e9ef;margin-top:6px;overflow:hidden;display:flex}.bar i{display:block;height:100%}
.sec{background:#fff;border-radius:8px;margin-bottom:12px;overflow:hidden;box-shadow:0 1px 2px #0001}
.sh{background:#c9defa;padding:11px 14px;font-weight:bold;font-size:12.5px;display:flex;justify-content:space-between}
.tabs{padding:12px 18px 6px;font-size:11.5px;color:#2a6fd6}.tabs span{margin-right:22px}.tabs .a{color:#111}
.tw{padding:4px 14px 14px;overflow-x:auto}.tw table{min-width:1100px}
table{border-collapse:collapse;width:100%}td,th{border:1px solid #dfe5ec;padding:5px 6px;text-align:center;font-size:10.5px}
th{background:#d9ead7;color:#244;font-weight:600}th.h{background:#5b8f6e;color:#fff}td.l,th.l{text-align:left}
td.c{color:#fff;font-weight:bold}
.pill{display:inline-block;padding:3px 9px;border-radius:10px;color:#fff;font-size:10px;font-weight:bold}
.note{font-size:10px;color:#777;padding:0 14px 10px}
.ban{background:#fff8e1;border:1px dashed #e0b84a;color:#7a5b00;border-radius:6px;padding:7px 12px;font-size:11px;margin-bottom:12px}
.stats{display:flex;gap:10px;margin-bottom:10px}.stat{flex:1;background:#f6f8fb;border-radius:6px;padding:8px 12px}
.stat .l{font-size:10.5px;color:#666}.stat .v{font-size:18px;font-weight:bold}
.legend{font-size:11px;line-height:19px;text-align:right}
.dot{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:6px;vertical-align:-2px}
.big{display:flex;gap:18px;margin-bottom:4px}.bc2{flex:1;background:#fff;border-radius:12px;box-shadow:0 2px 6px #0002;overflow:hidden}
.bh{padding:16px 18px;color:#fff}.bh h2{font-size:19px}.bh div{font-size:11.5px;opacity:.92;margin-top:3px}
.it{display:flex;align-items:center;gap:12px;padding:11px 18px;border-bottom:1px solid #edf0f4}
.ic{width:34px;height:34px;border-radius:9px;display:flex;align-items:center;justify-content:center;font-size:16px;color:#fff;flex-shrink:0}
.it .t{font-weight:bold;font-size:12.5px}.it .a{font-size:10.5px;color:#667;margin-top:2px}
.fq{margin-left:auto;text-align:right;font-size:10px}.fq span{display:inline-block;padding:3px 9px;border-radius:10px;font-weight:bold}
.bdg{display:inline-block;margin-left:6px;background:#d64545;color:#fff;border-radius:9px;padding:2px 7px;font-size:10px;font-weight:bold}
.sh2{display:flex;gap:12px;margin-top:16px}.sm{flex:1;background:#fff;border-radius:10px;padding:13px 14px;box-shadow:0 1px 3px #0002;border-top:4px solid #8899aa}
.sm b{display:block;font-size:12.5px;margin-bottom:3px}.sm div{font-size:10.5px;color:#667;line-height:1.5}
.fl{display:flex;align-items:center;margin-top:16px;background:#fff;border-radius:10px;padding:14px;box-shadow:0 1px 3px #0002}
.fl div.st{flex:1;text-align:center;font-size:11px}.fl div.st b{display:block;font-size:12px;margin-bottom:2px}.fl div.ar{color:#99a;font-size:20px}
"""

AUTOSIZE = """<script>
function fit(){try{var h=Math.ceil(document.body.getBoundingClientRect().height)+8;
window.frameElement.style.height=h+'px';}catch(e){}}
window.addEventListener('load',fit);
try{new ResizeObserver(fit).observe(document.body);}catch(e){}
</script>"""


def render(body, height=600):
    """Tampilkan blok HTML (dashboard) persis seperti mockup; tinggi iframe menyesuaikan otomatis."""
    doc = ("<html><head><meta charset='utf-8'><style>" + DASH_CSS + "</style></head><body><div class='wrap'>"
           + body + "</div>" + AUTOSIZE + "</body></html>")
    components.html(doc, height=height, scrolling=True)


# =====================================================================
# 2. DATA MASTER (dari dokumen Change Feature Request DRACORE)
# =====================================================================
BLN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
BLN3 = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]
STAFF = ["RR", "SN", "WF", "ND", "AL", "AG", "WN", "NA", "RN"]
HX = {"g": "#2e9e5b", "y": "#f0a81c", "r": "#d64545", "b": "#2a7ca8"}
SYM = {"g": "&#10003;", "y": "!", "r": "&#10007;"}
CL = {"Valid": "g", "Due Soon": "y", "Overdue": "r", "Not Compliant": "r", "Aktif": "g", "Warning": "y", "Expired": "r",
      "Need Evaluation": "r", "Terjadwal": "b", "Completed": "g", "Pass": "g", "Failed": "r", "Pending": "b",
      "Compliant": "g", "Need Follow-up": "y", "Belum Ada Data": "b"}

T_LO, T_HI, H_LO, H_HI = 18.0, 23.0, 40.0, 60.0
AREAS = ["R. Teknik CT", "R. Teknik MRI", "Lemari ALKES", "Lemari CD"]
SHIFTS = ["P - Pagi", "S - Siang", "M - Malam"]

DAILY = [("PHILIPS DIGITALDIAGNOST C50", "Suhu & QC Harian"), ("PHILIPS COMBIDIAGNOST R90", "Suhu & QC Harian"),
         ("PHILIPS MOBILEDIAGNOST WDR", "Suhu & QC Harian"), ("PHILIPS IQON SPECTRAL CT", "Suhu & QC Harian"),
         ("FUJIFILM DIGITAL MAMMOGRAPHY FDR", "Suhu & QC Harian"), ("PHILIPS MRI INGENIA AMBITION X 1,5 T", "Suhu & QC Harian"),
         ("PHILIPS USG AFINITY 70 G", "Suhu & QC Harian"), ("R. Teknik CT", "Suhu"), ("R. Teknik MRI", "Suhu"),
         ("Lemari ALKES", "Suhu"), ("Lemari CD", "Suhu")]
DAILY_MODALITIES = [x[0] for x in DAILY[:7]]

MRI_CHECKLIST = [
    ("Cek Kelengkapan & Peralatan Penunjang", [("Control box / panel listrik", "Lampu indikator menyala"), ("Oksigen central", "Menyala / Berfungsi"), ("AC", "Menyala / Berfungsi"), ("Pintu ruangan", "Dapat terkunci rapat"), ("Suhu Chiller", "Menyala dan suhu ada pada batas aman"), ("Tes publik sistem", "Terdengar secara jelas")]),
    ("Pemeriksaan Injektor", [("Injektor ON", "Injektor dan CT Scan terhubung tanpa kendala koneksi"), ("Cek Fungsi Syringe", "Dapat mengisi & membuang cairan untuk kedua channel syringe"), ("Tes publik sistem", "Terdengar secara jelas")]),
    ("Warming Up MRI", [("Pesawat & gantry ON", "Tidak terdapat notifikasi error dan booting secara normal"), ("Check up / Warming up Scan Phantom Botol", "Test scan phantom botol"), ("Monitor workstation dan ISP", "Menyala / berfungsi")]),
    ("Kondisi Ruangan", [("Kebersihan ruangan", "Ruangan bersih dan tidak terdapat sampah"), ("Kebersihan meja control & meja komputer", "Tidak terdapat debu & noda"), ("Suhu ruangan", "18 - 22 C"), ("Kelembapan Udara", "40 - 60 %"), ("Ketersediaan linen, bantal dan tissue gulung", "Tersedia secara lengkap"), ("Kelengkapan ALKES untuk tindakan", "Tersedia secara lengkap"), ("Selimut", "1 buah"), ("Bantal", "1 buah")]),
    ("Mematikan Pesawat", [("Matikan Injector", "Tekan switch pada monitor injector"), ("Pilih menu shutdown", "Pilih menu shutdown dan matikan sistem setiap malam hari")]),
    ("Kondisi Ruang Ganti / Locker Room 1 dan 2", [("Kelengkapan baju dan celana pasien", "baju dan celana tersedia untuk seluruh ukuran"), ("Kebersihan Loker", "kondisi loker pasien bersih dan tidak terdapat sampah"), ("Kondisi smart lock loker", "Kondisi loker bisa dibuka dan ditutup menggunakan gelang"), ("Kondisi pintu loker", "Kondisi pintu loker bisa terkunci dan tertutup rapat"), ("Kelengkapan jumlah gelang smart lock", "Satu changing room memiliki 4 buah smart lock loker")]),
]
# Checklist alat lain: template sederhana. Ganti/tambah sesuai form Daily QC masing-masing alat.
DEFAULT_CHECKLIST = [
    ("Cek Kelengkapan & Peralatan Penunjang", [("Control box / panel listrik", "Lampu indikator menyala"), ("AC", "Menyala / Berfungsi"), ("Pintu ruangan", "Dapat terkunci rapat"), ("Tes publik sistem", "Terdengar secara jelas")]),
    ("Pemeriksaan Alat", [("Pesawat ON & booting", "Tidak terdapat notifikasi error dan booting secara normal"), ("QC harian alat", "Hasil QC sesuai batas penerimaan")]),
    ("Kondisi Ruangan", [("Kebersihan ruangan", "Ruangan bersih dan tidak terdapat sampah"), ("Suhu ruangan", "18 - 22 C"), ("Kelembapan Udara", "40 - 60 %")]),
]
CHECKLISTS = {"PHILIPS MRI INGENIA AMBITION X 1,5 T": MRI_CHECKLIST}
NUMERIC = {"Suhu ruangan": (18.0, 22.0, "°C", 21.0), "Kelembapan Udara": (40.0, 60.0, "%", 52.0)}

WEEKLY_ROWS = [("CT Scan (iQon)", "CT Number, Uniformity, Noise, Low Contrast, Mono E, Air Calibration"), ("Digital Diagnost", "Homogenity"),
               ("Combi Diagnost", "Homogenity"), ("Mobile Diagnost", "Homogenity")]
WEEKLY_SPEC = {"CT Scan (iQon)": ["CT Number", "Uniformity", "Noise", "Low Contrast (mm)", "Mono E", "Air Calibration"],
               "Digital Diagnost": ["Homogenity"], "Combi Diagnost": ["Homogenity"], "Mobile Diagnost": ["Homogenity"]}

_KOL = ["Uji Intensitas Cahaya Kolimasi", "Uji Akurasi Berkas Kolimasi", "Ketegalurusan Berkas"]
# (parameter, batas) ; batas = (min, max, satuan, teks) bila numerik. Batas CT di bawah hanya ILUSTRASI -> sesuaikan dengan protokol QC internal.
BULANAN_SPEC = {
    "CT Scan": [("Uji Akurasi Pergerakan Meja", None), ("Linearitas CT Number - Water", (-5, 5, "HU", "0 ± 5 HU")), ("Linearitas CT Number - Nylon", None),
                ("Linearitas CT Number - Acrylic", (110, 130, "HU", "120 ± 10 HU")), ("Linearitas CT Number - Polyethylene", (-105, -85, "HU", "-95 ± 10 HU")),
                ("Linearitas CT Number - Teflon", (940, 1040, "HU", "990 ± 50 HU")), ("Linearitas CT Number - Lexan", None),
                ("Linearitas CT Number - Center Teflon Pin", None), ("Linearitas CT Number - Water Hole", None)],
    "Digital Diagnost": [(x, None) for x in _KOL + ["Gain Calibration", "Pixel Calibration", "Dynamic Range", "High Contrast Resolution", "Spatial Contrast Resolution"]],
    "Combi Diagnost": [(x, None) for x in _KOL + ["Dynamic Range", "High Contrast Resolution", "Spatial Contrast Resolution"]],
    "Mobile Diagnost WDR": [(x, None) for x in _KOL + ["Dynamic Range", "High Contrast Resolution", "Spatial Contrast Resolution"]],
    "Mammografi": [(x, None) for x in ["Missed tissue at chest wall side", "CNR", "System sensitivity", "Geometric distortion", "System artifact (the whole image)", "Uniformity", "Dynamic range", "Spatial resolution", "LCD", "Linearity/Beam quality"]],
    "MRI": [(x, None) for x in ["Flood Field Uniformity", "Spatial Linearity", "Slice Profile", "Spatial Resolution"]],
}
BULANAN_ROWS = [(k, ", ".join(p for p, _ in v)) for k, v in BULANAN_SPEC.items()]

LEAK_ITEMS = [("CT Scan", "Uji Kebocoran Tabung X-Ray + Radiasi", dt.date(2026, 3, 12)), ("Digital Diagnost", "Kebocoran Tabung X-Ray + Radiasi", dt.date(2026, 3, 14)),
              ("Combi Diagnost", "Kebocoran Tabung X-Ray + Radiasi", dt.date(2026, 3, 15)), ("wDR Mobile", "Kebocoran Tabung X-Ray", dt.date(2025, 11, 20)),
              ("Mammografi", "Kebocoran Tabung X-Ray", dt.date(2026, 4, 2)), ("Azurion Cathlab", "Kebocoran Radiasi", dt.date(2025, 10, 8)),
              ("C-Arm Ruang OT", "Kebocoran Radiasi", dt.date(2026, 1, 22)), ("BMD", "Kebocoran Radiasi", dt.date(2025, 9, 18)),
              ("Dental Clinic", "Kebocoran Radiasi", dt.date(2026, 5, 6))]

PM_ITEMS = [("PHILIPS DIGITALDIAGNOST C50", dt.date(2026, 8, 12)), ("PHILIPS COMBIDIAGNOST R90", dt.date(2026, 6, 3)), ("PHILIPS MOBILEDIAGNOST WDR", dt.date(2026, 7, 15)),
            ("PHILIPS IQON SPECTRAL CT", dt.date(2026, 9, 20)), ("FUJIFILM DIGITAL MAMMOGRAPHY FDR", dt.date(2026, 5, 28)), ("PHILIPS AZURION 7 M12", dt.date(2026, 8, 30)),
            ("PHILIPS BV ENDURA", dt.date(2026, 4, 14)), ("PHILIPS MRI INGENIA AMBITION X 1,5 T", dt.date(2026, 9, 5)), ("PHILIPS USG AFINITY 70 G", dt.date(2026, 6, 18))]
UK_ITEMS = [("PHILIPS DIGITALDIAGNOST C50", dt.date(2024, 11, 5), 48), ("PHILIPS COMBIDIAGNOST R90", dt.date(2022, 9, 20), 48), ("PHILIPS MOBILEDIAGNOST WDR", dt.date(2025, 2, 10), 48),
            ("PHILIPS IQON SPECTRAL CT", dt.date(2025, 9, 12), 48), ("ACTEON X-MIND PRIME 3D", dt.date(2024, 12, 1), 48), ("PHILIPS AZURION 7 M12", dt.date(2022, 12, 20), 48),
            ("PHILIPS BV ENDURA", dt.date(2024, 6, 30), 48), ("GE OEC ELITE", dt.date(2023, 11, 8), 48), ("GE LUNAR BMD", dt.date(2025, 3, 3), 48),
            ("ACTEON X-MIND DC", dt.date(2024, 10, 10), 48), ("FUJIFILM DIGITAL MAMMOGRAPHY FDR", dt.date(2024, 5, 20), 36)]
UK_DUMMY_NE = "GE OEC ELITE"
_KAL_NAMES = ["PHILIPS DIGITALDIAGNOST C50", "PHILIPS COMBIDIAGNOST R90", "PHILIPS MOBILEDIAGNOST WDR", "PHILIPS IQON SPECTRAL CT", "FUJIFILM DIGITAL MAMMOGRAPHY FDR",
              "ACTEON X-MIND PRIME 3D", "PHILIPS AZURION 7 M12", "PHILIPS BV ENDURA", "GE OEC ELITE", "GE LUNAR BMD", "ACTEON X-MIND DC", "PHILIPS MRI INGENIA AMBITION X 1,5 T",
              "PHILIPS USG AFINITY 70 G", "PENDOSE FUJI ELECTRIC / DOSE-I", "PENDOSE ALLOKA / PDM 222C-SH", "SURVEIMETER SE INTL RANGER", "INJEKTOR SYRINGE CT SCAN",
              "INJEKTOR SYRINGE MRI", "EWS MONITOR", "TIMBANGAN DIGITAL"]
_KAL_OFF = [40, 90, 150, 210, 260, 300, 330, 350, 380, 200, 120, 60, 30, 280, 340, 372, 100, 170, 240, 310]
KAL_ITEMS = list(zip(_KAL_NAMES, _KAL_OFF))
KAL_DUMMY_NC = "PENDOSE FUJI ELECTRIC / DOSE-I"


# =====================================================================
# 3. UTIL: tanggal, status, penyimpanan
# =====================================================================
def TODAY():
    v = st.session_state.get("sim_today")
    return v if isinstance(v, dt.date) else dt.date(2026, 10, 5)


def DUMMY():
    return bool(st.session_state.get("use_dummy", True))


def petugas():
    return str(st.session_state.get("user_name", "Berliana")).strip() or "Petugas"


def initials(name):
    p = [x for x in str(name).split() if x]
    return "".join(x[0].upper() for x in p[:2]) or "--"


def rng(*parts):
    return random.Random(zlib.crc32("|".join(str(x) for x in parts).encode()))


def dfmt(x):
    return "{:02d} {} {}".format(x.day, BLN3[x.month - 1], x.year)


def addm(x, m):
    y = x.year + (x.month - 1 + m) // 12
    mo = (x.month - 1 + m) % 12 + 1
    return dt.date(y, mo, min(x.day, 28))


def stt(left, soon=60):
    return "Overdue" if left < 0 else ("Due Soon" if left <= soon else "Valid")


ECOLS = ["id", "dibuat", "kategori", "objek", "tanggal", "shift", "nilai1", "nilai2", "status", "petugas", "temuan", "tindakan", "lampiran", "detail"]


def load_entries():
    if os.path.exists(ENTRY_FILE):
        try:
            df = pd.read_csv(ENTRY_FILE, dtype=str).fillna("")
        except Exception:
            df = pd.DataFrame(columns=ECOLS)
    else:
        df = pd.DataFrame(columns=ECOLS)
    for c in ECOLS:
        if c not in df.columns:
            df[c] = ""
    return df[ECOLS]


def save_entry(**kw):
    df = load_entries()
    row = {c: "" for c in ECOLS}
    for k, v in kw.items():
        row[k] = json.dumps(v) if k == "detail" else str(v)
    row["id"] = str(len(df) + 1)
    row["dibuat"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    df.to_csv(ENTRY_FILE, index=False)


def save_upload(f):
    """Simpan file unggahan, kembalikan path (atau '' bila tidak ada)."""
    if f is None:
        return ""
    name = "{}_{}".format(dt.datetime.now().strftime("%Y%m%d%H%M%S"), str(f.name).replace(" ", "_"))
    path = os.path.join(UPLOAD_DIR, name)
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(f.getvalue())
    return path


def recs(kategori, objek=None):
    df = load_entries()
    df = df[df["kategori"] == kategori]
    if objek is not None:
        df = df[df["objek"] == objek]
    out = []
    for r in df.to_dict("records"):
        try:
            r["tgl"] = dt.date.fromisoformat(r["tanggal"])
        except Exception:
            continue
        try:
            r["det"] = json.loads(r["detail"]) if r["detail"] else {}
        except Exception:
            r["det"] = {}
        out.append(r)
    out.sort(key=lambda x: (x["tgl"], x["dibuat"]))
    return out


def real_daily_map():
    """(nama, tanggal) -> 'g'/'r' dari input Daily QC & Suhu yang nyata."""
    m = {}
    for r in recs("Daily QC"):
        m[(r["objek"], r["tgl"])] = "g" if r["status"] == "Completed" else "r"
    tmp = {}
    for r in recs("Suhu & Kelembapan"):
        tmp.setdefault((r["objek"], r["tgl"]), []).append(r["status"])
    for k, v in tmp.items():
        m[k] = "r" if "Abnormal" in v else "g"
    return m


# =====================================================================
# 4. BUILDER HTML (tampilan = mockup PNG)
# =====================================================================
def pill(s):
    return "<span class='pill {}'>{}</span>".format(CL.get(s, "b"), s)


def kpi(items):
    o = "<div class='kpis'>"
    for t, v, s, seg in items:
        bars = "".join("<i class='{}' style='width:{}%'></i>".format(c, w) for c, w in seg)
        o += "<div class='kpi'><div class='t'>{}</div><div class='v'>{}</div><div class='s'>{}</div><div class='bar'>{}</div></div>".format(t, v, s, bars)
    return o + "</div>"


def sec(title, tabs, inner, note=""):
    tb = "".join("<span class='{}'>{}</span>".format("a" if i == 0 else "", x) for i, x in enumerate(tabs))
    return ("<div class='sec'><div class='sh'>{}<span>&#8963;</span></div><div class='tabs'>{}</div><div class='tw'>{}</div><div class='note'>{}</div></div>"
            .format(title, tb, inner, note))


def tbl(rows, heads):
    t = "<table><tr>" + "".join("<th class='{}'>{}</th>".format("l" if j == 1 else "", x) for j, x in enumerate(heads)) + "</tr>"
    for r in rows:
        t += "<tr>" + "".join("<td class='{}'>{}</td>".format("l" if j == 1 else "", x) for j, x in enumerate(r)) + "</tr>"
    return t + "</table>"


def info_card():
    return ("<div class='card info'><div>Hospital<b>Mandaya Royal Hospital Puri</b></div><div>Department<b>RADIOLOGY DEPARTMENT</b></div>"
            "<div>Role<b>RADIOLOGY DEPARTMENT</b></div></div>")


def legend_html(pme=False):
    last = "Terjadwal / Info" if pme else "Belum jatuh tempo / Info"
    red = "Overdue / Expired" if pme else "Not Compliant / Overdue"
    return ("<div class='legend'><span class='dot g'></span>Mencapai Target / Sesuai &nbsp; <span class='dot y'></span>Tidak Mencapai / Warning &nbsp; "
            "<span class='dot r'></span>{} &nbsp; <span class='dot b'></span>{}</div>".format(red, last))


BAN = "<div class='ban'>&#9888; SIMULASI TAMPILAN &ndash; data dummy ditampilkan bersama data yang diinput (matikan di sidebar bila tidak diperlukan).</div>"


# ---------- PMI: ringkasan harian ----------
def build_daily(y, m, today, dummy, rmap):
    n = calendar.monthrange(y, m)[1]
    r = rng("daily", y, m)
    h = ("<table style='table-layout:fixed'><tr><th style='width:30px'>No</th><th class='l' style='width:215px'>Jenis Modality / Area</th>"
         "<th style='width:105px'>Item Uji</th>" + "".join("<th class='h'>{:02d}</th>".format(i) for i in range(1, n + 1))
         + "<th style='width:52px'>Achv %</th></tr>")
    ratios = []
    for i, (name, item) in enumerate(DAILY):
        cells, ok, filled = "", 0, 0
        for dd in range(1, n + 1):
            dm = dt.date(y, m, dd)
            dv = r.choices(["g", "y", "r"], [90, 6, 4])[0]
            if dm > today:
                s = None
            elif (name, dm) in rmap:
                s = rmap[(name, dm)]
            elif dummy:
                s = dv
            else:
                s = None
            if s is None:
                cells += "<td></td>"
            else:
                filled += 1
                ok += 1 if s == "g" else 0
                cells += "<td class='c {}'>{}</td>".format(s, SYM[s])
        if filled:
            ratios.append(ok / filled)
        pct = "-" if not filled else "{}%".format(round(ok / filled * 100))
        h += "<tr><td>{}</td><td class='l'>{}</td><td>{}</td>{}<td><b>{}</b></td></tr>".format(i + 1, name, item, cells, pct)
    return h + "</table>", (sum(ratios) / len(ratios) if ratios else None)


# ---------- PMI: Suhu & Kelembapan ----------
def suhu_series(area, y, m, today, dummy):
    n = calendar.monthrange(y, m)[1]
    T, H, W = [None] * (n * 3), [None] * (n * 3), [""] * (n * 3)
    r = rng("suhu", area, y, m)
    real = {}
    for x in recs("Suhu & Kelembapan", area):
        if x["tgl"].year == y and x["tgl"].month == m:
            real[(x["tgl"].day, (x["shift"] or "P")[:1])] = x
    for d in range(1, n + 1):
        for k, code in enumerate("PSM"):
            idx = (d - 1) * 3 + k
            dv_t, dv_h = r.gauss(21, .9), r.gauss(50, 3.5)
            if r.random() < .06:
                dv_t = r.choice([r.uniform(16, 17.8), r.uniform(23.5, 26)])
            if r.random() < .05:
                dv_h = r.choice([r.uniform(33, 39), r.uniform(61, 65)])
            if (d, code) in real:
                x = real[(d, code)]
                try:
                    T[idx], H[idx] = float(x["nilai1"]), float(x["nilai2"])
                except ValueError:
                    pass
                W[idx] = initials(x["petugas"])
            elif dummy and dt.date(y, m, d) <= today:
                T[idx], H[idx], W[idx] = round(dv_t, 1), round(dv_h, 1), STAFF[(idx * 5 + idx // 7) % 9]
    return T, H, W


def chart_svg(title, vals, who, ymin, ymax, step, lo, hi, n, W=1518):
    L = 34
    cw = (W - L) / (n * 3)
    Hh, top, bot = 210, 34, 14
    y = lambda v: top + (ymax - v) / (ymax - ymin) * Hh
    s = "<svg width='{}' height='{}' xmlns='http://www.w3.org/2000/svg' font-family='DejaVu Sans' font-size='9'>".format(W, top + Hh + bot + 16)
    s += "<rect x='0' y='0' width='{}' height='16' fill='#e5e9ef'/><text x='{}' y='12' text-anchor='middle' font-weight='bold'>{}</text>".format(W, W / 2, title)
    for d in range(n):
        x = L + d * 3 * cw
        s += "<rect x='{}' y='16' width='{}' height='{}' fill='{}'/><text x='{}' y='26' text-anchor='middle' font-weight='bold'>{}</text>".format(
            x, 3 * cw, top - 16 + Hh, "#dff6f6" if d % 2 else "#fff", x + 1.5 * cw, d + 1)
        for k, nm in enumerate("PSM"):
            s += "<text x='{}' y='34' text-anchor='middle' font-size='7.5'>{}</text>".format(x + (k + .5) * cw, nm)
    s += "<rect x='{}' y='{}' width='{}' height='{}' fill='#d64545' opacity='.10'/>".format(L, y(ymax), W - L, y(hi) - y(ymax))
    s += "<rect x='{}' y='{}' width='{}' height='{}' fill='#d64545' opacity='.10'/>".format(L, y(lo), W - L, y(ymin) - y(lo))
    s += "<rect x='{}' y='{}' width='{}' height='{}' fill='#2e9e5b' opacity='.07'/>".format(L, y(hi), W - L, y(lo) - y(hi))
    v = ymin
    while v <= ymax:
        bad = v < lo or v > hi
        s += ("<line x1='{}' x2='{}' y1='{}' y2='{}' stroke='#dfe5ec'/><rect x='0' y='{}' width='30' height='12' fill='{}'/>"
              "<text x='15' y='{}' text-anchor='middle' fill='{}' font-weight='bold'>{:g}</text>").format(
            L, W, y(v), y(v), y(v) - 6, "#f7c9c9" if bad else "#fff", y(v) + 3, "#c0392b" if bad else "#222", v)
        v += step
    for t in (lo, hi):
        s += "<line x1='{}' x2='{}' y1='{}' y2='{}' stroke='#2e9e5b' stroke-dasharray='5 3' stroke-width='1.4'/>".format(L, W, y(t), y(t))
    pts = [(L + (i + .5) * cw, y(max(ymin, min(ymax, val)))) if val is not None else None for i, val in enumerate(vals)]
    seg = []
    for p in pts + [None]:
        if p is None:
            if len(seg) > 1:
                s += "<polyline fill='none' stroke='#2a7ca8' stroke-width='1' opacity='.55' points='{}'/>".format(" ".join("{:.1f},{:.1f}".format(a, b) for a, b in seg))
            seg = []
        else:
            seg.append(p)
    for p, val in zip(pts, vals):
        if p is not None:
            s += "<circle cx='{:.1f}' cy='{:.1f}' r='3' fill='{}'/>".format(p[0], p[1], "#d64545" if (val < lo or val > hi) else "#222")
    for i in range(n * 3):
        s += "<text x='{:.1f}' y='{}' text-anchor='middle' font-size='6.5' fill='#335'>{}</text>".format(L + (i + .5) * cw, top + Hh + 11, who[i])
    return s + "<text x='15' y='{}' text-anchor='middle' font-size='6.5' font-weight='bold'>PIC</text></svg>".format(top + Hh + 11)


def suhu_section_html(area, y, m, today, dummy):
    n = calendar.monthrange(y, m)[1]
    T, H, W = suhu_series(area, y, m, today, dummy)
    tv, hv = [v for v in T if v is not None], [v for v in H if v is not None]
    inT = sum(T_LO <= v <= T_HI for v in tv)
    inH = sum(H_LO <= v <= H_HI for v in hv)

    def st_(l, v, c="#111"):
        return "<div class='stat'><div class='l'>{}</div><div class='v' style='color:{}'>{}</div></div>".format(l, c, v)

    avg_t = "{:.1f} &deg;C <small style='font-size:10px;color:#666'>(target 18&ndash;23)</small>".format(sum(tv) / len(tv)) if tv else "-"
    avg_h = "{:.0f} % <small style='font-size:10px;color:#666'>(target 40&ndash;60)</small>".format(sum(hv) / len(hv)) if hv else "-"
    pt = "{}/{} &middot; {}%".format(inT, len(tv), round(inT / len(tv) * 100)) if tv else "-"
    ph = "{}/{} &middot; {}%".format(inH, len(hv), round(inH / len(hv) * 100)) if hv else "-"
    alert = (len(tv) - inT) + (len(hv) - inH)
    stats = ("<div class='stats'>" + st_("Ruang", area) + st_("Rata-rata Suhu", avg_t) + st_("Rata-rata Kelembapan", avg_h)
             + st_("Suhu dalam rentang", pt, "#2e9e5b") + st_("Kelembapan dalam rentang", ph, "#2e9e5b")
             + st_("Alert di luar batas", "{} &#9888;".format(alert), "#d64545") + "</div>")
    body = stats + chart_svg("SUHU (&deg;C) &ndash; Target Temperatur 18 &ndash; 23 &deg;C &middot; Shift P = Pagi, S = Siang, M = Malam", T, W, 15, 30, 1, T_LO, T_HI, n)
    body += "<div style='height:8px'></div>" + chart_svg("KELEMBAPAN (%) &ndash; Target 40 &ndash; 60 %", H, W, 30, 65, 5, H_LO, H_HI, n)
    return sec("PMI Harian &ndash; Form Digital Monitoring Suhu dan Kelembapan", ["Chart", "Table", "Alert", "Summary"], body,
               "Titik merah = nilai di luar target &middot; PIC = inisial petugas yang menginput pada tiap shift &middot; di luar batas wajib mengisi temuan & tindakan korektif (notifikasi ke Supervisor Radiologi)")


# ---------- PMI: Daily QC checklist ----------
def flat_params(modality):
    cats = CHECKLISTS.get(modality, DEFAULT_CHECKLIST)
    return cats, sum(len(items) for _, items in cats)


def qc_section_html(modality, y, m, today, dummy):
    n = calendar.monthrange(y, m)[1]
    cats, total = flat_params(modality)
    real = {}
    for x in recs("Daily QC", modality):
        if x["tgl"].year == y and x["tgl"].month == m:
            real[x["tgl"].day] = x
    r = rng("qc", modality, y, m)
    q = ("<table style='table-layout:fixed;font-size:9.5px'><tr><th style='width:26px'>NO</th><th class='l' style='width:215px'>KEGIATAN</th>"
         "<th class='l' style='width:235px'>PARAMETER</th>" + "".join("<th class='h'>{}</th>".format(i) for i in range(1, n + 1)) + "<th style='width:44px'>%</th></tr>")
    flat = 0
    for cat, items in cats:
        q += "<tr><td colspan='{}' style='background:#cfe3f7;font-weight:bold'>{}</td></tr>".format(n + 4, cat)
        for i, (k, p) in enumerate(items):
            q += "<tr><td>{}</td><td class='l'>{}</td><td class='l'>{}</td>".format(i + 1, k, p)
            ok = filled = 0
            for dd in range(1, n + 1):
                dm = dt.date(y, m, dd)
                dv = r.choices(["g", "y", "r"], [96, 2.5, 1.5])[0]
                if dm > today:
                    s = None
                elif dd in real:
                    stl = real[dd]["det"].get("st", [])
                    s = stl[flat] if flat < len(stl) else "g"
                elif dummy:
                    s = dv
                else:
                    s = None
                if s is None:
                    q += "<td></td>"
                else:
                    filled += 1
                    ok += 1 if s == "g" else 0
                    if s == "g":
                        q += "<td style='background:#e3f4ea;color:#2e9e5b;font-weight:bold'>&#10003;</td>"
                    else:
                        q += "<td class='c {}'>{}</td>".format(s, "!" if s == "y" else "&#10007;")
            q += "<td><b>{}</b></td></tr>".format("-" if not filled else "{}%".format(round(ok / filled * 100)))
            flat += 1
    q += "<tr><td colspan='3' style='background:#dcdcdc;font-weight:bold'>NAMA PEKERJA RADIASI</td>"
    for dd in range(1, n + 1):
        nm = initials(real[dd]["petugas"]) if dd in real else (STAFF[((dd - 1) * 4 + (dd - 1) // 5) % 9] if (dummy and dt.date(y, m, dd) <= today) else "")
        q += "<td style='background:#dcdcdc;font-weight:bold;font-size:8.5px'>{}</td>".format(nm)
    q += "<td style='background:#dcdcdc'></td></tr></table>"
    head = ("<div style='font-size:11px;margin:0 0 8px'><b>Modalitas :</b> {} &nbsp;&nbsp; <b>BULAN :</b> {} {} &nbsp;&nbsp; "
            "<i style='color:#777'>Form Digital Daily Quality Control Departemen Radiologi Tahun {}</i></div>").format(modality, BLN[m - 1].upper(), y, y)
    return sec("PMI Harian &ndash; Form Digital Daily Quality Control Alat Radiologi", ["Checklist", "Trend %", "Not Compliant", "Summary"], head + q,
               "Checklist mengikuti jenis alat &middot; sel merah/kuning = parameter tidak sesuai, wajib isi temuan & tindakan awal")


# ---------- PMI: mingguan, bulanan, tahunan ----------
def month_weeks(y, m):
    n = calendar.monthrange(y, m)[1]
    out, s = [], 1
    while s <= n:
        wd = dt.date(y, m, s).weekday()
        e = min(n, s + (6 - wd))
        out.append((dt.date(y, m, s), dt.date(y, m, e)))
        s = e + 1
    return out


def build_weekly(y, m, today, dummy):
    weeks = month_weeks(y, m)
    real = recs("QC Mingguan")
    r = rng("weekly", y, m)
    h = ("<table><tr><th class='l' style='width:200px'>Jenis Modality</th><th class='l'>Item Uji</th>"
         + "".join("<th class='h'>Minggu {}<br><small>{:02d}-{:02d} {}</small></th>".format(i + 1, a.day, b.day, BLN3[m - 1]) for i, (a, b) in enumerate(weeks))
         + "<th>Achv %</th></tr>")
    comp = started = 0
    for label, item in WEEKLY_ROWS:
        h += "<tr><td class='l'>{}</td><td class='l'>{}</td>".format(label, item)
        done = start = 0
        for ws, we in weeks:
            dv = r.choices(["Completed", "Failed", "Overdue"], [84, 8, 8])[0]
            ex = [x for x in real if x["objek"] == label and ws <= x["tgl"] <= we]
            if ex:
                s = "Completed" if ex[-1]["status"] == "Completed" else "Failed"
            elif we < today:
                s = dv if dummy else "Overdue"
            elif ws <= today <= we:
                s = "Pending"
            else:
                s = "Terjadwal"
            if ws <= today:
                start += 1
                done += 1 if s == "Completed" else 0
            h += "<td>{}</td>".format(pill(s))
        comp += done
        started += start
        h += "<td><b>{}</b></td></tr>".format("-" if not start else "{}%".format(round(done / start * 100)))
    return h + "</table>", (comp, started)


def build_monthly(y, m, today, dummy):
    real = recs("QC Bulanan")
    r = rng("monthly", y, m)
    h = ("<table><tr><th class='l' style='width:190px'>Jenis Modality</th><th class='l'>Item Uji</th><th>Parameter Selesai</th>"
         "<th style='width:170px'>Progres</th><th>Status Bulan Ini</th><th>Evidence</th></tr>")
    sd = st_ = 0
    for label, item in BULANAN_ROWS:
        total = len(BULANAN_SPEC[label])
        dr = r.choices(["Pass", "Pending", "Failed", "Overdue"], [55, 15, 15, 15])[0]
        dd_ = r.randint(max(1, total // 2), max(1, total - 1))
        ex = [x for x in real if x["objek"] == label and x["tgl"].year == y and x["tgl"].month == m]
        past, cur = (y, m) < (today.year, today.month), (y, m) == (today.year, today.month)
        if ex:
            s, done = ex[-1]["status"], int(ex[-1]["det"].get("done", total))
        elif not (past or cur):
            s, done = "Terjadwal", 0
        elif dummy:
            s = "Pending" if (cur and dr == "Overdue") else dr
            done = total if s == "Pass" else dd_
        else:
            s, done = ("Overdue" if past else "Pending"), 0
        sd += done
        st_ += total
        p = round(done / total * 100)
        c = {"Pass": "g", "Pending": "b", "Failed": "r", "Overdue": "r", "Terjadwal": "b"}[s]
        h += ("<tr><td class='l'><b>{}</b></td><td class='l'>{}</td><td>{} / {}</td><td><div class='bar' style='height:9px'><i class='{}' style='width:{}%'></i></div>{}%</td>"
              "<td>{}</td><td>{}</td></tr>").format(label, item, done, total, c, p, p, pill(s), "&#128206; Lengkap" if (done == total and s == "Pass") else "&#9888; Belum lengkap")
    return h + "</table>", (sd, st_)


def yearly_model(today, dummy):
    rows = []
    real_all = recs("Uji Kebocoran")
    for i, (name, item, base) in enumerate(LEAK_ITEMS):
        ex = [x for x in real_all if x["objek"] == name]
        if ex:
            last, nc = ex[-1]["tgl"], ex[-1]["status"] == "Tidak Memenuhi"
        elif dummy:
            last, nc = base, False
        else:
            rows.append(dict(i=i + 1, name=name, item=item, last=None, nx=None, left=None, status="Belum Ada Data"))
            continue
        nx = addm(last, 12)
        left = (nx - today).days
        s = "Not Compliant" if nc else ("Overdue" if left < 0 else ("Due Soon" if left <= 60 else "Compliant"))
        rows.append(dict(i=i + 1, name=name, item=item, last=last, nx=nx, left=left, status=s))
    return rows


def build_yearly(today, dummy):
    rows = yearly_model(today, dummy)
    cnt = {"Compliant": 0, "Due Soon": 0, "Overdue": 0}
    out = []
    for r in rows:
        if r["status"] in ("Overdue", "Not Compliant"):
            cnt["Overdue"] += 1
        elif r["status"] in cnt:
            cnt[r["status"]] += 1
        out.append([r["i"], r["name"], r["item"], dfmt(r["last"]) if r["last"] else "-", dfmt(r["nx"]) if r["nx"] else "-",
                    "<b>{}</b>".format(r["left"]) if r["left"] is not None else "-", pill(r["status"])])
    return tbl(out, ["No", "Jenis Modality", "Item Uji", "Terakhir Dilakukan", "Jatuh Tempo Berikutnya", "Sisa Hari", "Status"]), cnt, rows


def build_pmi(y, m):
    today, dummy = TODAY(), DUMMY()
    rmap = real_daily_map()
    daily_html, davg = build_daily(y, m, today, dummy, rmap)
    wk_html, (wc, ws) = build_weekly(y, m, today, dummy)
    mo_html, (md, mt) = build_monthly(y, m, today, dummy)
    yr_html, ycnt, yrows = build_yearly(today, dummy)
    dpct = round(davg * 100) if davg is not None else 0
    wpct = round(wc / ws * 100) if ws else 0
    mpct = round(md / mt * 100) if mt else 0
    nY = max(1, len(LEAK_ITEMS))
    kp = kpi([("PMI Harian ({} {})".format(BLN3[m - 1], y), "{}%".format(dpct) if davg is not None else "-", "Target 100% &middot; 11 item", [("g", dpct), ("y", (100 - dpct) * 6 // 10), ("r", (100 - dpct) * 4 // 10)]),
              ("PMI Mingguan", "{}%".format(wpct) if ws else "-", "{} dari {} kegiatan Completed".format(wc, ws), [("g", wpct), ("r", 100 - wpct)]),
              ("PMI Bulanan ({})".format(BLN3[m - 1]), "{}%".format(mpct), "{} dari {} parameter selesai".format(md, mt), [("g", mpct), ("r", 100 - mpct)]),
              ("PMI Tahunan", "{} item".format(len(LEAK_ITEMS)), "Compliant {} &middot; Due Soon {} &middot; Overdue {}".format(ycnt["Compliant"], ycnt["Due Soon"], ycnt["Overdue"]),
               [("g", ycnt["Compliant"] * 100 // nY), ("y", ycnt["Due Soon"] * 100 // nY), ("r", ycnt["Overdue"] * 100 // nY)])])
    TB = ["Table", "Trend %", "Top 5", "KPI %", "Summary"]
    return dict(
        top=kp + sec("PMI Harian &ndash; Ringkasan Achievement Suhu & QC Harian (11 Modality / Area)", TB, daily_html, "Sel hijau = sesuai &middot; kuning = perlu perhatian &middot; merah = tidak sesuai / terlambat &middot; kosong = belum ada data / belum jatuh tempo"),
        rest=(sec("PMI Mingguan &ndash; QC Mingguan (Fisikawan Medis)", TB, wk_html)
              + sec("PMI Bulanan &ndash; QC Bulanan (Fisikawan Medis)", TB, mo_html)
              + sec("PMI Tahunan &ndash; Uji Kebocoran Tabung X-Ray & Radiasi (PPR / Fisikawan Medis)", TB, yr_html)),
        yearly=yrows)


# ---------- PME ----------
def build_pme(today, dummy):
    MON = BLN3
    prio = []
    # --- Preventive Maintenance ---
    pm_rows, cp = [], {"Valid": 0, "Due Soon": 0, "Overdue": 0}
    pm_real = recs("Preventive Maintenance")
    for i, (name, base) in enumerate(PM_ITEMS):
        ex = [x for x in pm_real if x["objek"] == name]
        follow = False
        if ex:
            last = ex[-1]["tgl"]
            done = sorted(x["tgl"] for x in ex if x["tgl"].year == today.year)
            nx = addm(last, 4)
            try:
                nx = dt.date.fromisoformat(ex[-1]["det"].get("next", ""))
            except Exception:
                pass
            follow = ex[-1]["status"] == "Need Follow-up"
            rep = os.path.basename(ex[-1]["lampiran"]) if ex[-1]["lampiran"] else "-"
        elif dummy:
            last, rep = base, "SR-{}-{:02d}".format(base.strftime("%y%m"), i + 1)
            done = [x for x in (addm(base, -8), addm(base, -4), base) if x.year == today.year]
            nx = addm(base, 4)
        else:
            continue
        left = (nx - today).days
        s = "Need Follow-up" if follow else stt(left)
        cp["Valid" if s == "Valid" else ("Overdue" if s == "Overdue" else "Due Soon")] += 1
        pm_rows.append(dict(i=i, name=name, last=last, nx=nx, left=left, status=s, done=done, rep=rep))
        if s != "Valid":
            prio.append(("Preventive Maintenance", name, s, nx, left, "Elektromedis / Teknisi Vendor",
                         "Konfirmasi jadwal PM dengan vendor & unggah service report" if s in ("Due Soon", "Need Follow-up") else "Jadwalkan PM segera; catat risiko unplanned downtime"))
    t = ("<table style='table-layout:fixed'><tr><th style='width:28px'>No</th><th class='l' style='width:235px'>Jenis Modality</th>"
         + "".join("<th class='h' style='width:40px;{}'>{}</th>".format("outline:2px solid #235b9c;outline-offset:-2px" if k + 1 == today.month else "", mn) for k, mn in enumerate(MON))
         + "<th style='width:78px'>Realisasi</th><th>Terakhir PM</th><th>Berikutnya</th><th style='width:62px'>Sisa Hari</th><th>Status</th><th>Service Report</th></tr>")
    for k, r in enumerate(pm_rows):
        t += "<tr><td>{}</td><td class='l'>{}</td>".format(k + 1, r["name"])
        for mo in range(1, 13):
            ev = [x for x in r["done"] if x.month == mo]
            if ev:
                t += "<td class='c g'>{}</td>".format(ev[0].day)
            elif r["nx"].year == today.year and r["nx"].month == mo:
                t += "<td class='c {}'>{}</td>".format({"Valid": "b", "Due Soon": "y", "Need Follow-up": "y", "Overdue": "r"}[r["status"]], r["nx"].day)
            else:
                t += "<td style='background:{}'></td>".format("#fff6df" if mo == today.month else "#fff")
        t += "<td>{} / 3</td><td>{}</td><td>{}</td><td><b>{}</b></td><td>{}</td><td>&#128206; {}</td></tr>".format(len(r["done"]), dfmt(r["last"]), dfmt(r["nx"]), r["left"], pill(r["status"]), r["rep"])
    t += "</table>"
    pm_html = t if pm_rows else "<div style='padding:14px;color:#777'>Belum ada data PM.</div>"

    # --- Uji Kesesuaian ---
    uk_rows, cu = [], {"Aktif": 0, "Warning": 0, "Expired": 0, "Need Evaluation": 0}
    uk_real = recs("Uji Kesesuaian")
    for i, (name, base, per) in enumerate(UK_ITEMS):
        ex = [x for x in uk_real if x["objek"] == name]
        if ex:
            last, ne, noser = ex[-1]["tgl"], ex[-1]["status"] == "Need Evaluation", ex[-1]["det"].get("noser", "-")
            nx = addm(last, per)
            try:
                nx = dt.date.fromisoformat(ex[-1]["det"].get("next", ""))
            except Exception:
                pass
        elif dummy:
            last, ne, noser, nx = base, name == UK_DUMMY_NE, "UK/RAD/{}/{:03d}".format(base.year, i + 1), addm(base, per)
        else:
            continue
        left = (nx - today).days
        s = "Need Evaluation" if ne else ("Expired" if left < 0 else ("Warning" if left <= 180 else "Aktif"))
        cu[s] += 1
        uk_rows.append(dict(i=i, name=name, per=per, last=last, nx=nx, left=left, status=s, noser=noser, ne=ne))
        if s != "Aktif":
            prio.append(("Uji Kesesuaian", name, s, nx, left, "PPR / Fisikawan Medis",
                         "Perpanjang uji & lengkapi dokumen audit" if s == "Expired" else ("Evaluasi alat oleh Fisikawan Medis sebelum digunakan" if s == "Need Evaluation" else "Jadwalkan uji ulang; siapkan laporan/sertifikat")))
    uk_html = "<div style='padding:14px;color:#777'>Belum ada data Uji Kesesuaian.</div>"
    if uk_rows:
        starts = [x["last"] for x in uk_rows]
        ends = [x["nx"] for x in uk_rows]
        d0, d1 = dt.date(min(starts).year, 1, 1), dt.date(max(ends).year, 12, 31)
        LW, W, RH = 260, 1518, 24
        X = lambda x: LW + (x - d0).days / (d1 - d0).days * (W - LW - 60)
        g = "<svg width='{}' height='{}' xmlns='http://www.w3.org/2000/svg' font-family='DejaVu Sans' font-size='10'>".format(W, RH * len(uk_rows) + 40)
        for yv in range(d0.year, d1.year + 1):
            xx = X(dt.date(yv, 1, 1))
            g += "<line x1='{0}' x2='{0}' y1='24' y2='{1}' stroke='#dfe5ec'/><text x='{2}' y='15' font-weight='bold' fill='#456'>{3}</text>".format(xx, RH * len(uk_rows) + 30, xx + 4, yv)
        trows = []
        for k, r in enumerate(uk_rows):
            yy = 30 + k * RH
            g += ("<text x='6' y='{}'>{}</text><rect x='{}' y='{}' width='{}' height='14' rx='3' fill='{}' opacity='.85'/>"
                  "<text x='{}' y='{}' fill='#333' font-size='9'>{}</text>").format(yy + 14, r["name"], X(r["last"]), yy + 4, max(2, X(r["nx"]) - X(r["last"])), HX[CL[r["status"]]], X(r["nx"]) + 5, yy + 15, dfmt(r["nx"]))
            rem = "Eskalasi ke Supervisor" if r["left"] < 0 else ("Reminder H-180 terkirim" if r["left"] <= 180 else "H-180: " + dfmt(r["nx"] - dt.timedelta(days=180)))
            trows.append([k + 1, r["name"], "{} Tahun".format(r["per"] // 12), dfmt(r["last"]), dfmt(r["nx"]), r["left"], r["noser"], "Perlu Evaluasi" if r["ne"] else "Memenuhi", rem, pill(r["status"])])
        tx = X(today)
        g += "<line x1='{0}' x2='{0}' y1='20' y2='{1}' stroke='#d64545' stroke-width='1.5' stroke-dasharray='4 3'/><text x='{2}' y='{3}' fill='#d64545' font-weight='bold'>Hari ini {4}</text></svg>".format(tx, RH * len(uk_rows) + 30, tx + 4, RH * len(uk_rows) + 38, dfmt(today))
        uk_html = g + tbl(trows, ["No", "Jenis Modality", "Periode", "Tanggal Uji", "Masa Berlaku s.d.", "Sisa Hari", "No. Sertifikat", "Hasil Uji", "Reminder Otomatis", "Status"])

    # --- Kalibrasi ---
    kal_rows, ck = [], {"Valid": 0, "Due Soon": 0, "Overdue": 0, "Not Compliant": 0}
    kal_real = recs("Kalibrasi")
    for i, (name, off) in enumerate(KAL_ITEMS):
        ex = [x for x in kal_real if x["objek"] == name]
        if ex:
            last, nc, noser = ex[-1]["tgl"], ex[-1]["status"] == "Not Compliant", ex[-1]["det"].get("noser", "-")
            nx = last + dt.timedelta(days=365)
            try:
                nx = dt.date.fromisoformat(ex[-1]["det"].get("next", ""))
            except Exception:
                pass
        elif dummy:
            last = today - dt.timedelta(days=off)
            nx, nc, noser = last + dt.timedelta(days=365), name == KAL_DUMMY_NC, "KAL/2026/{:03d}".format(i + 1)
        else:
            continue
        left = (nx - today).days
        s = "Not Compliant" if nc else stt(left)
        ck[s] += 1
        kal_rows.append(dict(i=i, name=name, last=last, nx=nx, left=left, status=s, noser=noser))
        if s != "Valid":
            prio.append(("Kalibrasi", name, s, nx, left, "Elektromedis / PIC Alat",
                         "Evaluasi alat sebelum digunakan kembali; kalibrasi ulang" if s == "Not Compliant" else ("Kalibrasi segera & unggah sertifikat" if s == "Overdue" else "Jadwalkan kalibrasi dengan vendor")))
    mlist, yy_, mm_ = [], today.year, today.month
    for _ in range(12):
        mlist.append((yy_, mm_))
        mm_ += 1
        if mm_ > 12:
            mm_, yy_ = 1, yy_ + 1
    t = ("<table style='table-layout:fixed'><tr><th style='width:28px'>No</th><th class='l' style='width:235px'>Jenis Modality / Alat</th><th style='width:48px;background:#f3d1d1'>Lewat</th>"
         + "".join("<th class='h' style='width:44px;{}'>{} {}</th>".format("outline:2px solid #235b9c;outline-offset:-2px" if k == 0 else "", MON[mo - 1], str(yv)[2:]) for k, (yv, mo) in enumerate(mlist))
         + "<th style='width:95px'>Terakhir</th><th style='width:95px'>Berlaku s.d.</th><th style='width:55px'>Sisa</th><th>No. Sertifikat</th><th>Status</th></tr>")
    for k, r in enumerate(kal_rows):
        t += "<tr><td>{}</td><td class='l'>{}</td>".format(k + 1, r["name"])
        t += "<td class='c r'>{}</td>".format(r["nx"].day) if r["left"] < 0 else "<td></td>"
        for (yv, mo) in mlist:
            t += "<td class='c {}'>{}</td>".format(CL[r["status"]], r["nx"].day) if ((r["nx"].year, r["nx"].month) == (yv, mo) and r["left"] >= 0) else "<td></td>"
        t += "<td>{}</td><td>{}</td><td><b>{}</b></td><td>{}</td><td>{}</td></tr>".format(dfmt(r["last"]), dfmt(r["nx"]), r["left"], r["noser"], pill(r["status"]))
    kal_html = (t + "</table>") if kal_rows else "<div style='padding:14px;color:#777'>Belum ada data Kalibrasi.</div>"

    prio.sort(key=lambda x: x[4])
    pr = [[i + 1, c, nm, pill(s), dfmt(nx), "<b>{}</b>".format(lf), p, a, "Supervisor Radiologi + PIC" if lf < 0 else "PIC Alat"] for i, (c, nm, s, nx, lf, p, a) in enumerate(prio[:12])]
    pr_html = tbl(pr, ["No", "Kategori", "Jenis Modality / Alat", "Status", "Jatuh Tempo", "Sisa Hari", "PIC", "Tindakan / Rekomendasi", "Notifikasi Dikirim ke"]) if pr else "<div style='padding:14px;color:#2e9e5b'>&#10003; Tidak ada item yang perlu tindak lanjut.</div>"

    def seg(c, n, good, warn):
        n = max(1, n)
        gv = sum(c.get(k, 0) for k in good) * 100 // n
        wv = sum(c.get(k, 0) for k in warn) * 100 // n
        return [("g", gv), ("y", wv), ("r", max(0, 100 - gv - wv))]

    kp = kpi([
        ("Preventive Maintenance", "{}/{}".format(cp["Valid"], len(pm_rows)), "Sesuai jadwal &middot; Due Soon {} &middot; Overdue {}".format(cp["Due Soon"], cp["Overdue"]), seg(cp, len(pm_rows), ["Valid"], ["Due Soon"])),
        ("Uji Kesesuaian MRHP {}".format(today.year), "{}/{}".format(cu["Aktif"], len(uk_rows)), "Aktif &middot; Warning {} &middot; Expired {} &middot; Eval {}".format(cu["Warning"], cu["Expired"], cu["Need Evaluation"]), seg(cu, len(uk_rows), ["Aktif"], ["Warning"])),
        ("Kalibrasi MRHP {}".format(today.year), "{}/{}".format(ck["Valid"], len(kal_rows)), "Valid &middot; Due Soon {} &middot; Overdue {} &middot; NC {}".format(ck["Due Soon"], ck["Overdue"], ck["Not Compliant"]), seg(ck, len(kal_rows), ["Valid"], ["Due Soon"])),
        ("Perlu Tindak Lanjut", "{} item".format(len(prio)), "Reminder bertingkat aktif ke PIC &amp; Supervisor", [("r", 100)])])
    TB = ["Table", "Trend %", "Top 5", "KPI %", "Summary"]
    return dict(
        kpi=kp,
        prio=sec("PME &ndash; Prioritas Tindak Lanjut & Reminder Otomatis (urut jatuh tempo terdekat)", ["Prioritas", "Reminder", "Audit Trail", "Summary"], pr_html, "Notifikasi bertingkat otomatis: H-180 &rarr; H-90 &rarr; H-30 &rarr; Overdue (eskalasi ke Supervisor Radiologi)"),
        pm=sec("PME &ndash; Preventive Maintenance {} (Teknisi Vendor / Elektromedis / PIC Alat Radiologi)".format(today.year), ["Kalender", "Riwayat", "Due Soon", "Summary"], pm_html, "Angka dalam kalender = tanggal pelaksanaan &middot; hijau = sudah dilakukan &middot; biru = terjadwal &middot; kuning = due soon &middot; merah = overdue &middot; periode 2&ndash;4 kali setahun"),
        uk=sec("PME &ndash; Uji Kesesuaian MRHP {} &middot; Timeline Masa Berlaku (PPR / Fisikawan Medis / PIC Alat)".format(today.year), ["Timeline", "Table", "Reminder", "Summary"], uk_html, "Periode uji: 4 tahun (3 tahun untuk Fujifilm Digital Mammography FDR) &middot; garis merah putus-putus = hari ini"),
        kal=sec("PME &ndash; Kalibrasi MRHP {} &middot; Jadwal Jatuh Tempo 12 Bulan ke Depan (Elektromedis / PIC Alat)".format(today.year), ["Jadwal", "Sertifikat", "Overdue", "Summary"], kal_html, "Kolom bulan = bulan masa berlaku sertifikat berakhir &middot; kolom Lewat = sudah melewati jatuh tempo &middot; periode kalibrasi 1 tahun sekali"),
        pm_rows=pm_rows, uk_rows=uk_rows, kal_rows=kal_rows, prio_list=prio)


# ---------- Struktur menu ----------
def pending_counts(today, dummy):
    rs = recs("Suhu & Kelembapan")
    su = sum(1 for a in AREAS if not any(x["objek"] == a and x["tgl"] == today for x in rs))
    dq = recs("Daily QC")
    dqc = sum(1 for a in DAILY_MODALITIES if not any(x["objek"] == a and x["tgl"] == today for x in dq))
    ws, we = next(((a, b) for a, b in month_weeks(today.year, today.month) if a <= today <= b), (today, today))
    wk = recs("QC Mingguan")
    wkc = sum(1 for l, _ in WEEKLY_ROWS if not any(x["objek"] == l and ws <= x["tgl"] <= we for x in wk))
    bl = recs("QC Bulanan")
    blc = sum(1 for l, _ in BULANAN_ROWS if not any(x["objek"] == l and x["tgl"].year == today.year and x["tgl"].month == today.month for x in bl))
    ye = sum(1 for r in yearly_model(today, dummy) if r["status"] in ("Due Soon", "Overdue", "Not Compliant", "Belum Ada Data"))
    pme = build_pme(today, dummy)
    cnt = {"Preventive Maintenance": 0, "Uji Kesesuaian": 0, "Kalibrasi": 0}
    for c, *_ in pme["prio_list"]:
        cnt[c] += 1
    return dict(su=su, dq=dqc, wk=wkc, bl=blc, ye=ye, pm=cnt["Preventive Maintenance"], uk=cnt["Uji Kesesuaian"], kal=cnt["Kalibrasi"])


def menu_html():
    pc = pending_counts(TODAY(), DUMMY())

    def it(ic, c, t, a, fq, fc, n=0):
        b = "<span class='bdg'>{}</span>".format(n) if n else ""
        return ("<div class='it'><div class='ic' style='background:{}'>{}</div><div><div class='t'>{}{}</div><div class='a'>{}</div></div>"
                "<div class='fq'><span style='background:{}22;color:{}'>{}</span></div></div>").format(c, ic, t, b, a, fc, fc, fq)

    pmi = ("<div class='bc2'><div class='bh' style='background:linear-gradient(135deg,#1b4f86,#2f7fd0)'><h2>PMI &ndash; Pemantapan Mutu Internal</h2>"
           "<div>Kegiatan mutu rutin di dalam departemen &middot; dilakukan Radiografer, Fisikawan Medis, PPR</div></div>"
           + it("&#9728;", "#2f7fd0", "Suhu & Kelembapan", "Radiografer &middot; input per shift P / S / M", "Harian", "#2f7fd0", pc["su"])
           + it("&#10003;", "#2f7fd0", "Daily Quality Control", "Radiografer &middot; checklist per modality", "Harian", "#2f7fd0", pc["dq"])
           + it("&#9776;", "#7a4fb5", "QC Mingguan", "Fisikawan Medis &middot; Image Quality, Air Calibration, Homogeneity", "Mingguan", "#7a4fb5", pc["wk"])
           + it("&#9678;", "#7a4fb5", "QC Bulanan", "Fisikawan Medis &middot; parameter QC per modality", "Bulanan", "#7a4fb5", pc["bl"])
           + it("&#9762;", "#c0392b", "Uji Kebocoran Tabung X-Ray & Radiasi", "PPR / Fisikawan Medis &middot; laporan &amp; sertifikat", "Tahunan", "#c0392b", pc["ye"])
           + it("&#9636;", "#445566", "Dashboard PMI", "Achievement harian s.d. tahunan &middot; tren &amp; alert", "Monitoring", "#445566") + "</div>")
    pme = ("<div class='bc2'><div class='bh' style='background:linear-gradient(135deg,#0b6e70,#18a999)'><h2>PME &ndash; Pemantapan Mutu Eksternal</h2>"
           "<div>Kegiatan mutu dengan pihak luar / regulator &middot; Vendor, Elektromedis, PIC Alat, PPR</div></div>"
           + it("&#9874;", "#0f8b8d", "Preventive Maintenance", "Teknisi Vendor / Elektromedis &middot; service report", "2&ndash;4x / tahun", "#0f8b8d", pc["pm"])
           + it("&#9745;", "#0f8b8d", "Uji Kesesuaian MRHP 2026", "PPR / Fisikawan Medis / PIC Alat &middot; sertifikat", "3&ndash;4 tahun", "#0f8b8d", pc["uk"])
           + it("&#9878;", "#0f8b8d", "Kalibrasi MRHP 2026", "Elektromedis / PIC Alat &middot; 20 alat", "1 tahun", "#0f8b8d", pc["kal"])
           + it("&#9636;", "#445566", "Dashboard PME", "Timeline masa berlaku &middot; kalender PM &middot; prioritas", "Monitoring", "#445566")
           + "<div style='padding:22px 18px;font-size:11px;color:#778;background:#f8fafc'>Setiap input PME otomatis menghitung jadwal / masa berlaku berikutnya dan menjadwalkan reminder bertingkat.</div></div>")
    sm = "<div class='sh2'>" + "".join("<div class='sm'><b>{}</b><div>{}</div></div>".format(a, b) for a, b in [
        ("Dashboard DRACORE", "Manajemen memantau Quality, Compliance &amp; Operational Reliability dari satu layar"),
        ("Dokumen & Perizinan", "Repository sertifikat, izin alat, service report &middot; reminder masa berlaku"),
        ("Laporan & Audit Trail", "Ekspor laporan, telusur histori &amp; bukti pendukung untuk audit / akreditasi"),
        ("Notifikasi & Reminder", "H-180 / H-90 / H-30 / Overdue ke PIC &amp; Supervisor Radiologi"),
        ("Master Data & Hak Akses", "Modality, parameter &amp; batas QC, jadwal, role pengguna")]) + "</div>"
    fl = "<div class='fl'>" + "<div class='ar'>&rarr;</div>".join("<div class='st'><b>{}</b>{}</div>".format(a, b) for a, b in [
        ("1. Input", "Form digital PMI / PME"), ("2. Validasi", "Batas standar &amp; kelengkapan"), ("3. Dashboard", "Status real-time"),
        ("4. Reminder", "Notifikasi otomatis"), ("5. Laporan", "Audit &amp; histori")]) + "</div>"
    return "<div class='big'>" + pmi + pme + "</div>" + sm + fl


# =====================================================================
# 5. KOMPONEN UI NATIVE
# =====================================================================
def heading(title, crumb, role=None, color="#2a7ca8", sub=""):
    tag = "<span class='dc-tag' style='background:{}'>{}</span>".format(color, role) if role else ""
    st.markdown("<div class='dc-bc'>{}</div><div class='dc-h'>{}{}</div><div class='dc-sub'>{}</div>".format(crumb, title, tag, sub), unsafe_allow_html=True)


def hint(ok, text):
    st.markdown("<div class='dc-hint {}'>{}</div>".format("ok" if ok else "bad", text), unsafe_allow_html=True)


def panel(title, items, cls=""):
    st.markdown("<div class='dc-pn {}'><b>{}</b>{}</div>".format(cls, title, "".join("<div>{}</div>".format(x) for x in items)), unsafe_allow_html=True)


def cat_bar(text):
    st.markdown("<div class='dc-cat'>{}</div>".format(text), unsafe_allow_html=True)


def filter_row(pme=False):
    """Baris filter seperti mockup: Bulan, Tahun, Load, Export."""
    t = TODAY()
    dm = t.month - 1 if t.month > 1 else 12
    dy = t.year if t.month > 1 else t.year - 1
    c1, c2, c3, c4 = st.columns([2, 1.4, 0.7, 1.1])
    bln = c1.selectbox("Bulan", BLN, index=dm - 1, key="f_bln_" + ("pme" if pme else "pmi"))
    yrs = list(range(t.year - 3, t.year + 1))
    thn = c2.selectbox("Tahun", yrs, index=yrs.index(dy) if dy in yrs else len(yrs) - 1, key="f_thn_" + ("pme" if pme else "pmi"))
    c3.write("")
    c3.button("🔍 Load", type="primary", key="f_load_" + ("pme" if pme else "pmi"))
    return BLN.index(bln) + 1, int(thn), c4


def export_excel(pmi_yearly=None, pme=None):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        load_entries().to_excel(xw, sheet_name="Input", index=False)
        if pmi_yearly:
            pd.DataFrame(pmi_yearly).to_excel(xw, sheet_name="PMI_Tahunan", index=False)
        if pme:
            for k, nm in (("pm_rows", "PME_PM"), ("uk_rows", "PME_UkesKesesuaian"), ("kal_rows", "PME_Kalibrasi")):
                rows = [{kk: (str(v) if isinstance(v, (dt.date, list)) else v) for kk, v in r.items()} for r in pme[k]]
                pd.DataFrame(rows).to_excel(xw, sheet_name=nm, index=False)
    return bio.getvalue()


# =====================================================================
# 6. HALAMAN - DRACORE
# =====================================================================
def page_home():
    heading("Struktur Menu DRACORE - PMI vs PME", "Beranda &rsaquo; <b>DRACORE</b>", sub="Dashboard Radiology Assurance, Compliance & Operational Reliability Excellence · Mandaya Royal Hospital Puri · (angka merah = tugas yang masih perlu dikerjakan)")
    render(menu_html(), 700)


def page_dash_all():
    heading("Dashboard DRACORE (Manajemen)", "Beranda &rsaquo; DRACORE &rsaquo; <b>Dashboard</b>", sub="Ringkasan Quality, Compliance & Operational Reliability")
    t = TODAY()
    m = t.month - 1 if t.month > 1 else 12
    y = t.year if t.month > 1 else t.year - 1
    pmi = build_pmi(y, m)
    pme = build_pme(t, DUMMY())
    render(info_card() + BAN + "<div style='font-weight:bold;margin:4px 0 8px'>PMI &ndash; {} {}</div>".format(BLN[m - 1], y) + pmi["top"].split("<div class='sec'>")[0]
           + "<div style='font-weight:bold;margin:4px 0 8px'>PME &ndash; per {}</div>".format(dfmt(t)) + pme["kpi"] + pme["prio"], 700)


def page_rekap():
    heading("Rekap & Download Data Input", "Beranda &rsaquo; DRACORE &rsaquo; <b>Rekap</b>", sub="Seluruh data yang diinput melalui form PMI & PME")
    df = load_entries()
    if df.empty:
        st.info("Belum ada data input. Gunakan menu PMI / PME untuk menginput data.")
        return
    c1, c2 = st.columns(2)
    kat = c1.multiselect("Kategori", sorted(df["kategori"].unique()), default=sorted(df["kategori"].unique()))
    obj = c2.multiselect("Objek (alat / ruang)", sorted(df["objek"].unique()))
    v = df[df["kategori"].isin(kat)]
    if obj:
        v = v[v["objek"].isin(obj)]
    st.dataframe(v.drop(columns=["detail"]), use_container_width=True)
    c3, c4 = st.columns(2)
    c3.download_button("⬇️ Download CSV", v.to_csv(index=False).encode("utf-8"), "dracore_input.csv", "text/csv")
    c4.download_button("⬇️ Download Excel", export_excel(), "dracore_input.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.expander("Hapus semua data input (reset)"):
        if st.button("Hapus semua data", key="reset_all"):
            if os.path.exists(ENTRY_FILE):
                os.remove(ENTRY_FILE)
            st.success("Data input dihapus.")


# =====================================================================
# 7. HALAMAN - PMI
# =====================================================================
def page_suhu():
    heading("Input Suhu & Kelembapan Ruangan", "Beranda &rsaquo; PMI &rsaquo; Suhu & Kelembapan &rsaquo; <b>Input</b>", "Radiografer", sub="Pencatatan per shift (P/S/M), menggantikan Google Form")
    left, right = st.columns([3, 1.25], gap="large")
    with left:
        area = st.selectbox("Ruang / Area *", AREAS, index=2, key="su_area")
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal *", value=TODAY(), key="su_tgl")
        c2.time_input("Waktu pengukuran *", value=dt.time(8, 15), key="su_jam")
        shift = st.radio("Shift *", SHIFTS, horizontal=True, key="su_shift")
        c3, c4 = st.columns(2)
        suhu = c3.number_input("Suhu (°C) *", value=21.5, step=0.1, format="%.1f", key="su_suhu")
        hum = c4.number_input("Kelembapan (%) *", value=52.0, step=1.0, format="%.0f", key="su_hum")
        ok_t, ok_h = T_LO <= suhu <= T_HI, H_LO <= hum <= H_HI
        with c3:
            hint(ok_t, "✓ Dalam batas standar (18 – 23 °C)" if ok_t else "⚠ {} batas standar (18 – 23 °C)".format("DI ATAS" if suhu > T_HI else "DI BAWAH"))
        with c4:
            hint(ok_h, "✓ Dalam batas standar (40 – 60 %)" if ok_h else "⚠ {} batas standar (40 – 60 %)".format("DI ATAS" if hum > H_HI else "DI BAWAH"))
        st.text_input("Petugas (otomatis dari akun)", value="{} [{}]".format(petugas(), initials(petugas())), disabled=True, key="su_pet")
        abn = not (ok_t and ok_h)
        temuan, tindakan = "", ""
        if abn:
            temuan = st.text_area("Temuan * (wajib karena nilai di luar batas)", key="su_temuan")
            tindakan = st.radio("Tindakan korektif *", ["Lapor Elektromedis", "Cek AC / Dehumidifier", "Lainnya"], horizontal=True, key="su_tindakan")
        catatan = st.text_input("Catatan (opsional)", key="su_cat")
        code = shift[0]
        dup = any(x["objek"] == area and x["tgl"] == tgl and (x["shift"] or "")[:1] == code for x in recs("Suhu & Kelembapan"))
        if dup:
            st.error("Data shift {} untuk {} pada {} sudah ada (duplikasi ditolak).".format(shift, area, dfmt(tgl)))
        missing = abn and not temuan.strip()
        if missing:
            st.warning("Isi temuan terlebih dahulu sebelum menyimpan.")
        if st.button("⚠ Simpan & Kirim Notifikasi" if abn else "✓ Simpan", type="primary", disabled=dup or missing, key="su_save"):
            save_entry(kategori="Suhu & Kelembapan", objek=area, tanggal=tgl.isoformat(), shift=code, nilai1=suhu, nilai2=hum,
                       status="Abnormal" if abn else "Normal", petugas=petugas(), temuan=temuan or catatan, tindakan=tindakan,
                       detail={"notifikasi": "Supervisor Radiologi" if abn else ""})
            if abn:
                st.error("Status Abnormal tersimpan. Notifikasi dikirim ke Supervisor Radiologi.")
                st.toast("Notifikasi terkirim ke Supervisor Radiologi", icon="🔔")
            else:
                st.success("Data tersimpan, status ruangan = Normal. Dashboard PMI diperbarui.")
    with right:
        panel("Setelah disimpan (positif)", ["✓ Sistem validasi kelengkapan & batas", "✓ Status ruangan = <b style='display:inline'>Normal</b>", "✓ Titik muncul di grafik sesuai shift", "✓ Dashboard PMI diperbarui real-time"], "gr")
        panel("Aturan validasi (negatif)", ["⚠ Duplikasi shift yang sama ditolak", "⚠ Data wajib belum lengkap → tidak bisa disimpan", "⚠ Di luar batas → wajib isi temuan & tindakan", "⚠ Notifikasi ke <b style='display:inline'>Supervisor Radiologi</b>", "⚠ Status = <b style='display:inline'>Abnormal</b>, titik merah di grafik"], "rd")
    with st.expander("Live preview grafik bulan ini (ruang terpilih)", expanded=True):
        render(suhu_section_html(area, tgl.year, tgl.month, TODAY(), DUMMY()), 700)


def page_daily_qc():
    heading("Input Daily Quality Control Alat Radiologi", "Beranda &rsaquo; PMI &rsaquo; Daily Quality Control &rsaquo; <b>Input</b>", "Radiografer", sub="Checklist menyesuaikan jenis alat (template MRI Ambition X; alat lain memakai template sederhana - edit CHECKLISTS di app.py)")
    left, right = st.columns([3.2, 1], gap="large")
    with left:
        c1, c2, c3 = st.columns(3)
        mod = c1.selectbox("Modalitas *", DAILY_MODALITIES, index=5, key="qc_mod")
        tgl = c2.date_input("Tanggal *", value=TODAY(), key="qc_tgl")
        c3.text_input("Petugas (otomatis)", value="{} [{}]".format(petugas(), initials(petugas())), disabled=True, key="qc_pet")
        cats, total = flat_params(mod)
        ph = st.empty()
        answered, fails, errors, stl, detail_rows = 0, 0, [], [], []
        n = 0
        for cat, items in cats:
            cat_bar(cat)
            for k, p in items:
                a, b, c = st.columns([2.4, 2.6, 3])
                a.write(k)
                b.caption(p)
                key = "qc_{}_{}_{}".format(mod, tgl, n)
                if k in NUMERIC:
                    lo, hi, unit, dv = NUMERIC[k]
                    val = c.number_input(k, value=dv, step=0.5, key=key, label_visibility="collapsed")
                    ok = lo <= val <= hi
                    answered += 1
                    if ok:
                        c.markdown("<div class='dc-hint ok'>✓ sesuai ({:g}–{:g} {})</div>".format(lo, hi, unit), unsafe_allow_html=True)
                        stl.append("g")
                    else:
                        c.markdown("<div class='dc-hint bad'>⚠ di luar batas ({:g}–{:g} {})</div>".format(lo, hi, unit), unsafe_allow_html=True)
                        stl.append("r")
                        fails += 1
                        t1, t2 = st.columns(2)
                        tm = t1.text_input("Temuan *", key=key + "_t", placeholder="wajib diisi")
                        ta = t2.text_input("Tindakan awal *", key=key + "_a", placeholder="wajib diisi")
                        if not tm.strip() or not ta.strip():
                            errors.append("{}: temuan & tindakan awal wajib diisi".format(k))
                        detail_rows.append("{}: {} {}".format(k, val, unit))
                else:
                    ans = c.radio(k, ["Sesuai", "Tidak Sesuai", "N/A"], index=None, horizontal=True, key=key, label_visibility="collapsed")
                    if ans is None:
                        stl.append("g")
                        errors.append("{} belum diisi".format(k))
                    else:
                        answered += 1
                        if ans == "Tidak Sesuai":
                            stl.append("r")
                            fails += 1
                            t1, t2 = st.columns(2)
                            tm = t1.text_input("Temuan *", key=key + "_t", placeholder="wajib diisi")
                            ta = t2.text_input("Tindakan awal *", key=key + "_a", placeholder="wajib diisi")
                            if not tm.strip() or not ta.strip():
                                errors.append("{}: temuan & tindakan awal wajib diisi".format(k))
                            detail_rows.append("{}: {} / {}".format(k, tm, ta))
                        else:
                            stl.append("g")
                n += 1
        ph.markdown("Progres pengisian checklist: **{} / {}** parameter · status: {}".format(answered, total, "🟠 Draft" if answered < total else "🟢 Siap submit"))
        if fails:
            st.warning("{} parameter Tidak Sesuai → status QC akan menjadi **Not Compliant** dan notifikasi dikirim ke Supervisor Radiologi & Fisikawan Medis.".format(fails))
        if st.button("Submit - Not Compliant" if fails else "Submit", type="primary", key="qc_submit"):
            if errors:
                st.error("Lengkapi dulu:\n\n" + "\n".join("- " + e for e in errors[:12]))
            else:
                save_entry(kategori="Daily QC", objek=mod, tanggal=tgl.isoformat(), status="Not Compliant" if fails else "Completed", petugas=petugas(),
                           temuan="; ".join(detail_rows), detail={"st": stl, "fail": fails})
                if fails:
                    st.error("QC tersimpan dengan status Not Compliant. Notifikasi dikirim.")
                    st.toast("Notifikasi ke Supervisor Radiologi & Fisikawan Medis", icon="🔔")
                else:
                    st.success("QC harian tersimpan dengan status Completed.")
    with right:
        panel("Validasi saat Submit", ["✓ Semua parameter wajib terisi", "⚠ Parameter Tidak Sesuai → temuan & tindakan awal wajib", "⚠ Status QC = <b style='display:inline'>Not Compliant</b>", "⚠ Belum submit sampai batas waktu = <b style='display:inline'>Overdue</b>", "🔔 Notifikasi ke Supervisor Radiologi & Fisikawan Medis", "✓ Checklist menyesuaikan jenis alat", "✓ Histori tersimpan otomatis → sel merah di dashboard"], "am")


def page_qc_mingguan():
    heading("Input QC Mingguan", "Beranda &rsaquo; PMI &rsaquo; QC Mingguan &rsaquo; <b>Input</b>", "Fisikawan Medis", "#7a4fb5", "Image Quality Check, Air Calibration, Homogeneity sesuai jadwal dan jenis alat")
    left, right = st.columns([3.2, 1], gap="large")
    with left:
        c1, c2 = st.columns(2)
        mod = c1.selectbox("Modality *", [x[0] for x in WEEKLY_ROWS], key="wk_mod")
        tgl = c2.date_input("Tanggal pelaksanaan *", value=TODAY(), key="wk_tgl")
        fails, miss = 0, 0
        notes = []
        for i, it_ in enumerate(WEEKLY_SPEC[mod]):
            a, b, c = st.columns([2.2, 2, 2])
            a.write("**{}**".format(it_))
            res = b.radio(it_, ["Memenuhi", "Tidak Memenuhi"], index=None, horizontal=True, key="wk_{}_{}_{}".format(mod, tgl, i), label_visibility="collapsed")
            val = c.text_input("Nilai / catatan", key="wk_v_{}_{}_{}".format(mod, tgl, i), label_visibility="collapsed", placeholder="nilai ukur (opsional)")
            if res is None:
                miss += 1
            elif res == "Tidak Memenuhi":
                fails += 1
            notes.append("{}={}{}".format(it_, res, " ({})".format(val) if val else ""))
        up = st.file_uploader("Bukti pelaksanaan (bila dipersyaratkan)", key="wk_up_{}_{}".format(mod, tgl))
        tm = ta = ""
        if fails:
            st.error("Hasil tidak memenuhi batas penerimaan → status Failed. Catatan temuan & rencana tindak lanjut wajib diisi.")
            tm = st.text_area("Catatan temuan *", key="wk_tm")
            ta = st.text_area("Rencana tindak lanjut *", key="wk_ta")
        if st.button("Simpan Hasil - Failed" if fails else "Simpan Hasil", type="primary", key="wk_save"):
            if miss:
                st.error("{} item belum diisi.".format(miss))
            elif fails and (not tm.strip() or not ta.strip()):
                st.error("Temuan dan rencana tindak lanjut wajib diisi.")
            else:
                save_entry(kategori="QC Mingguan", objek=mod, tanggal=tgl.isoformat(), status="Failed" if fails else "Completed", petugas=petugas(), temuan=tm, tindakan=ta, lampiran=save_upload(up), detail={"hasil": notes})
                (st.error if fails else st.success)("QC mingguan tersimpan ({}).".format("Failed - reminder ke Fisikawan Medis & Supervisor" if fails else "Completed"))
    with right:
        panel("Alur sistem", ["✓ Hasil tersimpan & status kepatuhan diperbarui", "⚠ Failed / belum dilakukan sampai jatuh tempo → Overdue", "🔔 Reminder ke Fisikawan Medis & Supervisor Radiologi", "✓ Dashboard PMI Mingguan diperbarui"], "am")


def page_qc_bulanan():
    heading("Input QC Bulanan", "Beranda &rsaquo; PMI &rsaquo; QC Bulanan &rsaquo; <b>Input</b>", "Fisikawan Medis", "#7a4fb5", "Parameter sesuai daftar PMI Bulanan. Batas penerimaan numerik pada CT hanyalah ilustrasi - sesuaikan di BULANAN_SPEC.")
    left, right = st.columns([3.2, 1], gap="large")
    with left:
        c1, c2 = st.columns(2)
        mod = c1.selectbox("Modality *", list(BULANAN_SPEC.keys()), key="bl_mod")
        tgl = c2.date_input("Tanggal uji *", value=TODAY(), key="bl_tgl")
        spec = BULANAN_SPEC[mod]
        done = fails = 0
        for i, (p, lim) in enumerate(spec):
            a, b, c, d = st.columns([2.6, 1.4, 1.6, 1.2])
            a.write(p)
            key = "bl_{}_{}_{}".format(mod, tgl, i)
            if lim:
                lo, hi, unit, txt = lim
                b.caption("Batas: " + txt)
                val = c.number_input(p, value=float((lo + hi) / 2), step=0.1, key=key, label_visibility="collapsed")
                ok = lo <= val <= hi
                done += 1
                fails += 0 if ok else 1
                d.markdown(pill("Pass" if ok else "Failed"), unsafe_allow_html=True)
            else:
                b.caption("Pass / Failed")
                res = c.radio(p, ["Pass", "Failed"], index=None, horizontal=True, key=key, label_visibility="collapsed")
                if res:
                    done += 1
                    fails += 1 if res == "Failed" else 0
                    d.markdown(pill(res), unsafe_allow_html=True)
        st.caption("Parameter selesai: {} / {}".format(done, len(spec)))
        up = st.file_uploader("Bukti hasil pengujian *", key="bl_up_{}_{}".format(mod, tgl))
        kes = st.text_area("Kesimpulan" + (" * (wajib bila ada parameter Failed)" if fails else ""), key="bl_kes")
        if st.button("Simpan Hasil - Failed" if fails else "Simpan Hasil", type="primary", key="bl_save"):
            if done < len(spec):
                st.error("Semua parameter harus diisi.")
            elif up is None:
                st.error("Lampiran bukti hasil pengujian wajib diunggah.")
            elif fails and not kes.strip():
                st.error("Kesimpulan wajib diisi karena ada parameter Failed.")
            else:
                save_entry(kategori="QC Bulanan", objek=mod, tanggal=tgl.isoformat(), status="Failed" if fails else "Pass", petugas=petugas(), temuan=kes, lampiran=save_upload(up), detail={"done": done, "total": len(spec), "fails": fails})
                (st.error if fails else st.success)("QC bulanan tersimpan: {}.".format("Failed - notifikasi ke Supervisor & unit terkait" if fails else "Pass"))
    with right:
        panel("Alur sistem", ["✓ Hasil Pass/Failed dihitung otomatis dari batas", "✓ Lampiran bukti wajib diunggah", "⚠ Failed / lampiran kosong / lewat jadwal → temuan & tindak lanjut wajib", "🔔 Notifikasi ke Supervisor Radiologi & unit terkait", "✓ Dashboard PMI Bulanan: Pass / Failed / Overdue diperbarui"], "am")


def page_tahunan():
    heading("Input Uji Kebocoran Tabung X-Ray & Radiasi (Tahunan)", "Beranda &rsaquo; PMI &rsaquo; Uji Kebocoran Tahunan &rsaquo; <b>Input</b>", "PPR / Fisikawan Medis", "#c0392b")
    left, right = st.columns([3.2, 1], gap="large")
    with left:
        names = [x[0] for x in LEAK_ITEMS]
        alat = st.selectbox("Alat / area yang diuji *", names, key="ye_alat")
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal uji *", value=TODAY(), key="ye_tgl")
        hasil = c2.text_input("Hasil pengukuran *", key="ye_hasil", placeholder="mis. 0.8 mSv/jam")
        ok = st.radio("Kesimpulan *", ["Memenuhi", "Tidak Memenuhi"], horizontal=True, key="ye_ok")
        up = st.file_uploader("Laporan / sertifikat *", key="ye_up")
        st.date_input("Jadwal uji berikutnya (otomatis)", value=addm(tgl, 12), disabled=True, key="ye_next_{}".format(tgl))
        tm = st.text_area("Temuan keselamatan radiasi / corrective action" + (" *" if ok == "Tidak Memenuhi" else ""), key="ye_tm")
        if st.button("Simpan Uji Kebocoran", type="primary", key="ye_save"):
            if not hasil.strip() or up is None:
                st.error("Hasil pengukuran dan laporan/sertifikat wajib diisi.")
            elif ok == "Tidak Memenuhi" and not tm.strip():
                st.error("Temuan & corrective action wajib diisi.")
            else:
                save_entry(kategori="Uji Kebocoran", objek=alat, tanggal=tgl.isoformat(), nilai1=hasil, status=ok, petugas=petugas(), temuan=tm, lampiran=save_upload(up), detail={"next": addm(tgl, 12).isoformat()})
                (st.success if ok == "Memenuhi" else st.error)("Uji kebocoran tersimpan ({}). Jadwal berikutnya {}.".format(ok, dfmt(addm(tgl, 12))))
    with right:
        panel("Setelah disimpan", ["✓ Jadwal pengujian berikutnya dihitung otomatis", "✓ Status kepatuhan tampil di dashboard PMI Tahunan", "⚠ Tidak memenuhi / terlambat → Not Compliant / Overdue", "🔔 Notifikasi ke PPR & Supervisor Radiologi"], "gr")


def page_dash_pmi():
    heading("Dashboard PMI - Pemantapan Mutu Internal Radiologi", "Beranda &rsaquo; PMI &rsaquo; <b>Dashboard PMI</b>", sub="Pilih bulan dan tahun untuk menampilkan pelaksanaan PMI Harian, Mingguan, Bulanan, dan Tahunan beserta target dan achievement.")
    m, y, c_exp = filter_row()
    pmi = build_pmi(y, m)
    with c_exp:
        st.write("")
        st.download_button("📄 Export Excel", export_excel(pmi["yearly"]), "dashboard_pmi_{}_{:02d}.xlsx".format(y, m), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    render(info_card() + (BAN if DUMMY() else "") + legend_html(), 220)
    render(pmi["top"], 800)
    area = st.radio("Ruang / area (grafik Suhu & Kelembapan)", AREAS, index=2, horizontal=True, key="dp_area")
    render(suhu_section_html(area, y, m, TODAY(), DUMMY()), 700)
    mod = st.radio("Modalitas (checklist Daily QC)", DAILY_MODALITIES, index=5, horizontal=True, key="dp_mod")
    render(qc_section_html(mod, y, m, TODAY(), DUMMY()), 900)
    render(pmi["rest"], 900)


# =====================================================================
# 8. HALAMAN - PME
# =====================================================================
def page_pm():
    heading("Input Preventive Maintenance", "Beranda &rsaquo; PME &rsaquo; Preventive Maintenance &rsaquo; <b>Input</b>", "Teknisi Vendor / Elektromedis", "#0f8b8d", "Pencatatan PM 2-4 kali setahun per alat")
    left, right = st.columns([3, 1], gap="large")
    with left:
        c1, c2, c3 = st.columns(3)
        alat = c1.selectbox("Alat / Modality *", [x[0] for x in PM_ITEMS], index=3, key="pm_alat")
        c2.text_input("Jenis kegiatan", value="Preventive Maintenance", disabled=True, key="pm_jenis")
        tgl = c3.date_input("Tanggal pelaksanaan *", value=TODAY(), key="pm_tgl")
        c4, c5 = st.columns(2)
        pel = c4.selectbox("Pelaksana *", ["Teknisi Vendor (Philips)", "Teknisi Vendor (lainnya)", "Elektromedis RS"], key="pm_pel")
        pic = c5.selectbox("PIC Alat / Elektromedis *", ["Elektromedis RS", "PIC Alat Radiologi"], key="pm_pic")
        hasil = st.radio("Hasil pemeriksaan *", ["Baik / Normal", "Perlu Tindak Lanjut"], horizontal=True, key="pm_hasil")
        c6, c7 = st.columns(2)
        kerja = c6.text_area("Pekerjaan yang dilakukan *", key="pm_kerja")
        rek = c7.text_area("Rekomendasi teknis", key="pm_rek")
        up = st.file_uploader("Upload service report *", key="pm_up")
        nxt = st.date_input("Jadwal PM berikutnya (otomatis, interval 4 bulan - dapat diubah)", value=addm(tgl, 4), key="pm_next_{}".format(tgl))
        st.caption("✓ Reminder otomatis H-30 akan dijadwalkan")
        if st.button("✓ Simpan PM", type="primary", key="pm_save"):
            if not kerja.strip() or up is None:
                st.error("Pekerjaan yang dilakukan dan service report wajib diisi.")
            else:
                fu = hasil != "Baik / Normal"
                save_entry(kategori="Preventive Maintenance", objek=alat, tanggal=tgl.isoformat(), status="Need Follow-up" if fu else "Valid", petugas=pel + " / " + pic, temuan=kerja, tindakan=rek, lampiran=save_upload(up), detail={"next": nxt.isoformat()})
                (st.warning if fu else st.success)("PM tersimpan. {}".format("Status Need Follow-up - notifikasi ke PIC alat, Elektromedis, Supervisor." if fu else "Status alat = Valid, jadwal berikutnya " + dfmt(nxt)))
    with right:
        panel("Setelah disimpan", ["✓ Riwayat PM & service report masuk repository", "✓ Status alat = <b style='display:inline'>Valid</b> (kalender hijau)", "✓ Jadwal berikutnya dihitung otomatis", "✓ Dashboard PME diperbarui", "⚠ Perlu Tindak Lanjut / report belum diunggah → <b style='display:inline'>Need Follow-up</b> & notifikasi"], "gr")


def page_kalibrasi():
    heading("Input Kalibrasi Peralatan", "Beranda &rsaquo; PME &rsaquo; Kalibrasi &rsaquo; <b>Input</b>", "Elektromedis / PIC Alat", "#0f8b8d", "Tanggal kedaluwarsa dan reminder dihitung otomatis dari tanggal pelaksanaan")
    left, right = st.columns([3, 1], gap="large")
    with left:
        alat = st.selectbox("Alat *", _KAL_NAMES, index=14, key="kl_alat")
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal kalibrasi *", value=TODAY(), key="kl_tgl")
        nos = c2.text_input("No. Sertifikat *", key="kl_nos", placeholder="KAL/2026/015")
        ok = st.radio("Status kelayakan *", ["Layak", "Tidak Layak"], horizontal=True, key="kl_ok")
        tm = st.text_area("Temuan / catatan" + (" * (wajib bila Tidak Layak)" if ok == "Tidak Layak" else ""), key="kl_tm")
        up = st.file_uploader("Upload sertifikat *", key="kl_up")
        nx = st.date_input("Masa berlaku s.d. (otomatis, 1 tahun)", value=tgl + dt.timedelta(days=365), key="kl_nx_{}".format(tgl))
        st.caption("✓ Reminder kalibrasi berikutnya dijadwalkan")
        if st.button("⚠ Simpan & Kirim Notifikasi" if ok == "Tidak Layak" else "✓ Simpan Kalibrasi", type="primary", key="kl_save"):
            if not nos.strip() or up is None:
                st.error("No. sertifikat dan file sertifikat wajib diisi.")
            elif ok == "Tidak Layak" and not tm.strip():
                st.error("Temuan wajib diisi bila status Tidak Layak.")
            else:
                save_entry(kategori="Kalibrasi", objek=alat, tanggal=tgl.isoformat(), status="Not Compliant" if ok == "Tidak Layak" else "Valid", petugas=petugas(), temuan=tm, lampiran=save_upload(up), detail={"noser": nos, "next": nx.isoformat()})
                (st.error if ok == "Tidak Layak" else st.success)("Kalibrasi tersimpan. {}".format("Status Not Compliant - notifikasi ke Elektromedis, PIC alat, Supervisor." if ok == "Tidak Layak" else "Berlaku s.d. " + dfmt(nx)))
    with right:
        panel("Hasil pada dashboard", ["Tidak Layak → status <b style='display:inline'>Not Compliant</b>, notifikasi ke Elektromedis, PIC alat, Supervisor", "Alat dievaluasi sebelum digunakan kembali", "Sertifikat tersimpan di repository & dapat diaudit", "Lewat jatuh tempo → <b style='display:inline'>Overdue</b> + reminder bertingkat"], "am")


def page_uk():
    heading("Input Uji Kesesuaian", "Beranda &rsaquo; PME &rsaquo; Uji Kesesuaian &rsaquo; <b>Input</b>", "PPR / Fisikawan Medis / PIC Alat", "#0f8b8d", "Masa berlaku dihitung otomatis: 4 tahun (3 tahun untuk Fujifilm Digital Mammography FDR)")
    left, right = st.columns([3, 1], gap="large")
    with left:
        names = [x[0] for x in UK_ITEMS]
        alat = st.selectbox("Modality *", names, index=5, key="uk_alat")
        per = dict((x[0], x[2]) for x in UK_ITEMS)[alat]
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal uji *", value=TODAY(), key="uk_tgl")
        nos = c2.text_input("No. Laporan / Sertifikat *", key="uk_nos", placeholder="UK/RAD/2026/006")
        hasil = st.radio("Hasil uji *", ["Memenuhi", "Perlu Evaluasi"], horizontal=True, key="uk_hasil")
        up = st.file_uploader("Upload laporan / sertifikat *", key="uk_up")
        nx = st.date_input("Masa berlaku s.d. (otomatis, {} tahun)".format(per // 12), value=addm(tgl, per), key="uk_nx_{}_{}".format(tgl, per))
        st.caption("✓ Status Aktif · reminder bertingkat H-180 / H-90 / H-30")
        cat = st.text_input("Catatan", key="uk_cat")
        if st.button("✓ Simpan Uji Kesesuaian", type="primary", key="uk_save"):
            if not nos.strip() or up is None:
                st.error("No. laporan dan file laporan/sertifikat wajib diisi.")
            else:
                ne = hasil == "Perlu Evaluasi"
                save_entry(kategori="Uji Kesesuaian", objek=alat, tanggal=tgl.isoformat(), status="Need Evaluation" if ne else "Aktif", petugas=petugas(), temuan=cat, lampiran=save_upload(up), detail={"noser": nos, "next": nx.isoformat()})
                (st.error if ne else st.success)("Uji kesesuaian tersimpan. {}".format("Status Need Evaluation - alat perlu dievaluasi." if ne else "Aktif s.d. " + dfmt(nx)))
    with right:
        panel("Hasil pada dashboard", ["Status <b style='display:inline'>Aktif</b>, bar timeline hijau sampai masa berlaku", "Sertifikat tersimpan di repository & dapat diaudit", "Perlu Evaluasi → <b style='display:inline'>Need Evaluation</b> + notifikasi PIC & Supervisor", "Bila lewat jatuh tempo: <b style='display:inline'>Expired</b> + reminder bertingkat"], "gr")


def page_dash_pme():
    heading("Dashboard PME - Pemantapan Mutu Eksternal Radiologi", "Beranda &rsaquo; PME &rsaquo; <b>Dashboard PME</b>", sub="Monitoring jadwal Preventive Maintenance, masa berlaku Uji Kesesuaian, dan Kalibrasi beserta status, dokumen, dan reminder otomatis.")
    t = TODAY()
    pme = build_pme(t, DUMMY())
    c1, c2 = st.columns([4, 1])
    c1.caption("Per tanggal: **{}**".format(dfmt(t)))
    c2.download_button("📄 Export Excel", export_excel(None, pme), "dashboard_pme.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    render(info_card() + (BAN if DUMMY() else "") + legend_html(True) + pme["kpi"] + pme["prio"] + pme["pm"] + pme["uk"] + pme["kal"], 2400)


# =====================================================================
# 9. NAVIGASI (menu besar: DRACORE | PMI | PME)
# =====================================================================
pages = {
    "DRACORE": [
        st.Page(page_home, title="Struktur Menu", icon="🏠", url_path="home", default=True),
        st.Page(page_dash_all, title="Dashboard DRACORE", icon="📊", url_path="dashboard"),
        st.Page(page_rekap, title="Rekap & Download", icon="🗂️", url_path="rekap"),
    ],
    "PMI - Pemantapan Mutu Internal": [
        st.Page(page_suhu, title="Suhu & Kelembapan", icon="🌡️", url_path="pmi-suhu"),
        st.Page(page_daily_qc, title="Daily Quality Control", icon="✅", url_path="pmi-daily-qc"),
        st.Page(page_qc_mingguan, title="QC Mingguan", icon="📅", url_path="pmi-mingguan"),
        st.Page(page_qc_bulanan, title="QC Bulanan", icon="🗓️", url_path="pmi-bulanan"),
        st.Page(page_tahunan, title="Uji Kebocoran Tahunan", icon="☢️", url_path="pmi-tahunan"),
        st.Page(page_dash_pmi, title="Dashboard PMI", icon="📈", url_path="pmi-dashboard"),
    ],
    "PME - Pemantapan Mutu Eksternal": [
        st.Page(page_pm, title="Preventive Maintenance", icon="🛠️", url_path="pme-pm"),
        st.Page(page_uk, title="Uji Kesesuaian", icon="📋", url_path="pme-uk"),
        st.Page(page_kalibrasi, title="Kalibrasi", icon="⚖️", url_path="pme-kalibrasi"),
        st.Page(page_dash_pme, title="Dashboard PME", icon="📈", url_path="pme-dashboard"),
    ],
}
pg = st.navigation(pages)

with st.sidebar:
    st.markdown("---")
    st.markdown("<div style='background:linear-gradient(160deg,#1b4f86,#2a7ca8);color:#fff;border-radius:10px;padding:12px;text-align:center'>"
                "<b>{}</b><br><small>RADIOLOGY DEPARTMENT</small></div>".format(str(st.session_state.get("user_name", "Berliana")).upper()), unsafe_allow_html=True)
    st.text_input("Nama petugas (akun login)", value="Berliana", key="user_name")
    st.date_input("Tanggal simulasi", value=dt.date(2026, 10, 5), key="sim_today")
    st.checkbox("Tampilkan data dummy (simulasi)", value=True, key="use_dummy")
    st.caption("Mandaya Royal Hospital Puri · DRACORE v1.0")

pg.run()
