"""Analisa transisi hasil 'suspects' per Name.

Pemakaian: python scripts/analisa_suspect.py <input.xlsx> <output.xlsx>
"""
import sys
import datetime as dt
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC, OUT = sys.argv[1], sys.argv[2]
SEV = {"Neg": 0, "suspects": 1, "Positive": 2}
LABEL = {"Neg": "Neg", "suspects": "Suspect", "Positive": "Positive"}

# ---------- 1. Bersihkan data ----------
raw = pd.read_excel(SRC)
raw["Catatan"] = ""


def to_num(v):
    # Beberapa nilai Result Range tersimpan sebagai tanggal (format sel Excel salah);
    # dikembalikan ke angka serial Excel aslinya (mis. 1899-12-29 16:48 -> -0.3).
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return round((pd.Timestamp(v) - pd.Timestamp("1899-12-30")) / pd.Timedelta(days=1), 1)
    return float(v)


is_date = raw["Result Range"].apply(lambda v: isinstance(v, (dt.datetime, pd.Timestamp)))
raw["Result Range"] = raw["Result Range"].apply(to_num)
raw.loc[is_date, "Catatan"] = "Result Range terformat tanggal, dikonversi ke angka"
raw["Hasil"] = raw["Hasil"].astype(str).str.strip()
raw["sev"] = raw["Hasil"].map(SEV)
assert raw["sev"].notna().all(), raw.loc[raw["sev"].isna(), "Hasil"].unique()

# Satu hasil per Name per tanggal: jika ada duplikat, ambil yang paling berat
# (Positive > suspects > Neg, lalu Result Range tertinggi).
raw = raw.sort_values(["Name", "Sampling Date", "sev", "Result Range"])
dup_n = raw.groupby(["Name", "Sampling Date"])["Hasil"].transform("size")
data = raw.groupby(["Name", "Sampling Date"], as_index=False).tail(1).copy()
data.loc[dup_n.loc[data.index] > 1, "Catatan"] = (
    data["Catatan"].where(data["Catatan"] == "", data["Catatan"] + "; ")
    + "Duplikat tanggal sama digabung (diambil hasil terberat)"
)
data = data.sort_values(["Name", "Sampling Date"]).reset_index(drop=True)
n_dup_removed = len(raw) - len(data)

# ---------- 2. Info per Name ----------
data["urut"] = data.groupby("Name").cumcount()
g = data.groupby("Name")
last = g.tail(1).set_index("Name")


def bucket(r):
    if r["Hasil"] != "suspects":
        return None
    if 10 <= r["Result Range"] < 15:
        return "10-14.9"
    if 15 <= r["Result Range"] <= 20:
        return "15-20"
    return None


data["Bucket"] = data.apply(bucket, axis=1)
nxt = g.shift(-1)
data["Tgl Berikutnya"] = nxt["Sampling Date"]
data["RR Berikutnya"] = nxt["Result Range"]
data["Hasil Berikutnya"] = nxt["Hasil"].map(LABEL)


def detail(bk):
    ev = data[data["Bucket"] == bk]
    first = ev.groupby("Name").head(1).set_index("Name")
    cnt = ev.groupby("Name").size()
    rows = []
    for name, r in first.iterrows():
        lr = last.loc[name]
        has_next = pd.notna(r["Tgl Berikutnya"])
        rows.append([
            name, r["Sampling Date"].to_pydatetime(), r["Result Range"], int(cnt[name]),
            r["Tgl Berikutnya"].to_pydatetime() if has_next else None,
            r["RR Berikutnya"] if has_next else None,
            r["Hasil Berikutnya"] if has_next else "Belum ada sampling berikutnya",
            lr["Sampling Date"].to_pydatetime(), lr["Result Range"],
            LABEL[lr["Hasil"]] if has_next else "Belum ada sampling berikutnya",
        ])
    return rows


det = {bk: detail(bk) for bk in ("10-14.9", "15-20")}

# Semua kejadian suspect (tanpa dedup per Name) – sebagai pembanding
events = data[data["Bucket"].notna()]
ev_rows = [[r["Name"], r["Sampling Date"].to_pydatetime(), r["Result Range"], r["Bucket"],
            r["Tgl Berikutnya"].to_pydatetime() if pd.notna(r["Tgl Berikutnya"]) else None,
            r["RR Berikutnya"] if pd.notna(r["Tgl Berikutnya"]) else None,
            r["Hasil Berikutnya"] if pd.notna(r["Tgl Berikutnya"]) else "Belum ada sampling berikutnya"]
           for _, r in events.iterrows()]

# Per Name
per = pd.DataFrame({
    "Jumlah Sampling": g.size(),
    "Tgl Pertama": g["Sampling Date"].min(),
    "Tgl Terakhir": g["Sampling Date"].max(),
    "Jml Neg": g["Hasil"].apply(lambda s: (s == "Neg").sum()),
    "Jml Suspect": g["Hasil"].apply(lambda s: (s == "suspects").sum()),
    "Jml Positive": g["Hasil"].apply(lambda s: (s == "Positive").sum()),
    "Result Range Terakhir": last["Result Range"],
    "Hasil Akhir": last["Hasil"].map(LABEL),
    "Pernah Suspect 10-14.9": g["Bucket"].apply(lambda s: "Ya" if (s == "10-14.9").any() else "Tidak"),
    "Pernah Suspect 15-20": g["Bucket"].apply(lambda s: "Ya" if (s == "15-20").any() else "Tidak"),
}).reset_index()

# ---------- 3. Tulis workbook ----------
F = "Arial"
HDR_FILL = PatternFill("solid", start_color="1F4E78")
SUB_FILL = PatternFill("solid", start_color="DDEBF7")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
wb = Workbook()


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, 1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font = Font(name=F, bold=True, color="FFFFFF")
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=1)
    ws.auto_filter.ref = f"A{row}:{get_column_letter(len(cols))}{row}"


def write_rows(ws, start, rows, fmts):
    for ri, r in enumerate(rows, start):
        for ci, v in enumerate(r, 1):
            c = ws.cell(row=ri, column=ci, value=v)
            c.font = Font(name=F)
            if fmts.get(ci):
                c.number_format = fmts[ci]


# Ringkasan dibuat pertama (sheet aktif)
ws = wb.active
ws.title = "Ringkasan"

# Detail sheets
DET_COLS = ["Name", "Tgl Suspect (pertama di range)", "Result Range Suspect", "Jml Suspect di Range",
            "Tgl Sampling Berikutnya", "Result Range Berikutnya", "Hasil Sampling Berikutnya",
            "Tgl Sampling Terakhir", "Result Range Terakhir", "Hasil Akhir (sampling terbaru)"]
DET_W = [12, 16, 12, 12, 16, 13, 26, 16, 13, 26]
DFMT = {2: "yyyy-mm-dd", 3: "0.0", 5: "yyyy-mm-dd", 6: "0.0", 8: "yyyy-mm-dd", 9: "0.0"}
det_sheet = {"10-14.9": "Suspect 10-14.9", "15-20": "Suspect 15-20"}
for bk, name in det_sheet.items():
    s = wb.create_sheet(name)
    header(s, 1, DET_COLS, DET_W)
    write_rows(s, 2, det[bk], DFMT)

s = wb.create_sheet("Semua Kejadian Suspect")
header(s, 1, ["Name", "Tgl Suspect", "Result Range", "Range", "Tgl Sampling Berikutnya",
              "Result Range Berikutnya", "Hasil Sampling Berikutnya"], [12, 14, 12, 10, 16, 13, 26])
write_rows(s, 2, ev_rows, {2: "yyyy-mm-dd", 3: "0.0", 5: "yyyy-mm-dd", 6: "0.0"})

s = wb.create_sheet("Per Name")
header(s, 1, ["Name"] + list(per.columns[1:]), [12, 10, 13, 13, 9, 11, 11, 12, 12, 12, 12])
write_rows(s, 2, [[r[0], int(r[1]), r[2].to_pydatetime(), r[3].to_pydatetime(), int(r[4]), int(r[5]), int(r[6]),
                   float(r[7]), r[8], r[9], r[10]] for r in per.itertuples(index=False)],
           {3: "yyyy-mm-dd", 4: "yyyy-mm-dd", 8: "0.0"})

s = wb.create_sheet("Data Bersih")
header(s, 1, ["Sampling Date", "Name", "Result Range", "Hasil", "Catatan"], [14, 12, 12, 11, 55])
write_rows(s, 2, [[r["Sampling Date"].to_pydatetime(), r["Name"], r["Result Range"], r["Hasil"], r["Catatan"] or None]
                  for _, r in data.iterrows()], {1: "yyyy-mm-dd", 3: "0.0"})

# ---------- Ringkasan (pakai formula COUNTIFS ke sheet detail) ----------
ws.column_dimensions["A"].width = 44
for col in "BCDE":
    ws.column_dimensions[col].width = 16
ws["A1"] = "Analisa Transisi Hasil Suspect per Name"
ws["A1"].font = Font(name=F, bold=True, size=14)
ws["A2"] = ("1 Name dihitung 1 kali per kelompok range. Titik awal = suspect pertama Name tsb di range itu. "
            "Persentase dihitung dari Name yang punya sampling sesudah suspect.")
ws["A2"].font = Font(name=F, italic=True, size=9)

OUTCOMES = ["Positive", "Neg", "Suspect"]
r = 4
for bk, sname in det_sheet.items():
    q = f"'{sname}'"
    ws.cell(row=r, column=1, value=f"Suspect dengan Result Range {bk.replace('-', ' – ').replace('.', ',')}").font = Font(name=F, bold=True, size=12)
    r += 1
    for metric, col in (("Hasil pada sampling date berikutnya", "G"), ("Hasil akhir (sampling date terbaru)", "J")):
        hdr = [metric, "Jumlah Name", "Persentase"]
        for i, h in enumerate(hdr, 1):
            c = ws.cell(row=r, column=i, value=h)
            c.font = Font(name=F, bold=True, color="FFFFFF")
            c.fill = HDR_FILL
            c.border = BOX
            c.alignment = Alignment(horizontal="center" if i > 1 else "left")
        top = r + 1
        for k, oc in enumerate(OUTCOMES):
            rr = top + k
            lab = {"Positive": "Menjadi Positive", "Neg": "Menjadi Neg", "Suspect": "Tetap Suspect"}[oc]
            ws.cell(row=rr, column=1, value=lab)
            ws.cell(row=rr, column=2, value=f'=COUNTIFS({q}!${col}:${col},"{oc}")')
            ws.cell(row=rr, column=3, value=f"=IF($B${top + 3}=0,0,B{rr}/$B${top + 3})")
        tr = top + 3
        ws.cell(row=tr, column=1, value="Total Name dengan sampling berikutnya")
        ws.cell(row=tr, column=2, value=f"=SUM(B{top}:B{top + 2})")
        ws.cell(row=tr, column=3, value=f"=SUM(C{top}:C{top + 2})")
        ws.cell(row=tr + 1, column=1, value="Belum ada sampling berikutnya (tidak dihitung di %)")
        ws.cell(row=tr + 1, column=2, value=f'=COUNTIFS({q}!${col}:${col},"Belum ada sampling berikutnya")')
        ws.cell(row=tr + 2, column=1, value="Total Name suspect di range ini")
        ws.cell(row=tr + 2, column=2, value=f"=COUNTA({q}!$A:$A)-1")
        for rr in range(top, tr + 3):
            for cc in (1, 2, 3):
                c = ws.cell(row=rr, column=cc)
                c.font = Font(name=F, bold=rr in (tr, tr + 2))
                c.border = BOX
                if rr == tr:
                    c.fill = SUB_FILL
            ws.cell(row=rr, column=2).number_format = "#,##0"
            ws.cell(row=rr, column=3).number_format = "0.0%"
        r = tr + 4
    r += 1

# Pembanding: semua kejadian (tanpa dedup per Name)
ws.cell(row=r, column=1, value="Pembanding: semua kejadian suspect (Name bisa dihitung lebih dari 1x)").font = Font(name=F, bold=True, size=12)
r += 1
hdr = ["Hasil sampling berikutnya", "10–14,9 (jml)", "10–14,9 (%)", "15–20 (jml)", "15–20 (%)"]
for i, h in enumerate(hdr, 1):
    c = ws.cell(row=r, column=i, value=h)
    c.font = Font(name=F, bold=True, color="FFFFFF")
    c.fill = HDR_FILL
    c.border = BOX
top = r + 1
q = "'Semua Kejadian Suspect'"
for k, oc in enumerate(OUTCOMES):
    rr = top + k
    ws.cell(row=rr, column=1, value={"Positive": "Menjadi Positive", "Neg": "Menjadi Neg", "Suspect": "Tetap Suspect"}[oc])
    for bcol, pcol, bk in (("B", "C", "10-14.9"), ("D", "E", "15-20")):
        ws[f"{bcol}{rr}"] = f'=COUNTIFS({q}!$D:$D,"{bk}",{q}!$G:$G,"{oc}")'
        ws[f"{pcol}{rr}"] = f"=IF({bcol}${top + 3}=0,0,{bcol}{rr}/{bcol}${top + 3})"
tr = top + 3
ws.cell(row=tr, column=1, value="Total kejadian dengan sampling berikutnya")
for bcol, pcol in (("B", "C"), ("D", "E")):
    ws[f"{bcol}{tr}"] = f"=SUM({bcol}{top}:{bcol}{top + 2})"
    ws[f"{pcol}{tr}"] = f"=SUM({pcol}{top}:{pcol}{top + 2})"
for rr in range(top, tr + 1):
    for cc in range(1, 6):
        c = ws.cell(row=rr, column=cc)
        c.font = Font(name=F, bold=rr == tr)
        c.border = BOX
        c.number_format = "0.0%" if cc in (3, 5) else "#,##0"
        if rr == tr:
            c.fill = SUB_FILL
r = tr + 3

# Catatan metodologi
notes = [
    "Catatan & asumsi",
    f"• Data sumber: {len(raw):,} baris, {data['Name'].nunique():,} Name, periode "
    f"{data['Sampling Date'].min():%Y-%m-%d} s.d. {data['Sampling Date'].max():%Y-%m-%d}.",
    "• Kategori suspect ditentukan dari kolom Hasil = 'suspects', lalu dikelompokkan menurut Result Range: "
    "10 ≤ RR < 15 (10–14,9) dan 15 ≤ RR ≤ 20 (15–20).",
    "• 1 Name = 1 hasil: per kelompok range, tiap Name dihitung sekali, dengan titik awal suspect pertama di range tsb. "
    "Name yang suspect di kedua range muncul di kedua tabel.",
    "• 'Sampling date berikutnya' = tanggal sampling pertama setelah tanggal suspect tsb.",
    "• 'Hasil akhir' = hasil pada sampling date paling baru milik Name tsb.",
    f"• {n_dup_removed} baris duplikat (Name + tanggal sama) digabung jadi 1 hasil per tanggal; bila hasil berbeda, "
    "diambil yang terberat (Positive > Suspect > Neg). Hanya 1 kasus yang berbeda (Name 1242290, 2026-01-12: Positive 57 & Neg 2).",
    f"• {int(is_date.sum())} nilai Result Range tersimpan sebagai tanggal di file sumber; dikonversi kembali ke angka "
    "(semuanya Hasil = Neg, tidak mempengaruhi analisa suspect).",
    "• Name yang suspect-nya adalah sampling terakhir (belum ada sampling berikutnya) tidak masuk penyebut persentase.",
]
for i, t in enumerate(notes):
    c = ws.cell(row=r + i, column=1, value=t)
    c.font = Font(name=F, bold=i == 0, size=10 if i else 11)
for row in ws.iter_rows(min_row=3, max_row=r - 1):
    for c in row:
        if c.font.name != F:
            c.font = Font(name=F, bold=c.font.b, color=c.font.color)

wb.save(OUT)
print("saved", OUT)
