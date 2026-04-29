import requests
import pandas as pd
import time
import os
from datetime import datetime

# =========================
# CONFIG
# =========================
url = "https://insp.social/xrpc/app.bsky.feed.getFeed?feed=at://did:plc:cadeggccwqtp3yjkk7auhpil/app.bsky.feed.generator/innovator&limit=50"

# 🔥 AMBIL TOKEN + BERSIHKAN (ANTI ERROR HEADER)
raw_token = os.getenv("INSPIRED_TOKEN", "")

token = raw_token.replace("\n", "").replace("\r", "").strip()

if not token:
    print("❌ TOKEN NOT FOUND / EMPTY")
    exit()

print("✅ Token loaded. Length:", len(token))

headers = {
    "Authorization": f"Bearer {token}"
}

# 🔥 BIAR PASTI KE-COMMIT (SELALU BERUBAH)
timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
file_path = f"inspired_data_{timestamp}.xlsx"

# =========================
# FETCH DATA
# =========================
all_rows = []
cursor = None

while True:
    full_url = url + (f"&cursor={cursor}" if cursor else "")

    res = requests.get(full_url, headers=headers)

    if res.status_code != 200:
        print(f"❌ API ERROR: {res.status_code}")
        print(res.text)
        break

    data = res.json()
    feed = data.get("feed", [])

    if not feed:
        break

    for item in feed:
        post = item.get("post", {})
        author = post.get("author", {})

        handle = author.get("handle", "")
        safe_handle = handle.split('.')[0] if handle and '.' in handle else handle

        all_rows.append({
            "Name": author.get("displayName"),
            "Handle": handle,
            "Profile": f"https://app.inspired.ch/profile/{safe_handle}" if handle else "",
            "Country": author.get("country"),
            "Sector ID": author.get("sectorId"),
            "Stage ID": author.get("stageId"),
            "User Type": author.get("userType"),
            "Description": post.get("record", {}).get("text") or "",
            "Likes": post.get("likeCount") or 0,
        })

    cursor = data.get("cursor")
    if not cursor:
        break

    time.sleep(1)

print(f"✅ Total data: {len(all_rows)}")

if len(all_rows) == 0:
    print("⚠️ No data fetched. Skip file creation.")
    exit()

# =========================
# SAVE TO EXCEL
# =========================
df = pd.DataFrame(all_rows)
df.to_excel(file_path, index=False)

# =========================
# STYLING EXCEL
# =========================
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

wb = load_workbook(file_path)
ws = wb.active

# HEADER
header_fill = PatternFill(start_color="111827", end_color="111827", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True, name="Segoe UI", size=11)

for cell in ws[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")

# BODY
body_font = Font(name="Segoe UI", size=10)

for row in ws.iter_rows(min_row=2):
    for cell in row:
        cell.font = body_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

# WIDTH
column_widths = {
    "Name": 28,
    "Description": 60,
    "Profile": 35,
    "Handle": 18,
    "Country": 12,
    "Sector ID": 14,
    "Stage ID": 14,
    "User Type": 14,
    "Likes": 10,
}

for idx, cell in enumerate(ws[1], start=1):
    col_letter = get_column_letter(idx)
    ws.column_dimensions[col_letter].width = column_widths.get(cell.value, 15)

# HEIGHT LIMIT
for row in ws.iter_rows(min_row=2):
    ws.row_dimensions[row[0].row].height = 60

# FREEZE
ws.freeze_panes = "A2"

# ZEBRA
stripe_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
    if i % 2 == 0:
        for cell in row:
            cell.fill = stripe_fill

# SECTOR COLOR
sector_colors = {
    "climate": "DCFCE7",
    "tech": "DBEAFE",
    "finance": "FEF3C7",
    "health": "FEE2E2",
}

sector_col_index = None
for idx, cell in enumerate(ws[1], start=1):
    if cell.value == "Sector ID":
        sector_col_index = idx
        break

if sector_col_index:
    for row in ws.iter_rows(min_row=2):
        cell = row[sector_col_index - 1]
        val = str(cell.value).lower() if cell.value else ""

        for key, color in sector_colors.items():
            if key in val:
                cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")

# SAVE FINAL
wb.save(file_path)

print(f"🔥 DONE — FILE CREATED: {file_path}")
