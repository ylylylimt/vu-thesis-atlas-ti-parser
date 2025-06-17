import re
import xml.etree.ElementTree as ET
import pandas as pd
from collections import defaultdict


name_of_paper_xml = "paper1.xml"
tree = ET.parse(name_of_paper_xml)
root = tree.getroot()

code_to_name = {}
code_to_tactic_override = {}

for c in root.findall("./codes/code"):
    cid  = c.attrib["id"]
    raw  = c.attrib["name"]
    m    = re.search(r"\s*\(T(\d+)\)\s*$", raw)
    if m:
        # strip off the " (Tn)"
        tn = m.group(1)
        name_stripped = raw[:m.start()].strip()
        code_to_name[cid] = name_stripped
        code_to_tactic_override[cid] = tn
    else:
        code_to_name[cid] = raw

# 3. Extract all quotations, record text, order, and (ATn) if it’s a title
quotes = []
for idx, q in enumerate(root.findall(".//primDoc//quotations/q")):
    qid    = q.attrib["id"]
    text   = "\n\n".join(p.text.strip() for p in q.findall("content/p") if p.text)
    m_at   = re.search(r"\(AT(\d+)\)", q.attrib["name"])
    at_num = m_at.group(1) if m_at else None
    quotes.append({"qid": qid, "text": text, "order": idx, "tactic": at_num})
quotes_by_id = {q["qid"]: q for q in quotes}

# Titles for fallback
title_quotes = sorted([q for q in quotes if q["tactic"]], key=lambda q: q["order"])
def find_tactic_for(qid):
    order = quotes_by_id[qid]["order"]
    for tq in reversed(title_quotes):
        if tq["order"] <= order:
            return tq["tactic"]
    return None

# 4. Pull in families
families = {
    cf.attrib["id"]: (
        cf.attrib["name"],
        [item.attrib["id"] for item in cf.findall("item")]
    )
    for cf in root.findall("./families/codeFamilies/codeFamily")
}

# 5. Read coding links and assign each code→tactic
tactic_codes = defaultdict(set)
for link in root.findall("./links/objectSegmentLinks/codings/iLink"):
    cid, qid = link.attrib["obj"], link.attrib["qRef"]
    # primary: override if code name had "(Tn)"
    if cid in code_to_tactic_override:
        tac = code_to_tactic_override[cid]
    else:
        # fallback: nearest preceding ATn quotation
        tac = find_tactic_for(qid)
    if tac:
        tactic_codes[tac].add(cid)

# 6. Build output rows: one tactic → { familyName: [codeNames], ... }
output = {}
for tac, cids in tactic_codes.items():
    row = {}
    for fam_id, (fam_name, fam_code_ids) in families.items():
        hits = sorted(cids & set(fam_code_ids))
        row[fam_name] = "; ".join(code_to_name[c] for c in hits) if hits else ""
    # grab the title paragraph
    title_q = next(q for q in title_quotes if q["tactic"] == tac)
    row["Paragraph"] = title_q["text"]
    output[tac] = row

# 7. DataFrame, reorder columns
df = pd.DataFrame.from_dict(output, orient="index")
df.index.name = "Tactic"

cols = [
    "1. Title",
    "2. Description",
    "3. Participant",
    "4. Related Software Artifact",
    "5. Context",
    "6. Software Feature",
    "7. Tactic Intent",
    "8. Target Quality Attribute",
    "9. Other Related Quality Attributes",
    "10. Measured Impact",
    "Paragraph"
]
df = df.reindex(columns=cols)

# 8. Write nicely to Excel
output_path = "output.xlsx"
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Tactics", startrow=1, header=False)

    wb = writer.book
    ws = writer.sheets["Tactics"]

    hdr = wb.add_format({
        "bold": True, "bg_color": "#D7E4BC",
        "border": 1,  "text_wrap": True,
        "align": "center", "valign": "vcenter"
    })
    wrap = wb.add_format({"text_wrap": True, "valign": "top"})

    headers = [df.index.name] + df.columns.tolist()
    for i, h in enumerate(headers):
        ws.write(0, i, h, hdr)

    for i, col in enumerate(headers):
        if i == 0:
            w = max(df.index.astype(str).map(len).max(), len(col)) + 2
        else:
            w = max(df[col].astype(str).map(len).max(), len(col)) + 2
        ws.set_column(i, i, w, wrap)

    ws.freeze_panes(1, 1)

print(f"Wrote formatted file → {output_path}")
