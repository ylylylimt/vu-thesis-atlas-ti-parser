import re
import xml.etree.ElementTree as ET
import pandas as pd
from collections import defaultdict

# 1. Load the XML
name_of_paper_xml = "paper1.xml"
tree = ET.parse(name_of_paper_xml)
root = tree.getroot()

# 2. Build mapping of code IDs
code_to_name = {}
code_to_tactic_override = {}

for c in root.findall("./codes/code"):
    cid = c.attrib["id"]
    raw = c.attrib["name"]
    m   = re.search(r"\s*\(T(\d+)\)\s*$", raw)
    if m:
        tn = m.group(1)
        code_to_name[cid]              = raw[:m.start()].strip()
        code_to_tactic_override[cid]   = tn
    else:
        code_to_name[cid] = raw

# 3. Extract quotations to support ATn fallback
quotes = []
for idx, q in enumerate(root.findall(".//primDoc//quotations/q")):
    qid  = q.attrib["id"]
    m_at = re.search(r"\(AT(\d+)\)", q.attrib["name"])
    atn  = m_at.group(1) if m_at else None
    quotes.append({"qid": qid, "order": idx, "tactic": atn})
quotes_by_id = {q["qid"]: q for q in quotes}

title_quotes = sorted([q for q in quotes if q["tactic"]], key=lambda q: q["order"])
def find_tactic_for(qid):
    order = quotes_by_id[qid]["order"]
    for tq in reversed(title_quotes):
        if tq["order"] <= order:
            return tq["tactic"]
    return None

# 4. Read codeFamily (cf) definitions
families = {
    cf.attrib["id"]: (
        cf.attrib["name"],
        [item.attrib["id"] for item in cf.findall("item")]
    )
    for cf in root.findall("./families/codeFamilies/codeFamily")
}

tactic_codes = defaultdict(set)
for link in root.findall("./links/objectSegmentLinks/codings/iLink"):
    cid, qid = link.attrib["obj"], link.attrib["qRef"]
    tac = code_to_tactic_override.get(cid) or find_tactic_for(qid)
    if tac:
        tactic_codes[tac].add(cid)

# 5. Build output rows
output = {}
for tac, cids in tactic_codes.items():
    row = {}
    for fam_id, (fam_name, fam_code_ids) in families.items():
        hits = sorted(cids & set(fam_code_ids))
        row[fam_name] = "; ".join(code_to_name[c] for c in hits) if hits else ""
    output[tac] = row

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
    "11. Level of abstraction",
    "12. Tool or framework"
]
df = df.reindex(columns=cols)

df = df.sort_index(key=lambda idx: idx.astype(int))

# 6. Write to Excel
output_path = "tactics_overview_12cols.xlsx"
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    df.to_excel(writer, sheet_name="Tactics", startrow=1, header=False)

    wb = writer.book
    ws = writer.sheets["Tactics"]

    hdr = wb.add_format({
        "bold": True,
        "bg_color": "#D7E4BC",
        "border": 1,
        "text_wrap": True,
        "align": "center",
        "valign": "vcenter"
    })
    wrap = wb.add_format({"text_wrap": True, "valign": "top"})

    headers = [df.index.name] + df.columns.tolist()
    for col_idx, header in enumerate(headers):
        ws.write(0, col_idx, header, hdr)

    for col_idx, col in enumerate(headers):
        if col_idx == 0:
            width = max(df.index.astype(str).map(len).max(), len(col)) + 2
        else:
            width = max(df[col].astype(str).map(len).max(), len(col)) + 2
        ws.set_column(col_idx, col_idx, width, wrap)

    ws.freeze_panes(1, 1)

print(f"Wrote formatted file → {output_path}")
