import pdfplumber
import json
import re

PDF_PATH = "/home/brsrk94/Desktop/srmm questionaiire extractor/brsr.pdf"
OUTPUT_PATH = "/home/brsrk94/Desktop/srmm questionaiire extractor/companies.json"

# ── Step 1: collect text lines, skipping page-header junk ────────────────────
SKIP_RE = re.compile(
    r"^(BRSR Reports for FY|S\.\s*No\.\s*Company|S\.No\.|Page \d+).*",
    re.IGNORECASE,
)

all_lines = []
with pdfplumber.open(PDF_PATH) as pdf:
    for page in pdf.pages:
        text = page.extract_text() or ""
        for line in text.splitlines():
            line = line.strip()
            if line and not SKIP_RE.match(line):
                all_lines.append(line)

# ── Step 2: helpers ───────────────────────────────────────────────────────────
URL_RE      = re.compile(r"https://\S+")
SNO_URL     = re.compile(r"^(\d+)\s+(https://\S+)\s*$")          # sno + url only
SNO_CO_URL  = re.compile(r"^(\d+)\s+(.+?)\s+(https://\S+)\s*$")  # sno + name + url
SNO_CO      = re.compile(r"^(\d+)\s+(.+)$")                      # sno + name (no url)

# words that can finish a company name on their own (multi-line wraps)
COMPLETION_WORDS = {
    "limited", "ltd", "ltd.", "limited.", "corporation", "inc", "inc.",
    "pvt", "pvt.", "llp", "llp."
}

def is_plain_text(line):
    return not re.match(r"^\d+\s", line) and not URL_RE.search(line)

def is_completion_word(line):
    """True when line is just a single suffix-completion like 'Limited' or 'LIMITED'."""
    words = line.strip().split()
    return len(words) == 1 and words[0].rstrip(".,").lower() in COMPLETION_WORDS

# ── Step 3: parse ─────────────────────────────────────────────────────────────
records = {}       # sno -> {company, brsr_link}
prev_consumed = False   # True when i-1 was already consumed as a suffix

i = 0
while i < len(all_lines):
    line = all_lines[i]

    # ── A: sno + url only → company name fully outside this line ─────────────
    m = SNO_URL.match(line)
    if m:
        sno = int(m.group(1))
        url = m.group(2)
        prefix = (all_lines[i - 1]
                  if i > 0 and not prev_consumed and is_plain_text(all_lines[i - 1])
                  else "")
        suffix = ""
        if i + 1 < len(all_lines) and is_plain_text(all_lines[i + 1]):
            suffix = all_lines[i + 1]
            i += 1
            prev_consumed = True
        else:
            prev_consumed = False
        company = " ".join(filter(None, [prefix, suffix])).strip()
        records[sno] = {"company": company, "brsr_link": url}
        i += 1
        continue

    # ── B: sno + company_part + url ───────────────────────────────────────────
    m = SNO_CO_URL.match(line)
    if m:
        sno     = int(m.group(1))
        co_part = m.group(2).strip()
        url     = m.group(3)
        # prefix: previous plain-text line only if NOT already consumed
        prefix = (all_lines[i - 1]
                  if i > 0 and not prev_consumed and is_plain_text(all_lines[i - 1])
                  else "")
        # suffix: only a short completion word (e.g. "LIMITED")
        suffix = ""
        if i + 1 < len(all_lines) and is_completion_word(all_lines[i + 1]):
            suffix = all_lines[i + 1]
            i += 1
            prev_consumed = True
        else:
            prev_consumed = False
        parts = [p for p in [prefix, co_part, suffix] if p]
        company = " ".join(parts).strip()
        records[sno] = {"company": company, "brsr_link": url}
        i += 1
        continue

    # ── C: sno + name, no url (null-link rows) ────────────────────────────────
    m = SNO_CO.match(line)
    if m:
        sno     = int(m.group(1))
        company = m.group(2).strip()
        if i + 1 < len(all_lines) and is_plain_text(all_lines[i + 1]):
            nxt = all_lines[i + 1]
            if not re.match(r"^\d{1,4}\s", nxt):
                company = (company + " " + nxt).strip()
                i += 1
                prev_consumed = True
            else:
                prev_consumed = False
        else:
            prev_consumed = False
        records[sno] = {"company": company, "brsr_link": None}
        i += 1
        continue

    prev_consumed = False
    i += 1

# ── Step 4: normalise and write ───────────────────────────────────────────────
NULL_URL = "https://nsearchives.nseindia.com/corporate/null"
companies = []
for sno in sorted(records):
    rec = records[sno]
    company = re.sub(r"\s+", " ", rec["company"]).strip()
    link = rec["brsr_link"]
    if not link or link == NULL_URL:
        link = None
    companies.append({"sno": sno, "company": company, "brsr_link": link})

with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    json.dump(companies, f, indent=2, ensure_ascii=False)

# ── Report ────────────────────────────────────────────────────────────────────
print(f"Extracted {len(companies)} companies → {OUTPUT_PATH}")
print(f"  First : #{companies[0]['sno']}  {companies[0]['company']}")
print(f"  Last  : #{companies[-1]['sno']}  {companies[-1]['company']}")
with_links = sum(1 for c in companies if c["brsr_link"])
print(f"  With BRSR links : {with_links} | Without : {len(companies) - with_links}")

spot = [1, 2, 23, 42, 50, 51, 78, 104, 542, 637, 1021]
print("\nSpot-check:")
for s in spot:
    hit = next((c for c in companies if c["sno"] == s), None)
    if hit:
        link_str = (hit["brsr_link"] or "null")[:50]
        print(f"  #{hit['sno']:4d}  {hit['company'][:60]:<60}  {link_str}")
    else:
        print(f"  #{s}  MISSING")

present = {c["sno"] for c in companies}
missing = [i for i in range(1, max(present) + 1) if i not in present]
print(f"\nMissing sno count: {len(missing)}" + (f"  → {missing}" if missing else " — none!"))
