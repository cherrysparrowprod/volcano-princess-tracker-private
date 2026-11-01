"""Convert the Volcano Princess spreadsheet into data.js for the tracker app."""
import json
import re
import sys
import openpyxl

SRC = sys.argv[1]
OUT = sys.argv[2]

ws = openpyxl.load_workbook(SRC).active
STATS = ["Stamina", "Mind", "Charisma", "STR", "INT", "EMO", "IMG", "Sword", "Archery",
         "Science", "Beliefs", "Ritual", "Literature", "Art", "Tone", "Fame", "Darkness"]


def cell(r, c):
    v = ws.cell(r, c).value
    return v.strip() if isinstance(v, str) else v


def split_conditions(text):
    if not text:
        return []
    parts = re.split(r",\s+|\s+\+\s+", str(text).strip())
    return [p[0].upper() + p[1:] for p in (p.strip() for p in parts) if p]


def stat_req(name, v):
    if isinstance(v, (int, float)):
        return f"{name} {int(v):,}+"
    v = str(v).strip()
    m = re.fullmatch(r"(\d+)\s*(>=|<=)", v)
    if m and name == "Darkness":
        # Flipped reading confirmed: "30>=" means 30 or more, "80<=" means 80 or less
        n, op = int(m.group(1)), m.group(2)
        return f"Darkness {n:,} or more" if op == ">=" else f"Darkness {n:,} or less"
    m = re.fullmatch(r"Avg (\d+) \((\d+) total\)", v)
    if m:
        return f"Average of STR, INT, EMO, IMG {int(m.group(1)):,}+ ({int(m.group(2)):,} total)"
    m = re.fullmatch(r"Below Avg (\d+) \(< ?(\d+) total\)", v)
    if m:
        return f"Average of STR, INT, EMO, IMG below {int(m.group(1)):,} (under {int(m.group(2)):,} total)"
    raise ValueError(f"Unrecognised value for {name}: {v!r}")


# ---- Romance (rows 6-32) ----
romances, seen = [], set()
for r in range(6, 33):
    name, done, cond = cell(r, 1), cell(r, 2), cell(r, 3)
    if not name:
        continue
    if name == "Lebsa":
        label = "Lebsa (forgive)" if "Forgive" in cond else "Lebsa (punish)"
    else:
        label = name
    if label in seen:          # duplicate Connie / Jermaine rows
        continue
    seen.add(label)
    if r >= 27:                # group endings: condition lists the people
        people, extra = [p.strip() for p in cond.split("+")], []
    else:
        people, extra = [name], split_conditions(cond)
    romances.append({"id": f"rom{len(romances)}", "name": label, "people": people,
                     "extra": extra, "group": r >= 27, "done": done is True})

# ---- Careers (rows 48-129) ----
careers, path, sub = [], None, None
for r in range(48, 130):
    a = cell(r, 1)
    if not a:
        continue
    if cell(r, 3) == "Star":
        path, sub = a, None
        continue
    if not isinstance(ws.cell(r, 3).value, (int, float)):
        sub = a
        continue
    reqs = [stat_req(s, cell(r, 4 + i)) for i, s in enumerate(STATS) if cell(r, 4 + i) not in (None, "")]
    reqs += split_conditions(cell(r, 21))
    careers.append({"id": f"car{len(careers)}", "name": a, "path": path, "sub": sub,
                    "stars": int(cell(r, 3)), "reqs": reqs, "done": cell(r, 2) is True})

# ---- One-off achievements ----
sections = {"Dad": (36, 44), "Growth": (133, 187), "Events": (191, 204), "Social": (208, 233),
            "Combat": (237, 265), "Other": (269, 275), "Story": (279, 289)}
achievements = []
for cat, (a, b) in sections.items():
    for r in range(a, b + 1):
        if cell(r, 1) and ws.cell(r, 2).value is not None:
            achievements.append({"id": f"{cat.lower()}{r}", "cat": cat, "name": cell(r, 1),
                                 "desc": cell(r, 3) or "", "done": cell(r, 2) is True})

data = {"romances": romances, "careers": careers, "achievements": achievements}
with open(OUT, "w", encoding="utf-8") as f:
    f.write("// Generated from volcano_princess_100.xlsx\nwindow.VP_DATA = ")
    json.dump(data, f, ensure_ascii=False, indent=1)
    f.write(";\n")

print(f"Romances {sum(x['done'] for x in romances)}/{len(romances)}, "
      f"careers {sum(x['done'] for x in careers)}/{len(careers)}, "
      f"achievements {sum(x['done'] for x in achievements)}/{len(achievements)}")
