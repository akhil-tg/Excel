"""Scratch: assign regions N/S/E/W to the 31 sales rows hitting exact region totals.
DP over 4-tuple states (amounts in units of 500)."""

rows = [  # (tag, month, person, product, amount)
    ("L1", "Aug-26", "Sneha", "Laptops", 38000),
    ("L2", "Aug-26", "Sneha", "Laptops", 42000),
    ("L3", "Aug-26", "Sneha", "Laptops", 48000),
    ("L4", "Aug-26", "Divya", "Laptops", 45000),
    ("L5", "Sep-26", "Divya", "Laptops", 55000),
    ("L6", "Sep-26", "Priya", "Laptops", 43500),
    ("L7", "Sep-26", "Arjun", "Laptops", 38500),
    ("P1", "Sep-26", "Arjun", "Printers", 34000),
    ("P2", "Sep-26", "Priya", "Printers", 34000),
    ("P3", "Sep-26", "Rahul", "Printers", 40000),
    ("P4", "Sep-26", "Sneha", "Printers", 40000),
    ("P5", "Oct-26", "Vikram", "Printers", 32000),
    ("P6", "Oct-26", "Divya", "Printers", 40000),
    ("N1", "Oct-26", "Arjun", "Notebooks", 45000),
    ("N2", "Aug-26", "Priya", "Notebooks", 54000),
    ("N3", "Aug-26", "Sneha", "Notebooks", 67500),
    ("N4", "Oct-26", "Priya", "Notebooks", 27000),
    ("N5", "Oct-26", "Rahul", "Notebooks", 31500),
    ("N6", "Oct-26", "Divya", "Notebooks", 25000),
    ("N7", "Sep-26", "Vikram", "Notebooks", 23000),
    ("N8", "Sep-26", "Sneha", "Notebooks", 11000),
    ("I1", "Oct-26", "Arjun", "Ink & Toner", 40500),
    ("I2", "Oct-26", "Priya", "Ink & Toner", 34000),
    ("I3", "Aug-26", "Rahul", "Ink & Toner", 25500),
    ("I4", "Oct-26", "Sneha", "Ink & Toner", 10500),
    ("I5", "Sep-26", "Vikram", "Ink & Toner", 34000),
    ("K1", "Oct-26", "Vikram", "Stationery Kits", 35000),
    ("K2", "Oct-26", "Divya", "Stationery Kits", 24000),
    ("K3", "Sep-26", "Sneha", "Stationery Kits", 9000),
    ("K4", "Oct-26", "Priya", "Stationery Kits", 20000),
    ("K5", "Sep-26", "Rahul", "Stationery Kits", 17000),
]

U = 500
tn, ts, te, tw = 294000 // U, 267500 // U, 258000 // U, 244000 // U
caps = (tn, ts, te, tw)
amts = [r[4] // U for r in rows]
assert sum(amts) == tn + ts + te + tw

# DP: state (a0,a1,a2,a3) -> (parent_state, region_choice) via backpointers
states = {(0, 0, 0, 0): None}
for a in amts:
    nxt = dict(states)
    for st in states:
        for r in range(4):
            if st[r] + a > caps[r]:
                continue
            ns = list(st)
            ns[r] += a
            ns = tuple(ns)
            if ns not in nxt:
                nxt[ns] = (st, r)
    states = nxt

final = (tn, ts, te, tw)
if final in states:
    path = []
    st = final
    while states[st] is not None:
        parent, r = states[st]
        path.append(r)
        st = parent
    path.reverse()
    names = ["North", "South", "East", "West"]
    for row, r in zip(rows, path):
        print(names[r], row)
    totals = {nm: 0 for nm in names}
    for row, r in zip(rows, path):
        totals[names[r]] += row[4]
    print(totals)
else:
    print("no region solution")
