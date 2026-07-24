import re, os
p = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Catalogo_de_Livros.py")
linhas = open(p, encoding="utf-8").read().splitlines()
print("TOTAL DE LINHAS:", len(linhas))
print("import json presente?:", any(re.match(r'\s*import\s+json\b', l) for l in linhas))
print("\n--- DEFS de método/função (linha: nome) ---")
for i, l in enumerate(linhas, 1):
    m = re.match(r'\s*def\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(', l)
    if m:
        print(f"{i:5d}: {m.group(1)}")
print("\n--- REATRIBUIÇÕES LivroCatalogApp.X = ... (linha: nome) ---")
for i, l in enumerate(linhas, 1):
    m = re.match(r'\s*LivroCatalogApp\.([A-Za-z_][A-Za-z0-9_]*)\s*=', l)
    if m:
        print(f"{i:5d}: {m.group(1)}")
print("\n--- MARCADORES DE PATCH (linhas com 'PATCH') ---")
for i, l in enumerate(linhas, 1):
    if "PATCH" in l:
        print(f"{i:5d}: {l.strip()}")