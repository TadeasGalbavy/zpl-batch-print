# 🖨️ ZPL Batch Print – Excel → PRN (GUI)

Jednoduchá **Tkinter aplikácia** na dávkovú tlač ZPL/PRN štítkov podľa Excel súboru.

---

## 🚀 Funkcie
- Načíta Excel so zoznamom štítkov a množstiev  
- Spáruje ich so `.prn` súbormi podľa názvu  
- Automaticky vloží správny príkaz `^PQ` podľa množstva  
- Umožní tlač:
  - cez **sieť (IP:9100)**  
  - alebo cez **Windows/USB tlačiareň**  
- Možnosť **Simulácie** (neposiela do tlačiarne, iba loguje)

---

## 📁 Štruktúra projektu
```
zpl-batch-print/
├─ src/
│  └─ zpl_batch_print/
│     ├─ app.py
│     └─ __init__.py
├─ requirements.txt
├─ README.md
└─ .gitignore
```

---

## 📦 Inštalácia

### 1️⃣ Vytvor a aktivuj virtuálne prostredie (odporúčané)
```bash
python -m venv .venv
.\.venv\Scriptsctivate     # Windows
source .venv/bin/activate    # Linux / macOS
```

### 2️⃣ Nainštaluj závislosti
```bash
pip install -r requirements.txt
```

---

## ▶️ Spustenie aplikácie

### A) Priamo cez Python modul
```bash
python -m zpl_batch_print.app
```

### B) (Voliteľne) ako príkaz po inštalácii balíčka
Ak neskôr pridáš `pyproject.toml` s entry pointom, môžeš aplikáciu spustiť príkazom:
```bash
zpl-print
```

---

## 📊 Formát Excel súboru
- **List:** `TLAČ` (predvolený názov)
- **Stĺpce:**
  | sablona | mnozstvo |
  |----------|----------|
  | produkt1.prn | 3 |
  | produkt2.prn | 1 |

> ⚠️ Stĺpce musia byť presne `sablona` a `mnozstvo` (lowercase, bez diakritiky)

---

## 🖨️ Ukážka PRN súboru
Každý `.prn` obsahuje ZPL príkazy, napríklad:
```
^XA
^FO50,40^A0N,30,30^FDHello ZPL^FS
^XZ
```
Aplikácia automaticky vloží alebo aktualizuje príkaz:
```
^PQ5,0,1,N
```
(pred `^XZ`), podľa zadaného množstva.

---

## ⚙️ Voliteľné: Balenie do EXE
Ak chceš mať spustiteľný `.exe`:
```bash
pip install pyinstaller
pyinstaller -F -w -n ZPL_Batch_Print src/zpl_batch_print/app.py
```

---

## 🧩 Poznámky
- `pywin32` je potrebné iba pre USB/Windows režim  
- Pri sieťovej tlači (IP:9100) sa príkazy posielajú priamo cez socket  
- `.prn` súbory sa čítajú s kódovaním `latin-1`

---

## 🧑‍💻 Autor
**Ing. Tadeáš Galbavý**  
E-shop: [www.medical-uniforms.sk](https://www.medical-uniforms.sk)  
GitHub: [tadeasgalbavy](https://github.com/tadeasgalbavy)
