# Tlač štítkov - Label Print Automation

Desktopová aplikácia na hromadnú tlač ZPL štítkov na Zebra tlačiarni.  
Nahrádza manuálny proces - **šetrí 5–7 hodín práce týždenne**.

---

## Čo aplikácia robí

Pred nasadením: operátor musel otvárať každý `.prn` súbor zvlášť, nastavovať množstvo ručne a spúšťať tlač jeden po druhom.

Po nasadení: stačí vyplniť Excel s dvoma stĺpcami, vybrať súbor v aplikácii a kliknúť „Spustiť tlač". Zvyšok ide automaticky.

**Tok:**
1. Operátor vyplní Excel - názov šablóny štítka + počet kusov
2. Aplikácia načíta zoznam, nájde `.prn` súbory na zdieľanom disku
3. Do každého ZPL súboru vloží správne množstvo (`^PQ` príkaz)
4. Odošle tlačové úlohy priamo na Zebra tlačiareň cez Windows spooler

---

## Požiadavky

- Python 3.8+
- Windows (vyžadované pre tlač cez `win32print`)
- Zebra tlačiareň nainštalovaná ako Windows tlačiareň
- Excel súbor `.xlsx` alebo `.xls`

### Knižnice

```bash
pip install pandas openpyxl pywin32
```

> `pywin32` je potrebné len na reálnu tlač. Bez neho funguje aplikácia v **simulačnom režime**.


## Štruktúra Excel súboru

Aplikácia očakáva jednoduchý Excel bez špeciálnych požiadaviek na názov listu.

- **Stĺpec A** - názov `.prn` súboru (iba názov súboru, nie celá cesta)
- **Stĺpec B** - počet kusov (celé číslo)
- Prázdne riadky sú automaticky ignorované

Príklad:
```
nazov_sablony.prn    10
iny_stitok.prn        5
produkt_xyz.prn      25
```

---

## Konfigurácia

V súbore `label_print.py` nastav cestu k zložke so šablónami štítkov:

```python
FIXED_PRN_DIR = r"path_to_DB"   # zmeň na skutočnú cestu, napr. sieťový disk
```

---

## Popis UI

**Simulačný režim** - prechádza celým procesom (načíta Excel, nájde súbory, upraví ZPL), ale nič neposiela na tlačiareň. Ideálne na testovanie pred ostrým spustením.

**imgs** - mrkni zložku imgs pre UI

---

## Čo sa loguje

Každý krok je viditeľný v log okne priamo v aplikácii:

- `OK: nazov.prn × 10` - úspešne odoslaná tlačová úloha
- `SKIP: Nenájdené: ...` - `.prn` súbor neexistuje v zložke
- `CHYBA pri ...` - technická chyba pri konkrétnom súbore
- Na konci: súhrn - spracované / preskočené / celkový počet štítkov

---

## Časté problémy

**„win32print nie je dostupný"**  
→ Spusti `pip install pywin32` a reštartuj aplikáciu.

**„Nenájdené: xyz.prn"**  
→ Skontroluj, že názov v Exceli presne zodpovedá názvu súboru v zložke (vrátane `.prn`).

**Tlačiareň sa nezobrazuje v zozname**  
→ Aplikácia zobrazuje len tlačiarne, ktorých názov obsahuje `zdesigner` (Zebra). Skontroluj názov nainštalovanej tlačiarne vo Windows.

**Nesprávny počet kusov**  
→ Skontroluj, že Stĺpec B obsahuje celé čísla bez textu (nie napr. „10 ks").

---

## Technické poznámky

- ZPL príkaz `^PQ` riadi počet kópií - aplikácia ho automaticky vloží alebo prepíše ak už v súbore existuje.
- Kódovanie `.prn` súborov: `latin-1` (štandard pre staršie ZPL šablóny).
- Tlač prebieha v samostatnom vlákne - UI počas tlače nezamrzne.

---

## Autor

Tadeáš Galbavý - interná automatizácia, Medical Uniforms  
Projekt vznikol ako náhrada manuálneho procesu s úsporou 5–7 hodín/týždeň.
