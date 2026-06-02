# Finance Tracker

Jednoduchý osobní finanční tracker (příjmy, výdaje, úspory) s grafy, kategoriemi
a importem bankovních CSV výpisů (Moneta, Česká spořitelna). Frontend je jediný
soubor `index.html`, data se ukládají do Google tabulky přes Google Apps Script.

```
┌──────────────┐      heslo + data       ┌─────────────────┐     ┌───────────────┐
│  index.html  │  ───────────────────▶   │  Apps Script    │ ──▶ │ Google Sheet  │
│ (prohlížeč)  │  ◀───────────────────   │  (Code.gs)      │     │  (List 1)     │
└──────────────┘                         └─────────────────┘     └───────────────┘
```

---

## Co budeš potřebovat
- Google účet (kvůli tabulce + Apps Scriptu)
- prohlížeč
- (volitelně) GitHub účet, pokud chceš appku hostovat online

---

## Nastavení krok za krokem

### 1) Vytvoř Google tabulku
1. Otevři <https://sheets.google.com> → **Prázdná tabulka**.
2. Spodní list přejmenuj na **`List 1`** (přesně tak, včetně mezery).
3. Do prvního řádku (A1–G1) dej hlavičku:

   | A  | B    | C    | D      | E   | F    | G          |
   |----|------|------|--------|-----|------|------------|
   | id | type | desc | amount | cat | date | created_at |

   *(Hlavička je jen pro tvůj přehled — kód čte podle pozic sloupců.)*

### 2) Nasaď backend (Code.gs)
1. V tabulce: **Rozšíření → Apps Script**.
2. Smaž ukázkový kód a vlož celý obsah souboru **`Code.gs`** z tohoto repa.
3. Najdi řádek `const PASSWORD = 'ZMEN_TOHLE_HESLO';` a nastav **vlastní heslo**.
4. Vpravo nahoře **Nasadit → Nové nasazení**:
   - Typ: **Webová aplikace**
   - Spustit jako: **Já (tvůj účet)**
   - Kdo má přístup: **Kdokoli**
   - **Nasadit** → povol oprávnění → **zkopíruj URL** (končí na `/exec`).

### 3) Propoj frontend
1. V souboru **`index.html`** najdi řádek (cca ř. 472):
   ```js
   const API_URL = 'https://script.google.com/macros/s/.../exec';
   ```
2. Nahraď URL tou svou z kroku 2.

### 4) Spusť appku
- **Rychle:** dvojklik na `index.html` → otevře se v prohlížeči.
- **Online (doporučeno):** nahraj `index.html` na GitHub Pages
  (Settings → Pages → Deploy from branch → `main`).

Při prvním otevření zadej **heslo** z kroku 2. Uloží se v prohlížeči, takže
příště už ho psát nemusíš. Odhlášení je v ⚙ (nastavení).

---

## Import bankovního výpisu
Tlačítko **⬆ Import výpisu** → nahraj CSV. Podporováno:
- **Moneta** (běžný účet i platby kartou)
- **Česká spořitelna** (běžný účet i kreditní karta)

Aplikace transakce automaticky roztřídí na příjem/výdaj/úspory a navrhne
kategorii; před uložením můžeš vše upravit. Přesuny mezi vlastními účty se
vyfiltrují (šedě, nezaškrtnuté).

---

## Bezpečnost
- Heslo žije **jen v `Code.gs`** (na serveru), ne v `index.html`. Server odmítne
  jakýkoli požadavek bez správného hesla — tabulka je tak chráněná i proti
  přímému volání API.
- Zvol si **unikátní heslo**, které nepoužíváš jinde.
- Když změníš kód v `Code.gs`, je potřeba **znovu nasadit**: Nasadit → Spravovat
  nasazení → ✏️ → Verze: **Nová verze** → Nasadit (URL zůstane stejná).

---

## Soubory v repu
| Soubor       | K čemu je                                                  |
|--------------|------------------------------------------------------------|
| `index.html` | Celá aplikace (frontend). Stačí upravit `API_URL`.         |
| `Code.gs`    | Backend do Google Apps Scriptu. Nastav `PASSWORD`.         |
| `README.md`  | Tento návod.                                               |
