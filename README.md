# ARETE – Valutazione Rischio Arboreo

App PWA per la compilazione delle schede di rilevamento ARETE (TRG-P, TRG-S, ORD).

---

## 🚀 Come pubblicare su GitHub Pages (per ottenere l'APK)

### PASSO 1 — Crea un account GitHub
1. Vai su **https://github.com** e clicca **Sign up**
2. Scegli un nome utente (es. `mario-rossi`)
3. Inserisci email e password → verifica l'email

---

### PASSO 2 — Crea il repository
1. Una volta loggato, clicca il **+** in alto a destra → **New repository**
2. Nome repository: `arete` (tutto minuscolo, senza spazi)
3. Lascia **Public** selezionato
4. **NON** spuntare "Add README"
5. Clicca **Create repository**

---

### PASSO 3 — Carica i file
1. Nella pagina del repository appena creato, clicca **uploading an existing file**
2. Trascina TUTTI i file contenuti in questa cartella:
   - `index.html`
   - `app.js`
   - `manifest.json`
   - `sw.js`
   - `icon-192.png`
   - `icon-512.png`
3. Scrivi nel campo "Commit message": `Prima versione ARETE`
4. Clicca **Commit changes**

---

### PASSO 4 — Attiva GitHub Pages
1. Vai su **Settings** (tab in alto nel repository)
2. Nel menu a sinistra clicca **Pages**
3. Sotto "Source" seleziona **Deploy from a branch**
4. Branch: **main** — Folder: **/ (root)**
5. Clicca **Save**
6. Attendi 1-2 minuti, poi aggiorna la pagina
7. Vedrai il link: `https://mario-rossi.github.io/arete/`

---

### PASSO 5 — Installa come app su Android (senza APK)
1. Apri Chrome su Android
2. Vai all'indirizzo `https://tuonome.github.io/arete/`
3. Menu ⋮ → **Aggiungi a schermata Home** → Installa
4. L'app appare come icona nativa sul telefono ✅

---

### PASSO 6 — Genera l'APK con PWA Builder (opzionale)
1. Vai su **https://www.pwabuilder.com**
2. Incolla l'URL: `https://tuonome.github.io/arete/`
3. Clicca **Start** → poi **Package for stores**
4. Scegli **Android** → **Generate Package**
5. Scarica il file `.apk`
6. Trasferiscilo sul telefono e installalo
   (Impostazioni → App → Installa da fonte sconosciuta → Abilita)

---

## Funzionalità
- Schede TRG-P, TRG-S e ORD complete
- Formule automatiche (Circ, Bio, CO₂, O₂, I.A., Impulso)
- Menu a tendina con tutte le voci dal protocollo originale
- Calcolo rischio automatico con tabella colori
- Export PDF fedele al modello originale
- Condivisione PDF via WhatsApp e Mail
- Funziona offline (PWA)
- Utente registrato n° 0368
