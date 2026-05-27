# Luftfracht-Rechner für Outlook

## 1. Kurzbeschreibung

Das **Luftfracht-Rechner**-Add-in berechnet Luftfracht-Volumen und Chargeable Weight direkt in Outlook. Mitarbeiter markieren Packstück-Dimensionen aus E-Mails, die Dimensionen werden automatisch erkannt und verarbeitet. Das Add-in ist über das Microsoft 365 Admin Center für das gesamte Büro in wenigen Minuten ausrollbar und bietet eine zentrale, wartungsfreie Lösung für die tägliche Logistikarbeit.

---

## 2. Voraussetzungen

- **Microsoft 365** (mit Exchange Online)
- **GitHub Account** (kostenlos) für das Hosting der App
- **IT-Admin-Zugang** zum M365 Admin Center
- (Optional) Basis-Kenntnisse in Git/GitHub oder Bereitschaft, Dateien per Browser hochzuladen

---

## 3. Schritt 1: GitHub Pages einrichten (Hosting)

### 3.1 GitHub Repository erstellen

1. **GitHub Account**: Falls noch nicht vorhanden, Account erstellen unter [https://github.com](https://github.com)
2. **Neues Repository erstellen**:
   - Klick auf `+` → `New repository`
   - Repository-Name: `luftfracht-rechner`
   - Wichtig: **Public** auswählen (privat funktioniert nicht mit GitHub Pages)
   - `Create repository` klicken

### 3.2 Dateien hochladen

Folgende Dateien hochladen:
- `index.html`
- `styles.css`
- `app.js`
- `manifest.xml`

**Methode A (Browser, einfachste Methode)**:
- Im Repository-Fenster `Add files` → `Upload files` klicken
- Dateien per Drag & Drop oder `choose your files` auswählen
- `Commit changes` klicken

**Methode B (Git Command Line)**:
```bash
git clone https://github.com/DEIN-USERNAME/luftfracht-rechner.git
cd luftfracht-rechner
# Dateien in das Verzeichnis kopieren
git add .
git commit -m "Initial commit: Luftfracht-Rechner App"
git push origin main
```

### 3.3 GitHub Pages aktivieren

1. Im Repository → `Settings` → `Pages` (linke Seitenleiste)
2. **Source**: "Deploy from a branch" auswählen
3. **Branch**: `main` und `/root` auswählen
4. `Save` klicken
5. Nach 1–2 Minuten ist die App verfügbar unter:
   ```
   https://DEIN-USERNAME.github.io/luftfracht-rechner/
   ```

### 3.4 App testen

1. Browser öffnen und die URL besuchen: `https://DEIN-USERNAME.github.io/luftfracht-rechner/`
2. Testdaten eingeben (z.B. `2x 60*40*30 cm`)
3. "Einlesen" klicken → Ergebnis überprüfen

---

## 4. Schritt 2: manifest.xml anpassen

Das Manifest ist die Konfigurationsdatei für das Outlook-Add-in.

### 4.1 GitHub URL einsetzen

1. `manifest.xml` öffnen (über GitHub UI direkt bearbeiten oder im Texteditor)
2. **PLACEHOLDER durch deinen GitHub-Username ersetzen**:
   - Suche nach: `PLACEHOLDER.github.io`
   - Ersetze mit: `DEIN-USERNAME.github.io`
   - Es gibt ca. 3–4 Vorkommen (Suchfunktion: `Strg+H` in Editor oder GitHub)

3. Beispiel:
   ```xml
   <!-- Vorher -->
   <SourceLocation DefaultValue="https://PLACEHOLDER.github.io/luftfracht-rechner/index.html"/>
   
   <!-- Nachher -->
   <SourceLocation DefaultValue="https://max-mueller.github.io/luftfracht-rechner/index.html"/>
   ```

### 4.2 UUID überprüfen

Die `<Id>` Zeile im Manifest enthält bereits eine generierte eindeutige ID:
```xml
<Id>12345678-1234-1234-1234-123456789012</Id>
```

- Diese ID **beibehalten** (keine Änderung nötig)
- Falls neue ID gewünscht: PowerShell eingeben: `[guid]::NewGuid()` → kopieren und einfügen

### 4.3 Angepasste manifest.xml hochladen

1. Datei auf GitHub hochladen oder Update `Commit` durchführen
2. GitHub Pages deployed die Änderung automatisch innerhalb von Sekunden

---

## 5. Schritt 3: Icons erstellen (Optional, empfohlen)

Das Add-in funktioniert auch ohne eigene Icons (Microsoft-Standard wird verwendet), sieht aber professioneller mit Luftfahrt-bezogenen Icons aus.

### 5.1 Icons beschaffen

Benötigte Größen:
- `icon-16.png` (16×16 Pixel)
- `icon-32.png` (32×32 Pixel)
- `icon-80.png` (80×80 Pixel)

**Einfachste Lösung**: Icon-Download von [https://icons8.com/icons](https://icons8.com/icons)
- Suche nach "airplane" oder "cargo"
- PNG herunterladen
- Im Browser mit kostenlosen Tools (z.B. [https://pixlr.com](https://pixlr.com)) auf die nötigen Größen skalieren

### 5.2 Icons in GitHub hochladen

1. Icons ins Repository hochladen (neuer Ordner `assets/` oder direkt im Root)
2. In `manifest.xml` die Icon-URLs anpassen:
   ```xml
   <bt:Image id="icon16" DefaultValue="https://DEIN-USERNAME.github.io/luftfracht-rechner/icon-16.png"/>
   <bt:Image id="icon32" DefaultValue="https://DEIN-USERNAME.github.io/luftfracht-rechner/icon-32.png"/>
   <bt:Image id="icon80" DefaultValue="https://DEIN-USERNAME.github.io/luftfracht-rechner/icon-80.png"/>
   ```

---

## 6. Schritt 4: Add-in im M365 Admin Center deployen

Dies ist der entscheidende Schritt zum Rollout für alle Mitarbeiter.

### 6.1 Zum Admin Center navigieren

1. Anmelden unter [https://admin.microsoft.com](https://admin.microsoft.com) mit Admin-Konto
2. Linke Seitenleiste → `Einstellungen` (Settings) → `Integrierte Apps`
3. Tab: `Apps hochladen` (Upload apps) auswählen

### 6.2 Manifest hochladen

1. `Office Add-in hochladen` (Upload Office Add-in) klicken
2. Die **angepasste** `manifest.xml` Datei auswählen und hochladen
3. Das System validiert das Manifest automatisch

### 6.3 Zielgruppe und Bereitstellung

1. Im Dialog: **Bereitstellung** wählen:
   - Option A: "Bestimmte Benutzer/Gruppen" (Pilotgruppe)
   - Option B: "Gesamte Organisation" (vollständiger Rollout)
2. Falls Pilotgruppe: Benutzer oder Azure AD Gruppen auswählen
3. `Bereitstellen` (Deploy) klicken

### 6.4 Aktivierungsstatus

Nach dem Upload:
- Grüner Haken = erfolgreich deployed
- **Wichtig**: Es kann bis zu **24 Stunden** dauern, bis das Add-in bei allen Benutzern im Outlook-Client erscheint
- Caching in Outlook kann die Anzeige verzögern (Neustart ggf. erforderlich)

---

## 7. Schritt 5: Testen (vor Office-weitem Rollout)

Vor dem Rollout auf die gesamte Organisation sollte das Add-in lokal getestet werden.

### 7.1 Lokales Testen in Outlook Desktop

1. **Outlook Desktop** öffnen
2. `Datei` → `Add-Ins verwalten` (Manage Add-ins)
3. `Meine Add-Ins` → `Benutzerdefiniertes Add-in hinzufügen` (Add custom add-in)
4. `Aus Datei...` (From a file...) auswählen
5. Die lokale `manifest.xml` Datei auswählen

### 7.2 Im Outlook arbeiten

1. **E-Mail öffnen** oder neu erstellen
2. In der oberen Leiste sollte der Button **"Luftfracht Rechner"** erscheinen
3. Button klicken → Add-in Seitenleiste öffnet sich

### 7.3 Funktionalität prüfen

Testdaten eingeben:
```
2x 60*40*30 cm
120 x 80 x 100cm / 50 kg
```

1. Daten in das Textfeld einfügen
2. `Einlesen` klicken
3. Ergebnis überprüfen:
   - Packstücke sollten erkannt werden
   - Volumengewichte berechnet
   - Chargeable Weight bestimmt

---

## 8. Benutzung für Mitarbeiter (Kurzanleitung)

### 8.1 Add-in öffnen

- **Outlook öffnen** (Web oder Desktop)
- **E-Mail öffnen** oder neu verfassen
- In der Ribbon/Leiste den Button **"Luftfracht Rechner"** anklicken
- Die Seitenleiste öffnet sich rechts im Fenster

### 8.2 Daten verarbeiten

1. **Dimensionen aus der E-Mail markieren** (z.B. "2x 145*90*100 cm")
2. **Markierten Text ins Textfeld kopieren** oder manuell eingeben
3. **"Einlesen" klicken** → Packstücke werden automatisch erkannt
4. **Fehlende Gewichte eintragen** (gelber Hinweis weist auf fehlende Gewichte hin)
5. **"Tabelle kopieren" klicken** → in E-Mail oder Word-Dokument einfügen

### 8.3 Häufige Fragen

- **Das Add-in ist nicht sichtbar?** Outlook neu starten, evtl. bis zu 24h warten
- **Gewichte werden nicht erkannt?** Manuell hinzufügen, dann "Einlesen" erneut klicken
- **Ergebnis sieht falsch aus?** Dimensionen überprüfen (Trennzeichen: *, x, oder Space)

---

## 9. Unterstützte Formate

Das Parser-Modul erkennt verschiedene Schreibweisen. Beispiele:

```
1x 145*90*100 cm
2x 85*38*40 cm
120 x 80 x 156cm / 181 kg  
1x 120x82x70cm
Gw 285 kg  (Gewicht für das vorherige Packstück)
3x 60 x 40 x 30cm @ 15kg each
```

**Formate die erkannt werden:**
- Trennzeichen: `*`, `x`, `X`, oder Leerzeichen
- Gewicht: `/ GEWICHT kg`, `@ GEWICHT kg`, oder `Gw GEWICHT kg`
- Menge: Ziffer + `x` am Anfang (z.B. `2x`, `10x`)
- Einheit wird ignoriert (cm, mm, etc.)

---

## 10. Berechnungsformel

Das Add-in nutzt IATA-Standard:

**Volumengewicht** (VLM):
```
VLM = Länge (cm) × Breite (cm) × Höhe (cm) ÷ 6.000
```

**Chargeable Weight (CW):**
```
CW = Größerer Wert aus:
     - Ist-Gewicht (Actual Weight)
     - Volumengewicht (VLM)
```

Beispiel:
- Packstück: 2 × 60 × 40 × 30 cm, 5 kg
- VLM = 60 × 40 × 30 ÷ 6.000 = 12 kg
- CW = max(5 kg, 12 kg) = **12 kg** (chargeable)

---

## 11. Updates und Wartung

Die Schönheit dieser Lösung: Updates laufen zentral über GitHub Pages.

### 11.1 App aktualisieren

Falls Verbesserungen oder Fehlerbehebungen anstehen:

1. **Dateien auf GitHub aktualisieren**:
   - `app.js` (Logik)
   - `styles.css` (Design)
   - `index.html` (Layout)

2. **GitHub Pages deployed automatisch** innerhalb von Minuten
3. **Kein Update im Admin Center nötig** (solange die manifest.xml URL gleich bleibt)
4. **Alle Benutzer erhalten Update automatisch** beim nächsten Laden des Add-ins

### 11.2 manifest.xml ändern

Falls die `manifest.xml` Datei geändert wird (z.B. neue Features, URL-Änderung):
1. Datei auf GitHub aktualisieren
2. **Neu im Admin Center hochladen** (gleiche Schritte wie 6.2–6.3)
3. Benutzer erhalten Update beim nächsten Neustart von Outlook

---

## Zusammenfassung: Checkliste für IT-Admin

- [ ] GitHub Account (falls nicht vorhanden): https://github.com
- [ ] Repository `luftfracht-rechner` erstellt und auf Public gesetzt
- [ ] Alle 4 Dateien hochgeladen: `index.html`, `styles.css`, `app.js`, `manifest.xml`
- [ ] GitHub Pages aktiviert (Settings → Pages → Deploy from branch)
- [ ] `manifest.xml` angepasst (PLACEHOLDER durch Username ersetzen)
- [ ] GitHub URL getestet: https://DEIN-USERNAME.github.io/luftfracht-rechner/
- [ ] Icons (optional) hochgeladen und in `manifest.xml` aktualisiert
- [ ] Im M365 Admin Center angemeldet
- [ ] `manifest.xml` hochgeladen: Einstellungen → Integrierte Apps → Apps hochladen
- [ ] Bereitstellung gewählt (Pilotgruppe oder ganze Organisation)
- [ ] Lokal getestet (Outlook → Datei → Add-Ins verwalten)
- [ ] Rollout aktiviert ✓

---

## Support & Kontakt

Bei Fragen oder Problemen:
- **Lokal testen** gemäß Schritt 5 (lokale manifest.xml in Outlook)
- **GitHub Pages Logs** überprüfen (Repository → Deployments)
- **Manifest validieren** im M365 Admin Center (Fehler werden angezeigt)
- **Benutzer informieren**: Add-in kann bis zu 24h zum Erscheinen brauchen

---

**Version**: 1.0 | **Datum**: Mai 2026 | **Status**: Produktiv
