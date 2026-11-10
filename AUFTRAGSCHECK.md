# Auftragscheck (V5) – Microsoft Forms

**Kurzfassung:** Ein kurzes Formular pro neuem Auftrag. 4 Abschnitte. Nur intern, **keine Kunden-Mail** am Ende. Antworten landen automatisch in **Excel**.

---

## Aufbau in Forms

Forms → **Neues Formular** → Titel: **Auftragscheck Druckvorstufe**

### Abschnitt 1: Bearbeiter + Status

| Frage | Typ | Optionen | Pflicht |
|---|---|---|---|
| **Auftrag / Betreff** | Text | – | Ja |
| **Bearbeiter** | Auswahl | Florian Stauber, Simone Ruppli, Thomas Heuberger | Ja |
| **Status** | Auswahl | zu erledigen, in Bearbeitung :), GzD erstellen, Warten GzD, Warten Kunde, Warten LT, Info | Ja |
| **Priorität?** | Auswahl | Ja, Nein | Ja |

### Abschnitt 2: Anhang

| Frage | Typ | Optionen | Pflicht |
|---|---|---|---|
| **Druckdaten im Anhang?** | Auswahl | Ja, Nein | Ja |

**Verzweigung** (… → Verzweigung hinzufügen):

- **Ja** → weiter zu Abschnitt 3
- **Nein** → zu Abschnitt "Daten fehlen"

### Abschnitt "Daten fehlen"

Nur ein Hinweistext:

> **Standard-Mail "Bitte Daten senden" verschicken.**
> Status in Outlook auf **Warten Kunde** setzen.

Danach → **Ende des Formulars**.

*Die Standard-Mail ist eine Outlook-Vorlage. Sie wird nicht vom Formular erzeugt.*

### Abschnitt 3: Details

Nur zwei Felder:

| Frage | Typ | Pflicht |
|---|---|---|
| **Was?** | Text (kurz) | Ja |
| **Bis wann?** | Datum | Ja |

### Abschnitt 4: Technik (alles Dropdown)

| Frage | Optionen |
|---|---|
| **Farbraum** | RGB, CMYK, gemischt, ungeprüft |
| **Preflight gemacht?** | Ja, Nein |

**Farbraum** steht nicht in der Mail. Er kommt aus dem **PDF** (Acrobat **Preflight**). Noch nicht geprüft → **ungeprüft**.

*Weitere Dropdowns nach Bedarf ergänzen (z. B. Format, Beschnitt). Bitte im Team festlegen.*

---

## Einstellungen

- **Wer kann antworten:** Nur Personen in der Organisation
- **Namen erfassen:** an
- **Antworten in Excel öffnen:** an (Registerkarte Antworten)

---

## Ablauf im Alltag

1. Neue Mail kommt rein.
2. **Formular ausfüllen.**
3. In Outlook **Name + Status** als Kategorie setzen (gleich wie im Formular).
4. Fehlen Daten → Standard-Mail senden, Status **Warten Kunde**.
