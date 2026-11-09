/* Auftrags-Parser für die Druckvorstufe
 * Liest den Text einer E-Mail und sucht die wichtigen Auftragsangaben heraus.
 * Alles läuft lokal im Browser/Outlook, es wird nichts an einen Server geschickt.
 */
(function (root) {
  'use strict';

  // ---------- Felder, die geprüft werden ----------
  var FIELDS = [
    { key: 'format',    label: 'Format',             frage: 'In welchem Format sollen die Ausdrucke erstellt werden (z. B. A3, A2 oder Maße in mm/cm)?' },
    { key: 'menge',     label: 'Menge',              frage: 'Wie viele Exemplare benötigen Sie?' },
    { key: 'papier',    label: 'Papier',             frage: 'Auf welchem Papier bzw. in welcher Grammatur soll gedruckt werden?' },
    { key: 'farbe',     label: 'Farbe / Seiten',     frage: 'Soll farbig oder schwarz-weiß gedruckt werden, einseitig oder beidseitig?' },
    { key: 'termin',    label: 'Termin',             frage: 'Bis wann (Datum und Uhrzeit) benötigen Sie die Ausdrucke?' },
    { key: 'daten',     label: 'Druckdaten',         frage: 'Bitte senden Sie uns die Druckdaten (bevorzugt als PDF) zu.' },
    { key: 'lieferung', label: 'Lieferung / Abholung', frage: 'Sollen die Ausdrucke abgeholt oder geliefert werden? Bei Lieferung bitte die Adresse angeben.' },
    { key: 'weiter',    label: 'Weiterverarbeitung', frage: 'Gibt es Wünsche zur Weiterverarbeitung (schneiden, falzen, laminieren, binden ...)?' },
    { key: 'auftrag',   label: 'Auftrags-/Projekt-Nr.', frage: 'Unter welcher Auftrags- bzw. Projektnummer läuft das Ganze?' }
  ];

  // Diese Felder sind nur Zusatzinfo und gelten nie als "fehlt".
  var OPTIONAL = { weiter: true, auftrag: true };

  // ---------- Pflichtfelder je Auftragsart ----------
  // Gibt die Schlüssel der Felder zurück, die für DIESE Mail Pflicht sind.
  // Alles, was hier fehlt, wird in der Checkliste als "fehlt" markiert.
  // Regeln (anpassbar):
  //  - Standard: Format, Menge, Papier, Farbe, Termin, Daten, Lieferung sind Pflicht.
  //  - Verweis auf früheren Auftrag ("wie letztes Mal", "Alt-Auftrag"): Format, Papier und Farbe
  //    stehen im Alt-Auftrag und werden nicht angemahnt. Menge, Termin, Daten, Lieferung bleiben.
  //  - Nur Großformat (Plakat, Banner, Roll-up, Plot, LFP, Leinwand): Farbe ist nicht Pflicht,
  //    weil sie dort fast immer 4-farbig ist.
  //  - Nur Visitenkarten: Format ist nicht Pflicht (Standardgrösse).
  //  Mischaufträge (z. B. Plakat + Flyer wie auf dem Laufzettel) behalten alle Pflichtfelder.
  function pflichtfelder(jobText) {
    var t = String(jobText || '');
    var pflicht = FIELDS.filter(function (f) { return !OPTIONAL[f.key]; })
                        .map(function (f) { return f.key; });
    function ohne(keys) {
      pflicht = pflicht.filter(function (k) { return keys.indexOf(k) === -1; });
    }

    var gross = /(?<![A-Za-zÄÖÜäöüß])(Plakate?|Poster|Banner|Roll-?ups?|LFP|Gro(?:ß|ss)format|Plot(?:s|ter)?|Leinwand|Fotoposter)(?![A-Za-zÄÖÜäöüß])/i.test(t);
    var klein = /(?<![A-Za-zÄÖÜäöüß])(Laserdruck|Flyer|Briefe?|Brosch(?:ü|ue)ren?|Karten|Folder|Prospekte?|Kopien|Etiketten?|Kataloge?|Mappen?|Kalender)(?![A-Za-zÄÖÜäöüß])/i.test(t);
    var visit = /Visitenkarten?/i.test(t);

    if (gross && !klein && !visit) ohne(['farbe']);
    if (visit && !gross) ohne(['format']);
    if (/(?<![A-Za-zÄÖÜäöüß])(wie\s+(?:letztes\s+Mal|beim\s+letzten\s+Mal|gehabt|bisher|bei\s+Auftrag|Auftrag)|analog\s+(?:zu|Auftrag)|gleich\s+wie|wie\s+Alt-?Auftrag)/i.test(t)) {
      ohne(['format', 'papier', 'farbe']);
    }
    return pflicht;
  }

  // ---------- Hilfsfunktionen ----------
  var WORD_L = '(?<![A-Za-zÄÖÜäöüß0-9_])';
  var WORD_R = '(?![A-Za-zÄÖÜäöüß0-9_])';

  function wb(pattern, flags) {
    return new RegExp(WORD_L + '(?:' + pattern + ')' + WORD_R, flags || 'gi');
  }

  function context(text, index, length) {
    var from = Math.max(0, index - 45);
    var to = Math.min(text.length, index + length + 45);
    var snip = text.slice(from, to).replace(/\s+/g, ' ').trim();
    return (from > 0 ? '… ' : '') + snip + (to < text.length ? ' …' : '');
  }

  function collect(text, regexes) {
    var out = [];
    var seen = {};
    regexes.forEach(function (re) {
      re.lastIndex = 0;
      var m;
      while ((m = re.exec(text)) !== null) {
        var val = m[0].replace(/\s+/g, ' ').trim();
        var k = val.toLowerCase();
        if (!seen[k]) {
          seen[k] = true;
          out.push({ text: val, context: context(text, m.index, m[0].length), index: m.index });
        }
        if (m.index === re.lastIndex) re.lastIndex++;
      }
    });
    out.sort(function (a, b) { return a.index - b.index; });
    return out;
  }

  var MENGE_EINHEIT = 'Stück|Stk\\.?|St\\.|Exemplare?|Ex\\.?|Expl\\.?|Ausdrucke?|Drucke?|Plakate?|Poster|Banner|Flyer|Karten|Visitenkarten|Postkarten|Aufkleber|Schilder|Roll-?ups?|Broschüren|Mappen|Sets?|Bögen|Blatt|Kopien|Pläne|Plots?|Folien|fach';

  // ---------- Muster je Feld ----------
  var PATTERNS = {
    format: function () {
      return [
        wb('(?:DIN[\\s-]*)?(?:SRA|RA|A)\\s?[0-8](?:\\+)?', 'g'),
        wb('DIN[\\s-]*lang|DIN[\\s-]*lang', 'gi'),
        wb('B[0-4]|C[4-6]', 'g'),
        /(?<![\d.,])\d{2,4}(?:[.,]\d+)?\s?(?:mm|cm|m)?\s?[x×*]\s?\d{2,4}(?:[.,]\d+)?\s?(?:mm|cm|m)(?![A-Za-z])/gi,
        wb('Visitenkarten?|Postkarten?|Roll-?up|Quadratisch(?:es)?|Hochformat|Querformat|Poster|Plan|Pl[äa]ne', 'gi')
      ];
    },
    menge: function () {
      return [
        /(?:Menge|Auflage|Anzahl|Stückzahl|Stueckzahl)\s*[:=]?\s*\d[\d.']*/gi,
        new RegExp("(?<![A-Za-zÄÖÜäöüß\\d.,'\\-\\/])\\d{1,3}(?:[.']\\d{3})+\\s?(?:" + MENGE_EINHEIT + ")" + WORD_R, 'gi'),
        new RegExp("(?<![A-Za-zÄÖÜäöüß\\d.,'\\-\\/])\\d+\\s?(?:" + MENGE_EINHEIT + ")" + WORD_R, 'gi'),
        /(?<![\d.,])\d{1,5}\s?[x×](?!\s?\d)/gi,
        wb('einmal|zweimal|dreimal|viermal|fünfmal|einfach|doppelt|je\\s\\d+', 'g')
      ];
    },
    papier: function () {
      return [
        /(?<![\d.,])\d{2,3}\s?(?:g\/m²|g\/m2|g\/qm|g\/sm|gm²|gm2|gsm|gr\b|g(?![A-Za-z])|Gramm)/gi,
        wb('Image Impact|DCP|Maranello|Bilderdruck|Offsetpapier|Offset|Naturpapier|Recyclingpapier|Recycling|Karton|Pappe|Plakatpapier|Fotopapier|Photopapier|Fotoglanz|Glanzpapier|Mattpapier|gestrichen(?:es)?|ungestrichen(?:es)?|Kopierpapier|Bond|Tyvek|Banner(?:stoff)?|Mesh|Klebefolie|Folie|Leinwand|Canvas|Dibond|Forex|Kapa|Synthetik(?:papier)?|Blueback|Whiteback|Vlies|Textil|Transparentpapier|Pergament|Zeichenpapier|Seidenmatt|Hochglanz|Glossy|Satin|Papier|Material', 'gi'),
        wb('matt(?:es|e)?|glänzend|glanz', 'g')
      ];
    },
    farbe: function () {
      return [
        /(?<![\d./])(?:4\/[014]|1\/[01]|5\/[05]|6\/[06]|0\/[14])(?![\d/])/g,
        wb('CMYK|RGB|Pantone|PMS|HKS|Sonderfarbe|Schmuckfarbe|Euroskala|s\\/w|schwarz[\\s-]?wei(?:ß|ss)|Graustufen|einfarbig|vierfarbig|4c|4-farbig|farbig|in Farbe|Farbe|Farbdruck|beidseitig|einseitig|doppelseitig|einseitig bedruckt|beidseitig bedruckt|4\\+0|4\\+4|1\\+0|1\\+1', 'gi')
      ];
    },
    termin: function () {
      return [
        /(?<![\d.])(?:0?[1-9]|[12]\d|3[01])\.\s?(?:0?[1-9]|1[0-2])\.(?:\s?(?:20)?\d{2})?(?![\d])/g,
        wb('(?:bis|spätestens|am|zum|für|vor)\\s+(?:zum\\s+|kommenden\\s+|nächsten\\s+|diesen\\s+)?(?:Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag)', 'gi'),
        wb('Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag', 'g'),
        wb('heute|morgen|übermorgen|nächste Woche|diese Woche|Ende der Woche|Ende des Monats|(?:bis|vor|am|zum|übers)\\s+Wochenende', 'g'),
        wb('KW\\s?\\d{1,2}|Kalenderwoche\\s?\\d{1,2}', 'gi'),
        wb('(?:bis\\s+)?\\d{1,2}(?:[:.]\\d{2})?\\s?Uhr', 'gi'),
        wb('Liefertermin|Abgabetermin|Abholtermin|Deadline|Fertigstellung|ASAP|asap|sofort|dringend|eilig|Express|schnellstmöglich|so schnell wie möglich', 'gi')
      ];
    },
    lieferung: function () {
      return [
        wb('Postaufgabe|Porto|B-Post|A-Post|Frankier\\w*|verpacken|Versandarbeiten|Paketinhalt|Schachteln?|Gummiband|Lieferung|liefern|geliefert|Lieferadresse|Lieferanschrift|Versand|versenden|versendet|Abholung|abholen|abgeholt|Selbstabholer|Zustellung|Kurier|Spedition|DHL|UPS|Paket|vorbeibringen|bringen|Auslieferung|Lieferschein|frei Haus', 'gi'),
        /(?<![\d])\d{5}\s+[A-ZÄÖÜ][a-zäöüß]+/g
      ];
    },
    weiter: function () {
      return [
        wb('Nutzen|Ntz\\.?|gefalzt|geschnitten|laminiert|gebunden|gelocht|gestanzt|aufgezogen|kaschiert|geklebt|Falzung|zuschneiden|schneiden|Schnitt|Beschnitt|Anschnitt|Konturschnitt|falzen|Falz|Wickelfalz|Zickzackfalz|laminieren|Laminierung|kaschieren|binden|Bindung|Klammerheftung|Heftung|Spiralbindung|Spiral|Ringbuch|Ösen|Hohlsaum|kleben|aufziehen|aufkaschieren|stanzen|lochen|Lochung|perforieren|rillen|Rillung|Veredelung|UV-Lack|Lackierung|Folierung|einschweißen|konfektionieren|Konfektion|Nutzen', 'gi')
      ];
    },
    auftrag: function () {
      return [
        /(?:Auftrags?(?:nummer|-?Nr\.?)?|Projekt(?:nummer|-?Nr\.?)?|Bestell(?:ung|nummer|-?Nr\.?)?|Angebot(?:s-?Nr\.?)?|Vorgang(?:s-?Nr\.?)?|Kommission)\s*[:#-]?\s*[A-Z0-9][A-Z0-9\-\/]{3,}/gi
      ];
    }
  };

  var DATEN_FOLGEN = wb('Daten (?:folgen|kommen|schicke ich nach|reiche ich nach|werden nachgereicht|nachgereicht|später)|Datei(?:en)? (?:folgen|kommen|schicke ich nach|werden nachgereicht|später)|folgen (?:noch|nach)|sende (?:ich )?nach|nachliefern|nachreichen|noch nicht fertig', 'gi');
  var DATEN_TEXT = [
    wb('Druckdaten|Druckdatei|Druck-PDF|PDF|im Anhang|anbei|beigefügt|angehängt|angefügt|Anhang|WeTransfer|Download|Dropbox|Cloud|OneDrive|SharePoint|Laufwerk|Server|FTP|Link', 'gi'),
    /https?:\/\/\S+/gi
  ];
  var DATEI_ENDUNGEN = /\.(pdf|ai|indd|idml|psd|tif|tiff|jpg|jpeg|png|eps|svg|zip|rar|7z|dwg|dxf|docx?|pptx?|xlsx?|cdr|afdesign|afpub|afphoto)$/i;
  var DATEI_IGNORE = /^(image\d*|outlook-[\w-]+|~?WRD\d+|logo[\w-]*)\.(png|jpg|jpeg|gif)$/i;

  function DATEN_VORHANDEN_TREFFER(t) {
    return collect(t, [wb('(?:druckfertige[sn]?\\s+)?PDF\\s+geliefert|Daten\\s+(?:sind\\s+)?(?:geliefert|vorhanden)|Daten\\s+liegen\\s+(?:vor|bereit|auf\\s+dem\\s+Server)|Daten\\s+auf\\s+dem\\s+Server|druckfertige[sn]?\\s+(?:PDF|Daten)|Daten\\s+im\\s+Auftrag', 'gi')]);
  }

  var REFERENZ = wb('wie\\s+(?:letztes\\s+Mal|beim\\s+letzten\\s+Mal|gehabt|bisher|immer|bei\\s+Auftrag|Auftrag)|analog\\s+(?:zu|Auftrag)|gleich\\s+wie|wie\\s+Alt-?Auftrag|Alt-?Auftrag(?:s-?Nr\\.?)?', 'gi');

  var URGENT = wb('dringend|eilig|ASAP|asap|sofort|schnellstmöglich|so schnell wie möglich|Express|heute noch|rush|wichtig', 'gi');

  // ---------- Hauptfunktion ----------
  function analyse(input) {
    input = input || {};
    var text = String(input.text || '');
    var subject = String(input.subject || '');
    var attachments = input.attachments || [];
    var full = (subject ? subject + '\n' : '') + text;
    full = full.slice(0, 60000);
    // Kopfzeilen weitergeleiteter Mails (Gesendet, An, Cc ...) sind keine Auftragsangaben
    full = full.replace(/^[ \t>]*(?:Gesendet|Sent|Von|From|An|To|Cc|Bcc|Datum|Date)\s*:.*$/gmi, '');

    var result = {};
    FIELDS.forEach(function (f) {
      var maker = PATTERNS[f.key];
      result[f.key] = {
        key: f.key,
        label: f.label,
        frage: f.frage,
        optional: !!OPTIONAL[f.key],
        values: maker ? collect(full, maker()) : [],
        state: 'fehlt'
      };
    });

    // Besonderheiten Format: Plain "Plan/Poster" allein reicht nicht
    // -> nur zählen, wenn auch A-Format oder Maße da sind; sonst als schwacher Treffer markieren
    var fmt = result.format;
    var starkF = fmt.values.filter(function (v) {
      return !/^(Visitenkarten?|Postkarten?|Roll-?up|Quadratisch(?:es)?|Hochformat|Querformat|Poster|Plan|Pl[äa]ne)$/i.test(v.text);
    });
    if (starkF.length) {
      fmt.values = starkF;
    } else if (fmt.values.length) {
      fmt.weak = true; // nur ein Hinweis, kein echtes Format
    }

    // Papier: "Papier"/"Material" allein ist nur ein Hinweis
    var pap = result.papier;
    var starkP = pap.values.filter(function (v) { return !/^(Papier|Material)$/i.test(v.text); });
    if (starkP.length) { pap.values = starkP; } else if (pap.values.length) { pap.weak = true; }

    // Farbe: "Farbe" allein = schwach
    var far = result.farbe;
    var starkFa = far.values.filter(function (v) { return !/^Farbe$/i.test(v.text); });
    if (starkFa.length) { far.values = starkFa; } else if (far.values.length) { far.weak = true; }

    // Termin: nur Uhrzeit / nur Wochentag ohne "bis" ist schwach
    var ter = result.termin;
    var datum = ter.values.filter(function (v) {
      return /\d{1,2}\.\s?\d{1,2}\./.test(v.text) || /^KW|Kalenderwoche/i.test(v.text) ||
             /^(bis|spätestens|am|zum|für|vor)\s/i.test(v.text) ||
             /^(heute|morgen|übermorgen|nächste Woche|diese Woche|Ende der Woche|Ende des Monats|Wochenende)$/.test(v.text) ||
             /(Liefertermin|Abgabetermin|Abholtermin|Deadline|Fertigstellung)/i.test(v.text);
    });
    ter.hasDatum = datum.length > 0;
    ter.hasUhrzeit = ter.values.some(function (v) { return /Uhr/i.test(v.text); });
    ter.eilig = ter.values.some(function (v) { return /^(ASAP|sofort|dringend|eilig|Express|schnellstmöglich|so schnell wie möglich)$/i.test(v.text); });
    // Wochentage allein (ohne bis/am) herausfiltern wenn schon ein Datum da ist
    if (datum.length) {
      ter.values = ter.values.filter(function (v) {
        if (/^(Montag|Dienstag|Mittwoch|Donnerstag|Freitag|Samstag|Sonntag)$/.test(v.text)) {
          // nur behalten, wenn nicht schon im längeren Treffer enthalten
          return !ter.values.some(function (w) { return w !== v && w.text.indexOf(v.text) !== -1 && w.text.length > v.text.length; });
        }
        return true;
      });
    } else if (ter.values.length) {
      ter.weak = true; // z. B. nur "dringend" oder nur "15 Uhr" ohne Tag
    }

    // Daten
    var dat = result.daten;
    var dateiAnhang = attachments
      .map(function (a) { return typeof a === 'string' ? a : (a && (a.name || a.fileName)) || ''; })
      .filter(function (n) { return n && DATEI_ENDUNGEN.test(n) && !DATEI_IGNORE.test(n); });
    dat.values = dateiAnhang.map(function (n) { return { text: n, context: 'Anhang: ' + n, index: -1 }; });
    var folgen = collect(full, [DATEN_FOLGEN]);
    var textHinweise = collect(full, DATEN_TEXT).filter(function (v) {
      return !/^(Anhang|Server|Cloud|Link|Download|PDF)$/i.test(v.text) || /^(PDF)$/i.test(v.text);
    });
    var hatLink = textHinweise.some(function (v) { return /^https?:/i.test(v.text) || /WeTransfer|Dropbox|OneDrive|SharePoint|FTP|Laufwerk/i.test(v.text); });
    if (dateiAnhang.length || hatLink) {
      dat.state = 'ok';
      textHinweise.filter(function (v) { return /^https?:/i.test(v.text); }).forEach(function (v) { dat.values.push(v); });
    } else if (folgen.length) {
      dat.values = folgen;
      dat.folgt = true;
    } else if (DATEN_VORHANDEN_TREFFER(full).length) {
      // z. B. "Daten: druckfertiges PDF geliefert" (liegt schon auf dem Server)
      dat.values = DATEN_VORHANDEN_TREFFER(full);
      dat.state = 'ok';
      dat.vorhanden = true;
    } else if (textHinweise.some(function (v) { return /anbei|beigefügt|angehängt|angefügt|im Anhang/i.test(v.text); })) {
      // Text sagt "anbei", aber es ist keine Datei erkannt (z. B. Bild im Text oder Anhang nicht lesbar)
      dat.values = textHinweise.filter(function (v) { return /anbei|beigefügt|angehängt|angefügt|im Anhang/i.test(v.text); });
      dat.weak = true;
    } else {
      dat.values = [];
    }
    dat.anhaenge = dateiAnhang.length;

    // Lieferung: "Post" etc. nicht nötig; alles Treffer sind okay
    // Zustand bestimmen
    var pflicht = pflichtfelder(full);
    FIELDS.forEach(function (f) {
      var r = result[f.key];
      var has = r.values.length > 0;
      if (f.key === 'daten') {
        r.state = r.state === 'ok' ? 'ok' : (r.folgt ? 'folgt' : (r.weak ? 'unklar' : 'fehlt'));
      } else if (f.key === 'termin') {
        r.state = r.hasDatum ? 'ok' : (has ? 'unklar' : 'fehlt');
      } else if (has && !r.weak) {
        r.state = 'ok';
      } else if (has && r.weak) {
        r.state = 'unklar';
      } else {
        r.state = 'fehlt';
      }
      r.required = pflicht.indexOf(f.key) !== -1;
    });

    // ---------- Verlauf, Weiterleitung, fehlende Anhänge ----------
    var thread = analyseVerlauf(text, subject, attachments, input.from);
    var risiken = [];
    var attNames = thread.attachmentNames;
    if (thread.missingFiles.length) {
      risiken.push({
        typ: 'anhang_fehlt',
        titel: attNames.length
          ? 'Anhang unvollständig: ' + thread.missingFiles.length + ' von ' + thread.mentionedFiles.length + ' genannten Dateien fehlen'
          : 'Anhang nicht vorhanden: Die Mail nennt ' + thread.mentionedFiles.length + ' Datei' + (thread.mentionedFiles.length === 1 ? '' : 'en') + ', es hängt aber nichts an',
        details: thread.missingFiles,
        weitergeleitet: thread.forwarded
      });
    } else if (!attNames.length && thread.erwaehntAnhang) {
      risiken.push({
        typ: 'anhang_fehlt',
        titel: 'Anhang nicht vorhanden: Die Mail erwähnt Daten im Anhang, es hängt aber nichts an',
        details: [],
        weitergeleitet: thread.forwarded
      });
    } else if (thread.forwarded && !attNames.length) {
      risiken.push({
        typ: 'weiterleitung_ohne_anhang',
        titel: 'Weitergeleitet ohne Anhang: Original-Mail prüfen, ob dort Dateien dranhingen',
        details: [],
        weitergeleitet: true
      });
    }
    var ref = collect(full, [REFERENZ]);
    if (ref.length) {
      risiken.push({
        typ: 'referenz',
        titel: 'Verweis auf früheren Auftrag ("' + ref[0].text + '"): Alt-Auftrag öffnen und Unterschiede klären',
        details: [],
        weitergeleitet: false
      });
    }
    if (risiken.length && risiken[0].typ === 'anhang_fehlt') {
      result.daten.state = attNames.length ? 'unklar' : 'fehlt';
      result.daten.risiko = true;
      if (!attNames.length) { result.daten.values = []; result.daten.folgt = false; }
    }

    var missing = FIELDS.filter(function (f) {
      var r = result[f.key];
      return r.required && r.state !== 'ok';
    }).map(function (f) { return result[f.key]; });

    // Wichtige Punkte: Zeilen/Sätze mit den meisten Treffern
    var punkte = wichtigePunkte(full, result);

    // Dringlichkeit
    var urgent = collect(full, [URGENT]).length > 0 || result.termin.eilig;

    // Status-Vorschlag
    var vorschlag = null;
    if (result.daten.folgt) vorschlag = 'Daten folgen';
    else if (/(warte|Rückmeldung|Freigabe|Korrektur)/i.test(full) && missing.length) vorschlag = 'Warten';


    // ---------- Angaben für die Polygrafen + Sonderwünsche ----------
    var polygraf = bauePolygraf(full, result);
    var sonder = baueSonder(full, result, input.now ? new Date(input.now) : new Date());

    var out = {
      polygraf: polygraf,
      sonder: sonder,
      fields: result,
      order: FIELDS.map(function (f) { return f.key; }),
      missing: missing,
      punkte: punkte,
      urgent: urgent,
      vorschlag: vorschlag,
      risiken: risiken,
      thread: thread,
      positionen: analysePositionen(full)
    };
    out.kern = baueKern(out, full);
    return out;
  }

  // ---------- Polygrafen: Seitenumfang, wie gedruckt, Farbprofil, Papier ----------
  function bauePolygraf(full, result) {
    var texte = function (list) { return list.map(function (v) { return v.text; }); };
    var seiten = collect(full, [
      /(?:Seitenumfang|Umfang)\s*[:=]?\s*\d{1,4}(?:\s?(?:Seiten|S\.?))?/gi,
      /(?<![\d.,])\d{1,4}\s?(?:Seiten|Seiter|S\.)(?![A-Za-zÄÖÜäöüß0-9_])/gi,
      /(?<![\d.,])\d{1,4}[\s-]?seitig(?![A-Za-zÄÖÜäöüß0-9_])/gi
    ]);
    var druck = collect(full, [
      /(?<![\d./])(?:4\/[014]|1\/[01]|5\/[05]|6\/[06]|0\/[14]|4\+[04]|1\+[01])(?![\d/])/g,
      wb('einseitig(?:\\s+bedruckt)?|beidseitig(?:\\s+bedruckt)?|doppelseitig|Duplex|Simplex|Vorder-?\\s?und\\s+Rückseite|Schön-?\\s?und\\s+Widerdruck|s\\/w|schwarz[\\s-]?wei(?:ß|ss)|Graustufen|einfarbig|vierfarbig|4-?farbig|farbig|in Farbe|Farbdruck', 'gi')
    ]);
    var profil = collect(full, [
      wb('ISO[\\s-]?Coated(?:[\\s-]?v2)?(?:[\\s-]?300)?|PSO[\\s-]?(?:Coated|Uncoated|LWC)[\\w\\s-]{0,12}?|FOGRA\\s?\\d{2}[a-z]?|ECI[\\s-]?RGB(?:\\s?v2)?|sRGB|Adobe[\\s-]?RGB|ICC[\\s-]?Profile?n?|Farbprofile?n?|Euroskala|Pantone|HKS|CMYK|RGB', 'gi')
    ]);
    return {
      seiten: texte(seiten),
      druck: texte(druck),
      profil: texte(profil),
      papier: texte(result.papier.values)
    };
  }

  // ---------- Sonderwünsche: kein DIN-Format, dringend, Zeitangaben für heute/morgen ----------
  var DIN_MM = [[841,1189],[594,841],[420,594],[297,420],[210,297],[148,210],[105,148],[74,105],
                [1000,1414],[707,1000],[500,707],[353,500],[250,353],[176,250],
                [1030,1456],[728,1030],[515,728],[364,515],[257,364],[182,257],
                [99,210],[110,220],[324,458],[229,324],[162,229],[114,162],[85,55],[55,85]];
  function mm(zahl, einheit) {
    var n = parseFloat(String(zahl).replace(',', '.'));
    einheit = (einheit || '').toLowerCase();
    if (einheit === 'cm') return n * 10;
    if (einheit === 'm') return n * 1000;
    return n;
  }
  function istDin(a, b) {
    return DIN_MM.some(function (d) {
      return (Math.abs(d[0] - a) <= 1 && Math.abs(d[1] - b) <= 1) || (Math.abs(d[0] - b) <= 1 && Math.abs(d[1] - a) <= 1);
    });
  }

  function baueSonder(full, result, jetzt) {
    var out = [];
    var seen = {};
    function add(typ, text, hinweis) {
      var k = typ + '|' + text.toLowerCase();
      if (seen[k]) return;
      seen[k] = true;
      out.push({ typ: typ, text: text, hinweis: hinweis || '' });
    }

    // 1) Kein DIN-Format (Maße, die keine Normgröße sind, oder ausdrücklich "Sonderformat")
    var re = /(?<![\d.,])(\d{2,4}(?:[.,]\d+)?)\s?(mm|cm|m)?\s?[x×*]\s?(\d{2,4}(?:[.,]\d+)?)\s?(mm|cm|m)(?![A-Za-z])/gi;
    var m;
    while ((m = re.exec(full)) !== null) {
      var einheit = m[4];
      var a = mm(m[1], m[2] || einheit), b = mm(m[3], einheit);
      if (!istDin(a, b)) add('format', m[0].replace(/\s+/g, ' ').trim(), 'kein DIN-Format');
    }
    collect(full, [wb('Sonderformat|Sonderma(?:ß|ss)e?|Sondergr(?:ö|oe)(?:ß|ss)e|freies Format|kein(?:e)? DIN(?:[\\s-]*Format)?|nicht DIN|individuelles Format|Wunschformat', 'gi')])
      .forEach(function (v) { add('format', v.text, 'kein DIN-Format'); });

    // 2) Dringend
    collect(full, [wb('dringend|eilig|ASAP|asap|sofort|schnellstmöglich|so schnell wie möglich|Express|heute noch|rush|Eilauftrag|eilt', 'gi')])
      .forEach(function (v) { add('dringend', v.text, 'dringend'); });

    // 3) Zeitangaben, die HEUTE oder MORGEN gelten
    var p2 = function (x) { return ('0' + x).slice(-2); };
    var morgen = new Date(jetzt.getFullYear(), jetzt.getMonth(), jetzt.getDate() + 1);
    var heuteD = p2(jetzt.getDate()) + '.' + p2(jetzt.getMonth() + 1) + '.';
    var morgenD = p2(morgen.getDate()) + '.' + p2(morgen.getMonth() + 1) + '.';
    var tage = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];
    var wtHeute = tage[jetzt.getDay()], wtMorgen = tage[morgen.getDay()];

    collect(full, [wb('(?:bis\\s+|am\\s+|für\\s+|vor\\s+)?heute(?:\\s+noch)?(?:\\s+(?:bis\\s+)?\\d{1,2}(?:[:.]\\d{2})?\\s?Uhr)?|(?:bis\\s+|am\\s+|für\\s+|vor\\s+)?morgen(?:\\s+(?:früh|Vormittag|Mittag|Nachmittag|Abend))?(?:\\s+(?:bis\\s+)?\\d{1,2}(?:[:.]\\d{2})?\\s?Uhr)?', 'gi')])
      .forEach(function (v) {
        var low = v.text.toLowerCase();
        add('zeit', v.text, /morgen/.test(low) ? 'morgen' : 'heute');
      });
    collect(full, [/(?<![\d.])(?:0?[1-9]|[12]\d|3[01])\.\s?(?:0?[1-9]|1[0-2])\.(?:\s?(?:20)?\d{2})?(?![\d])/g])
      .forEach(function (v) {
        var d = v.text.replace(/\s+/g, '').replace(/\.(?:20)?\d{2}$/, '.').replace(/^0?(\d)\./, '0$1.').replace(/\.0?(\d)\.$/, '.0$1.');
        if (d.indexOf(heuteD) === 0) add('zeit', v.text, 'heute');
        else if (d.indexOf(morgenD) === 0) add('zeit', v.text, 'morgen');
      });
    var wtRe = new RegExp(WORD_L + '(?:bis\\s+|am\\s+|für\\s+|vor\\s+|spätestens\\s+)?(' + wtHeute + '|' + wtMorgen + ')' + WORD_R, 'gi');
    collect(full, [wtRe]).forEach(function (v) {
      var low = v.text.toLowerCase();
      add('zeit', v.text, low.indexOf(wtHeute.toLowerCase()) !== -1 ? 'heute' : 'morgen');
    });

    return out;
  }

  // ---------- Kernaussage: Was / Bis wann / Anhang ----------
  function baueKern(a, full) {
    var f = a.fields;
    var first = function (r) { return r.values.length ? r.values[0].text : ''; };

    // WAS
    var was = { text: '', fehlt: [] };
    if (a.positionen.length >= 2) {
      var unvoll = a.positionen.filter(function (p) { return p.fehlt.length; }).length;
      was.text = a.positionen.length + ' Positionen: ' +
        a.positionen.slice(0, 3).map(function (p) { return p.text.split(/[,(]/)[0].trim(); }).join(', ') +
        (a.positionen.length > 3 ? ' …' : '');
      if (unvoll) was.fehlt.push(unvoll + ' unvollständig');
    } else {
      var teile = [];
      var menge = f.menge.state === 'ok'
        ? first(f.menge).replace(/^(?:Menge|Auflage|Anzahl|Stückzahl|Stueckzahl)\s*[:=]?\s*/i, '') : '';
      var art = collect(full, [DRUCKART]);
      art = art.length ? art[0].text : '';
      if (menge) teile.push(menge);
      if (art && !(menge && menge.toLowerCase().indexOf(art.toLowerCase()) !== -1)) teile.push(art);
      var fmt = f.format.state === 'ok' ? first(f.format) : '';
      if (fmt) teile.push(fmt);
      was.text = teile.join(' · ');
      if (!menge && f.menge.required) was.fehlt.push('Menge');
      if (!fmt && f.format.required) was.fehlt.push('Format');
      if (!was.text) was.text = (a.punkte[0] || '').slice(0, 110);
    }

    // BIS WANN
    var t = f.termin;
    var teilT = t.values.filter(function (v) {
      return /\d{1,2}\.\s?\d{1,2}\.|^KW|Kalenderwoche|^(bis|spätestens|am|zum|für|vor)\s|^(heute|morgen|übermorgen|nächste Woche|diese Woche|Ende der Woche|Ende des Monats)$|Uhr/i.test(v.text);
    });
    var datum = teilT.filter(function (v) { return !/Uhr/i.test(v.text); })[0];
    var uhr = teilT.filter(function (v) { return /Uhr/i.test(v.text); })[0];
    var wann = '';
    if (datum) wann = datum.text.replace(/\s+\d{1,2}$/, '');
    if (uhr) wann = (wann ? wann + ' ' : '') + uhr.text.replace(/^bis\s+/i, '');
    var eilig = !!t.eilig || a.urgent;
    var bisWann = {
      text: wann || (eilig ? 'dringend, Datum fehlt' : ''),
      state: t.state === 'ok' ? 'ok' : (wann || eilig ? 'unklar' : 'fehlt'),
      uhrzeit: !!t.hasUhrzeit,
      eilig: eilig
    };

    // ANHANG
    var d = f.daten;
    var namen = a.thread.attachmentNames.slice();
    var anhang = { state: 'nein', namen: namen, erwaehnt: false, fehlend: a.thread.missingFiles.slice() };
    if (namen.length && anhang.fehlend.length) anhang.state = 'teilweise';
    else if (namen.length) anhang.state = 'ja';
    else if (d.vorhanden) anhang.state = 'geliefert';
    else if (d.folgt) anhang.state = 'folgt';
    else if (d.state === 'ok') anhang.state = 'ja';
    else anhang.erwaehnt = a.risiken.some(function (r) { return r.typ === 'anhang_fehlt' || r.typ === 'weiterleitung_ohne_anhang'; });

    return { was: was, bisWann: bisWann, anhang: anhang };
  }

  // ---------- Positionen (mehrere Druckarbeiten in einer Mail / auf einem Laufzettel) ----------
  var DRUCKART = wb('Laserdruck|Digitaldruck|Offsetdruck|Gro(?:ß|ss)format|LFP|Plot(?:s|ter)?|Plakate?|Poster|Flyer|Briefe?|Brosch(?:ü|ue)ren?|Visitenkarten?|Postkarten?|Karten|Banner|Roll-?ups?|Folder|Prospekte?|Kataloge?|Etiketten?|Aufkleber|Couverts?|Kuverts?|Mappen?|Kopien|Ausdrucke?|Preislisten?|Pläne|Kalender|Speisekarten?|Schilder|Leinwand|Fotoposter', 'gi');
  var FORTSETZUNG = /^\s*(?:(?:Auflage|Format|Papier|Druck|Daten|Menge|Farbe|Seiten|Stück|Material)\s*:|\d[\s-]?seitig|ein-?seitig|zwei-?seitig|beid-?seitig|doppel-?seitig)/i;
  var SCHWACH = {
    format: /^(Visitenkarten?|Postkarten?|Roll-?up|Quadratisch(?:es)?|Hochformat|Querformat|Poster|Plan|Pl[äa]ne)$/i,
    papier: /^(Papier|Material)$/i,
    farbe: /^Farbe$/i
  };

  function ersterTreffer(text, key) {
    var vals = collect(text, PATTERNS[key]());
    if (SCHWACH[key]) vals = vals.filter(function (v) { return !SCHWACH[key].test(v.text); });
    return vals.length ? vals.map(function (v) { return v.text; }).join(', ') : '';
  }

  function analysePositionen(text) {
    var lines = String(text).split(/\r?\n/);
    var blocks = [];
    var cur = null;
    lines.forEach(function (line) {
      if (!line.trim()) { cur = null; return; }
      DRUCKART.lastIndex = 0;
      var hatArt = DRUCKART.test(line);
      if (hatArt && !FORTSETZUNG.test(line)) {
        cur = { zeilen: [line.trim()] };
        blocks.push(cur);
      } else if (cur) {
        cur.zeilen.push(line.trim());
      }
    });

    var out = [];
    blocks.forEach(function (b) {
      var t = b.zeilen.join('\n');
      var pos = {
        text: b.zeilen[0].replace(/^[-•*·\d.)\s]+(?=[A-Za-zÄÖÜ])/, '').slice(0, 80),
        menge: ersterTreffer(t, 'menge'),
        format: ersterTreffer(t, 'format'),
        papier: ersterTreffer(t, 'papier'),
        farbe: ersterTreffer(t, 'farbe')
      };
      // Zahl am Zeilenanfang = Menge (wie die Mengenspalte auf dem Laufzettel)
      var lead = /^\s*(\d{1,5})\s+(?=[A-Za-zÄÖÜ])/.exec(b.zeilen[0]);
      if (lead && !pos.menge) pos.menge = lead[1];
      // Eine Position braucht eine Menge oder zumindest Format und Papier,
      // sonst ist es nur eine Aufzählung ("Inhalte: - Poster A2")
      if (!(pos.menge || (pos.format && pos.papier))) return;
      pos.fehlt = [];
      if (!pos.menge) pos.fehlt.push('Menge');
      if (!pos.format) pos.fehlt.push('Format');
      if (!pos.papier) pos.fehlt.push('Papier');
      if (!pos.farbe) pos.fehlt.push('Farbe');
      out.push(pos);
    });
    return out.length >= 2 ? out : [];
  }

  // ---------- Verlauf zerlegen ----------
  var DATEI_IN_TEXT = /(?<![A-Za-zÄÖÜäöüß0-9_\-])[A-Za-zÄÖÜäöüß0-9_\-+.()]+\.(?:pdf|ai|indd|idml|psd|tif|tiff|jpg|jpeg|png|eps|svg|zip|rar|7z|dwg|dxf|docx?|pptx?|xlsx?|cdr|afdesign|afpub|afphoto)(?![A-Za-z0-9])/gi;
  var HEADER_START = /^[ \t>]*(?:Von|From)\s*:\s*(.*)$/gmi;
  var AM_SCHRIEB = /^[ \t>]*(?:Am|On)\s+(.{6,60}?)\s+(?:schrieb|wrote)\s+(.{2,80}?):\s*$/gmi;

  function istBrauchbareDatei(n) {
    return n && !DATEI_IGNORE.test(n) && /\.[A-Za-z0-9]{2,5}$/.test(n) && !/^\.+/.test(n);
  }

  function analyseVerlauf(text, subject, attachments, from) {
    var attachmentNames = (attachments || []).map(function (a) {
      return typeof a === 'string' ? a : (a && (a.name || a.fileName)) || '';
    }).filter(istBrauchbareDatei);

    var forwarded = /^\s*(WG|FW|FWD|Fwd|Weitergeleitet)\s*:/i.test(subject) ||
      /(Weitergeleitete Nachricht|Forwarded message|Begin forwarded message|Ursprüngliche Nachricht|Original Message|-----\s*Original)/i.test(text);

    // Genannte Dateien (nur ohne Signatur-Bilder)
    var mentioned = [], seen = {};
    DATEI_IN_TEXT.lastIndex = 0;
    var m;
    while ((m = DATEI_IN_TEXT.exec(text)) !== null) {
      var n = m[0].replace(/[.)(]+$/, '');
      if (!istBrauchbareDatei(n) || /^www\./i.test(n)) continue;
      var k = n.toLowerCase();
      if (!seen[k]) { seen[k] = true; mentioned.push(n); }
    }
    var have = attachmentNames.map(function (a) { return a.toLowerCase(); });
    var missingFiles = mentioned.filter(function (n) { return have.indexOf(n.toLowerCase()) === -1; });

    var erwaehntAnhang = /\b(im Anhang|anbei|beigefügt|angehängt|angefügt|siehe Anhang|Anhang:|Anlage|Dateien? anbei|Daten anbei)\b/i.test(text);

    // Nachrichten trennen
    var cuts = [];
    HEADER_START.lastIndex = 0;
    while ((m = HEADER_START.exec(text)) !== null) {
      cuts.push({ index: m.index, kind: 'header', absender: (m[1] || '').trim(), end: m.index + m[0].length });
    }
    AM_SCHRIEB.lastIndex = 0;
    while ((m = AM_SCHRIEB.exec(text)) !== null) {
      cuts.push({ index: m.index, kind: 'schrieb', absender: m[2].trim(), datum: m[1].trim(), end: m.index + m[0].length });
    }
    cuts.sort(function (a, b) { return a.index - b.index; });

    var messages = [];
    var firstText = cuts.length ? text.slice(0, cuts[0].index) : text;
    messages.push({ von: from || '', datum: '', body: firstText });
    cuts.forEach(function (c, i) {
      var stop = i + 1 < cuts.length ? cuts[i + 1].index : text.length;
      var block = text.slice(c.end, stop).replace(/^[ \t]*\r?\n/, '');
      var datum = c.datum || '';
      if (c.kind === 'header') {
        var dm = /(?:Gesendet|Sent|Datum|Date)\s*:\s*(.+)/i.exec(block.slice(0, 500));
        if (dm) datum = dm[1].trim();
        // Kopfzeilen (An:, Cc:, Betreff: ...) entfernen
        block = block.replace(/^(?:[ \t>]*(?:Gesendet|Sent|Datum|Date|An|To|Cc|Bcc|Betreff|Subject|Wichtigkeit|Importance|Anlagen|Attachments|Anhänge)\s*:.*\r?\n)+/i, '');
      }
      messages.push({ von: cleanAbsender(c.absender), datum: datum, body: block });
    });

    messages = messages.map(function (msg) {
      var pts = wichtigePunkte(msg.body, null).slice(0, 3);
      return {
        von: msg.von, datum: msg.datum, punkte: pts,
        score: infoScore(msg.body),
        leer: msg.body.replace(/\s+/g, '').length < 5
      };
    }).filter(function (msg, i) { return !msg.leer || (i === 0 && messages.length === 1); });

    // Welche Nachricht enthält die meiste Auftragsinfo? (bei Gleichstand die neueste)
    var best = 0;
    messages.forEach(function (m, i) { if (m.score > messages[best].score) best = i; });

    return {
      wichtigste: best,
      forwarded: forwarded,
      attachmentNames: attachmentNames,
      mentionedFiles: mentioned,
      missingFiles: missingFiles,
      erwaehntAnhang: erwaehntAnhang,
      messages: messages
    };
  }

  // Wie viel Auftragsinfo steckt in einem Textstück? 0 = Small Talk ("Danke, schönes Wochenende!")
  function infoScore(body) {
    var s = 0;
    ['format', 'menge', 'papier', 'farbe', 'termin', 'lieferung'].forEach(function (k) {
      var vals = collect(body, PATTERNS[k]());
      if (SCHWACH[k]) vals = vals.filter(function (v) { return !SCHWACH[k].test(v.text); });
      if (vals.length) s++;
    });
    DATEI_IN_TEXT.lastIndex = 0;
    if (DATEI_IN_TEXT.test(body) || /(?<![A-Za-zÄÖÜäöüß])(im Anhang|anbei|beigefügt|Daten folgen)(?![A-Za-zÄÖÜäöüß])/i.test(body)) s++;
    if (collect(body, [URGENT]).length) s++;
    return s;
  }

  function cleanAbsender(s) {
    return String(s || '').replace(/<[^>]*>/g, '').replace(/\[mailto:[^\]]*\]/gi, '').replace(/["']/g, '').replace(/\s+/g, ' ').trim();
  }

  function wichtigePunkte(full, result) {
    var lines = full.split(/\r?\n|(?<=[.!?])\s+(?=[A-ZÄÖÜ])/).map(function (s) { return s.trim(); })
      .filter(function (s) { return s.length > 8 && s.length < 400; });
    // Zitat-Ketten und Signaturen aussortieren
    lines = lines.filter(function (s) {
      return !/^(>|Von:|From:|Gesendet:|Sent:|An:|To:|Cc:|Betreff:|Subject:|-----)/i.test(s) &&
             !/^(Mit freundlichen|Freundliche|Viele Gr|Beste Gr|Gruß|Gruss|Hallo|Guten (Tag|Morgen)|Sehr geehrte|Liebe)/i.test(s);
    });
    var scored = lines.map(function (s) {
      var score = 0;
      FIELDS.forEach(function (f) {
        var maker = PATTERNS[f.key];
        if (!maker) return;
        if (collect(s, maker()).length) score += 1;
      });
      if (collect(s, [URGENT]).length) score += 2;
      return { text: s, score: score };
    }).filter(function (x) { return x.score > 0; });
    scored.sort(function (a, b) { return b.score - a.score; });
    return scored.slice(0, 6).map(function (x) { return x.text; });
  }

  // ---------- Rückfrage-Text ----------
  function rueckfrage(analysis, anrede) {
    var m = analysis.missing;
    if (!m.length) return '';
    var zeilen = [];
    zeilen.push(anrede || 'Guten Tag,');
    zeilen.push('');
    zeilen.push('vielen Dank für Ihre Anfrage. Damit wir den Auftrag sauber vorbereiten können, fehlen uns noch folgende Angaben:');
    zeilen.push('');
    m.forEach(function (r, i) {
      var f = analysis.fields[r.key];
      var prefix = (i + 1) + '. ';
      if (r.key === 'daten' && f.risiko) {
        var mf = (analysis.thread && analysis.thread.missingFiles) || [];
        zeilen.push(prefix + 'Druckdaten: Der Anhang ist bei uns nicht angekommen' +
          (mf.length ? ' (es fehlen: ' + mf.join(', ') + ')' : '') + '. Bitte senden Sie die Dateien erneut.');
      } else if (r.key === 'daten' && f.folgt) {
        zeilen.push(prefix + 'Druckdaten: Bis wann können Sie uns die Daten zusenden?');
      } else if (r.key === 'termin' && f.values.length) {
        zeilen.push(prefix + 'Termin: Bitte nennen Sie uns das genaue Datum' + (f.hasUhrzeit ? '' : ' und die Uhrzeit') + '.');
      } else {
        zeilen.push(prefix + f.label + ': ' + f.frage);
      }
    });
    zeilen.push('');
    zeilen.push('Sobald uns diese Angaben vorliegen, legen wir direkt los.');
    zeilen.push('');
    zeilen.push('Freundliche Grüße');
    return zeilen.join('\n');
  }

  function rueckfrageHtml(analysis, anrede) {
    var txt = rueckfrage(analysis, anrede);
    var esc = txt.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return '<div style="font-family:Segoe UI,Calibri,Arial,sans-serif;font-size:11pt">' +
           esc.replace(/\n/g, '<br>') + '</div>';
  }

  root.Auftrag = {
    FIELDS: FIELDS,
    analyse: analyse,
    rueckfrage: rueckfrage,
    rueckfrageHtml: rueckfrageHtml
  };
})(typeof window !== 'undefined' ? window : this);
