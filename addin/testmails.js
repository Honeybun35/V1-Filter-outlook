/* Beispielmails zum Testen (Demo-Modus und tests.html).
 * "erwartet" sagt, welcher Zustand je Feld herauskommen soll. */
window.TESTMAILS = [
  {
    name: 'A3 ohne Details',
    subject: 'Ausdrucke für Kunde Müller',
    text: 'Hallo zusammen,\nder Kunde Müller will A3 Ausdrucke. Bitte macht das bis Freitag.\nViele Grüße\nHerr H',
    attachments: [],
    erwartet: { format: 'ok', menge: 'fehlt', papier: 'fehlt', farbe: 'fehlt', termin: 'ok', daten: 'fehlt', lieferung: 'fehlt' }
  },
  {
    name: 'Fast komplett',
    subject: 'Auftrag 2024-0815 Plakate',
    text: 'Guten Morgen,\nbitte 250 Stück Plakate A2 auf 135 g Bilderdruck matt, 4/0, bis 14.11. 12 Uhr. Abholung durch den Kunden. Daten im Anhang.\nDanke',
    attachments: ['plakat_final.pdf', 'image001.png'],
    erwartet: { format: 'ok', menge: 'ok', papier: 'ok', farbe: 'ok', termin: 'ok', daten: 'ok', lieferung: 'ok', auftrag: 'ok' }
  },
  {
    name: 'Daten folgen',
    subject: 'Banner für Messe',
    text: 'Hi, wir brauchen 2 Banner 85 x 200 cm auf Blueback, farbig. Daten folgen noch. Wichtig, dringend! Lieferung an die Messe in 80331 München.',
    attachments: [],
    erwartet: { format: 'ok', menge: 'ok', papier: 'ok', farbe: 'ok', termin: 'unklar', daten: 'folgt', lieferung: 'ok' }
  },
  {
    name: 'Nur Gruß (nichts drin)',
    subject: 'AW: Anfrage',
    text: 'Guten Morgen,\nkönnt ihr mir bitte helfen?\nDanke und schöne Grüße',
    attachments: [],
    erwartet: { format: 'fehlt', menge: 'fehlt', papier: 'fehlt', farbe: 'fehlt', termin: 'fehlt', daten: 'fehlt', lieferung: 'fehlt' }
  },
  {
    name: 'Weiterleitung ohne Anhang (5 Dateien genannt)',
    subject: 'WG: Plakate Frühling',
    from: 'Max Muster',
    text: 'Zur Info, bitte umsetzen.\n\nVon: Max Muster <m.muster@kunde.ch>\nGesendet: Montag, 3. November 2025 09:12\nAn: Chef\nBetreff: Plakate Frühling\n\nHallo,\nbitte 50 Plakate A2 auf 150 g matt, 4/0, bis 14.11. Abholung.\nDateien: Plakat_1.pdf, Plakat_2.pdf, Plakat_3.pdf, Plakat_4.pdf, Plakat_5.pdf\nDanke',
    attachments: ['image001.png'],
    erwartet: { daten: 'fehlt', format: 'ok', menge: 'ok' },
    risiko: 'anhang_fehlt',
    fehlende: 5,
    nachrichten: 2
  },
  {
    name: 'Anhang erwähnt, nichts angehängt',
    subject: 'Flyer',
    text: 'Hallo, anbei die Daten für 500 Flyer A5, 135 g, 4/4, bis Freitag, Lieferung an die Firma.',
    attachments: [],
    erwartet: { daten: 'fehlt', menge: 'ok' },
    risiko: 'anhang_fehlt',
    fehlende: 0
  },
  {
    name: 'Weiterleitung ohne Anhang (keine Datei genannt)',
    subject: 'FW: Neue Anfrage',
    text: 'Bitte schaut euch das an.\n\nVon: Kunde <k@k.ch>\nGesendet: heute\nBetreff: Neue Anfrage\n\nWir brauchen A3 Ausdrucke.',
    attachments: [],
    erwartet: { format: 'ok' },
    risiko: 'weiterleitung_ohne_anhang',
    nachrichten: 2
  },
  {
    name: 'Alles vollständig angehängt',
    subject: 'WG: Plakat',
    text: 'Zur Info.\n\nVon: Max Muster\nGesendet: heute\nBetreff: Plakat\n\n20 Plakate A2, 150 g matt, 4/0, bis 20.11., Abholung. Datei: plakat.pdf',
    attachments: ['plakat.pdf'],
    erwartet: { daten: 'ok' },
    risiko: null
  },
  {
    name: 'Laufzettel (Schweizer Schreibweise, mehrere Positionen)',
    subject: 'Auftrag-Nr. 532182 Versand Frühlingsmarketing',
    text: 'Auftrag-Nr. 532182\nAlt-Auftrags-Nr. 532181\nLiefertermin 18.04.2023\nLieferadresse: Beispiel AG, Musterstrasse 11, 8804 Au ZH\n\nVersandarbeiten an 14 Adressen\nInhalte:\n- Brief mit Personalisierung\n- Info-Flyer A5, 2-seitig\n- Poster A2\n\n28 Laserdruck A4 farbig, 80 gm2 Offset\n1-seitig: 14x Brief\n\n14 Laserdruck A4 farbig, 120 gm2 DCP\n1-seitig: 14x Versicherungsliste intern\n\n14 Laminieren A4 - Versicherungsliste intern\n\n50 Laserdruck, SRA3, farbig zweiseitig, 160g Image Impact\nAuflage: 170 x Versicherungspartnerflyer, A5\nDruck zu 4 Ntz.\n\n40 Laserdruck, farbig zweiseitig, 160g Image Impact\nAuflage: 70 x Info-Flyer\nDruck zu 2 Ntz.\n\n6 Plakat A2 (LFP) Photopapier inkl. Datenübernahme\nFormat: 42 x 59.4 cm\nDruck: einseitig 4-farbig Digitaldruck\nDaten: druckfertiges PDF geliefert\nPapier: 195 gm2, Fotopapier Maranello seidenmatt PE\n\nPlakate zuschneiden\n\nPostaufgabe\nPorto B-Post (Pakete)',
    attachments: [],
    erwartet: { format: 'ok', menge: 'ok', papier: 'ok', farbe: 'ok', termin: 'ok', daten: 'ok', lieferung: 'ok', weiter: 'ok', auftrag: 'ok' },
    risiko: 'referenz',
    positionenMin: 4
  },
  {
    name: 'Pflicht: nur Banner (Farbe nicht Pflicht)',
    subject: 'Banner',
    text: 'Bitte 2 Banner 85 x 200 cm auf Blueback bis Freitag, Abholung.',
    attachments: ['banner.pdf'],
    erwartet: { farbe: 'fehlt' },
    pflichtNicht: ['farbe'],
    pflichtJa: ['format', 'menge', 'papier']
  },
  {
    name: 'Pflicht: Mischauftrag Plakat + Flyer (Farbe bleibt Pflicht)',
    subject: 'Plakat und Flyer',
    text: '6 Plakate A2 und 100 Flyer A5 auf 135 g, bis Freitag, Abholung.',
    attachments: ['daten.pdf'],
    erwartet: { farbe: 'fehlt' },
    pflichtJa: ['farbe', 'papier']
  },
  {
    name: 'Pflicht: wie letztes Mal',
    subject: 'Nachdruck',
    text: 'Bitte nochmal 200 Stück wie letztes Mal, bis Freitag, Lieferung wie immer.',
    attachments: [],
    erwartet: { menge: 'ok' },
    pflichtNicht: ['format', 'papier', 'farbe'],
    pflichtJa: ['menge', 'termin', 'daten', 'lieferung'],
    risiko: 'referenz'
  },
  {
    name: 'Pflicht: Visitenkarten (Format nicht Pflicht)',
    subject: 'Visitenkarten',
    text: '250 Visitenkarten auf 350 g Karton, 4/4, bis 14.11., Abholung. Daten anbei.',
    attachments: ['vk.pdf'],
    erwartet: { menge: 'ok' },
    pflichtNicht: ['format'],
    pflichtJa: ['papier', 'farbe']
  },
  {
    name: 'SRA3 mit Grammatur gm2',
    subject: 'Flyer',
    text: '100 Stück Flyer SRA3, 170 gm2 matt, farbig beidseitig, bis Freitag, Abholung. Daten folgen.',
    attachments: [],
    erwartet: { format: 'ok', menge: 'ok', papier: 'ok', farbe: 'ok', termin: 'ok', daten: 'folgt', lieferung: 'ok' }
  },
  {
    name: 'Maße und Weiterverarbeitung',
    subject: 'Plots Bauplan',
    text: 'Bitte 10x Plan 841 x 594 mm, s/w, 90 g/m² Plotterpapier, gefalzt auf A4. Bis morgen. Lieferung an Baustelle.',
    attachments: ['Plan_EG.pdf'],
    erwartet: { format: 'ok', menge: 'ok', papier: 'ok', farbe: 'ok', termin: 'ok', daten: 'ok', lieferung: 'ok', weiter: 'ok' }
  }
];
