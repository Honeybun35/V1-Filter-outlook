/* Das Team. Hier Kürzel und Farben ändern.
 * kuerzel : Kurzform in der Seitenleiste
 *           (1 Buchstabe Vorname + 2 Buchstaben Nachname, der zweite Nachnamen-Buchstabe klein)
 * name    : Anzeige in der Seitenleiste (hier bewusst nur das Kürzel, weil die Datei öffentlich liegt)
 * farbe   : Farbe in der Seitenleiste
 * preset  : Kategorie-Farbe in Outlook (0 Rot, 1 Orange, 2 Pfirsich, 3 Gelb, 4 Grün, 5 Türkis,
 *           6 Olive, 7 Blau, 8 Lila, 9 Bordeaux, 10 Stahl, 11 Dunkelstahl, 12 Grau, 13 Dunkelgrau,
 *           14 Schwarz, 15 Dunkelrot, 16 Dunkelorange ...)
 */
window.TEAM = [
  // kategorie: genauer Name der Kategorie in Outlook (so wie in eurer Kategorienliste)
  { kuerzel: 'SRu', name: 'SRu', kategorie: 'Simone Ruppli',    farbe: '#e07b00', preset: 1 },   // Orange
  { kuerzel: 'THe', name: 'THe', kategorie: 'Thomas Heuberger', farbe: '#1d6fd1', preset: 7 },   // Blau
  { kuerzel: 'FSt', name: 'FSt', kategorie: 'Florian Stauber',  farbe: '#1e8f4e', preset: 4 },   // Grün
  { kuerzel: 'JRe', name: 'JRe', kategorie: 'JRe',              farbe: '#8e44ad', preset: 8 },   // Lila (vollen Namen eintragen)
  { kuerzel: 'VHa', name: 'VHa', kategorie: 'VHa',              farbe: '#6b7280', preset: 12 },  // Grau (vollen Namen eintragen)
  // Kein Bearbeiter: es wird nur der Status gesetzt, keine Namens-Kategorie
  { kuerzel: 'Offen', name: 'Kein Bearbeiter', kategorie: '', farbe: '#9ca3af', preset: 13, keinePerson: true }
];
