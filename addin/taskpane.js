/* Seitenleiste "Auftrags-Check" für Outlook (Office.js, Mailbox 1.8)
 * Reiter 1 "Diese Mail": Kernaussage (Was / Bis wann / Anhang), Bearbeiter + Status, Aktivität, Verlauf
 * Reiter 2 "Übersicht": alle Mails mit Filter nach MA und Status (Datenquelle: später Microsoft Graph) */
(function () {
  'use strict';

  var STATUS = ['Warten', 'Daten folgen', 'In Bearbeitung', 'GZD warten', 'Erledigt', 'Zur Kenntnisnahme'];
  // Schriftfarbe je Status (Akzent, sonst ist alles weiß)
  var STATUS_FARBE = {
    'Warten': '#b7791f', 'Daten folgen': '#c05621', 'In Bearbeitung': '#2b6cb0',
    'GZD warten': '#6b46c1', 'Erledigt': '#2f855a', 'Zur Kenntnisnahme': '#0f766e'
  };
  var SEP = ' · ';
  var STATUS_RE = new RegExp('^(.+?)' + SEP.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(' + STATUS.join('|') + ')$');
  var TEAM = window.TEAM || [];
  var CORE = ['menge', 'format', 'termin', 'daten']; // Was (Menge, Format), Bis wann, Anhang

  function person(k) {
    for (var i = 0; i < TEAM.length; i++) if (TEAM[i].kuerzel === k) return TEAM[i];
    return { kuerzel: k, name: k, farbe: '#6b7280', preset: 12 };
  }

  var state = {
    demo: false,
    item: null,
    analysis: null,
    categories: [],
    me: '',
    team: TEAM.map(function (t) { return t.kuerzel; }),
    person: '',
    itemKey: '',
    subject: '',
    log: [],
    props: null,
    demoLogs: {},
    demoCats: {},
    detailsInReply: false,
    overview: { items: [], q: '', ma: 'alle', st: 'alle', sort: 'wichtig' }
  };

  var $ = function (id) { return document.getElementById(id); };
  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text != null) e.textContent = text;
    return e;
  }

  // ---------- Einstellungen ----------
  function loadSettings() {
    var me = '';
    try {
      if (!state.demo && Office.context.roamingSettings) me = Office.context.roamingSettings.get('me') || '';
      else me = localStorage.getItem('ac_me') || '';
    } catch (e) { /* ignorieren */ }
    state.me = me;
    state.person = me || state.team[0];
    var sel = $('meSelect');
    sel.innerHTML = '<option value="">– bitte wählen –</option>';
    TEAM.forEach(function (t) {
      if (t.keinePerson) return;
      var o = el('option', null, t.name);
      o.value = t.kuerzel;
      sel.appendChild(o);
    });
    sel.value = me;
    applyMe();
  }

  function saveSettings() {
    var me = $('meSelect').value;
    state.me = me;
    try {
      if (!state.demo && Office.context.roamingSettings) {
        Office.context.roamingSettings.set('me', me);
        Office.context.roamingSettings.saveAsync(function () {});
      } else {
        localStorage.setItem('ac_me', me);
      }
      $('settingsMsg').textContent = 'Gespeichert.';
    } catch (e) {
      $('settingsMsg').textContent = 'Konnte nicht gespeichert werden.';
    }
    applyMe();
    syncPersonFromCategories();
    render();
  }

  // Akzentfarbe der Oberfläche = Farbe der eigenen Person
  function applyMe() {
    var p = state.me ? person(state.me) : null;
    document.documentElement.style.setProperty('--accent', p ? p.farbe : '#374151');
    var tag = $('meTag');
    tag.textContent = p ? p.kuerzel : '';
    tag.style.color = p ? p.farbe : '';
  }

  // ---------- Mail laden ----------
  function loadItem() {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { showEmpty(); return; }
    state.item = item;
    state.itemKey = item.itemId || item.conversationId || item.subject || '';
    var subject = item.subject || '';
    var attachments = (item.attachments || []).map(function (a) { return a.name; });

    holeMailtext(item, function (text) {
      readCategories(function () {
        loadLog(function () {
          var from = (item.from && (item.from.displayName || item.from.emailAddress)) || '';
          state.analysis = Auftrag.analyse({ subject: subject, text: text, attachments: attachments, from: from });
          state.subject = subject;
          syncPersonFromCategories();
          render();
        });
      });
    });
  }

  // Mailtext holen. Manche Outlook-Versionen liefern den Text ohne Zeilenumbrüche.
  // Dann wird der HTML-Text geholt und selbst in Zeilen zerlegt.
  function holeMailtext(item, done) {
    item.body.getAsync(Office.CoercionType.Text, function (res) {
      var text = (res.status === Office.AsyncResultStatus.Succeeded) ? (res.value || '') : '';
      var zeilen = (text.match(/\n/g) || []).length;
      if (text.length < 120 || zeilen >= 3) { done(text); return; }
      item.body.getAsync(Office.CoercionType.Html, function (hr) {
        if (hr.status !== Office.AsyncResultStatus.Succeeded || !hr.value) { done(text); return; }
        var t = htmlZuText(hr.value);
        done(t.length >= text.length * 0.5 ? t : text);
      });
    });
  }

  function htmlZuText(html) {
    try {
      var doc = new DOMParser().parseFromString(html, 'text/html');
      var i, rm = doc.querySelectorAll('script,style,head');
      for (i = 0; i < rm.length; i++) rm[i].parentNode.removeChild(rm[i]);
      var br = doc.querySelectorAll('br');
      for (i = 0; i < br.length; i++) br[i].parentNode.replaceChild(doc.createTextNode('\n'), br[i]);
      var li = doc.querySelectorAll('li');
      for (i = 0; i < li.length; i++) li[i].insertBefore(doc.createTextNode('- '), li[i].firstChild);
      var blk = doc.querySelectorAll('p,h1,h2,h3,h4,h5,h6,blockquote');
      for (i = 0; i < blk.length; i++) blk[i].appendChild(doc.createTextNode('\n\n'));
      var row = doc.querySelectorAll('div,li,tr,table,ul,ol');
      for (i = 0; i < row.length; i++) row[i].appendChild(doc.createTextNode('\n'));
      var t = (doc.body ? doc.body.textContent : '') || '';
      return t.replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{4,}/g, '\n\n\n');
    } catch (e) {
      return '';
    }
  }

  function readCategories(done) {
    var item = state.item;
    if (item && item.categories && item.categories.getAsync) {
      item.categories.getAsync(function (r) {
        state.categories = (r.status === Office.AsyncResultStatus.Succeeded && r.value)
          ? r.value.map(function (c) { return c.displayName; }) : [];
        done();
      });
    } else {
      state.categories = [];
      done();
    }
  }

  function currentStatus() {
    for (var i = 0; i < state.categories.length; i++) {
      var m = STATUS_RE.exec(state.categories[i]);
      if (m) return { person: m[1], status: m[2] };
    }
    return null;
  }

  function syncPersonFromCategories() {
    var cur = currentStatus();
    if (cur) {
      if (state.team.indexOf(cur.person) === -1) state.team.push(cur.person);
      state.person = cur.person;
    } else {
      state.person = state.me || state.team[0];
    }
  }

  // ---------- Aktivität (wer hat wann was gesetzt) ----------
  // Gespeichert als eigene Eigenschaft an der Mail, damit alle im gemeinsamen Postfach es sehen.
  // Format je Eintrag: Akteur|Status-Nr (oder x)|Minuten seit 1970|Bearbeiter
  function parseLog(raw) {
    return String(raw || '').split(';').filter(Boolean).map(function (p) {
      var a = p.split('|');
      var s = a[1] === 'x' ? 'x' : STATUS[parseInt(a[1], 10)];
      return { p: a[0], s: s, t: parseInt(a[2], 10) * 60000, a: a[3] || '' };
    }).filter(function (e) { return e.s && !isNaN(e.t); });
  }

  function serializeLog(list) {
    return list.map(function (e) {
      return [e.p, e.s === 'x' ? 'x' : STATUS.indexOf(e.s), Math.floor(e.t / 60000), e.a || ''].join('|');
    }).join(';');
  }

  function loadLog(done) {
    state.log = [];
    state.notiz = '';
    state.props = null;
    if (!state.item || !state.item.loadCustomPropertiesAsync) { done(); return; }
    state.item.loadCustomPropertiesAsync(function (r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        state.props = r.value;
        state.log = parseLog(r.value.get('ac_log'));
        state.notiz = r.value.get('ac_notiz') || '';
      }
      done();
    });
  }

  function addLog(actor, status, assignee) {
    state.log.unshift({ p: actor, s: status, t: Date.now(), a: assignee || '' });
    state.log = state.log.slice(0, 12);
    if (state.demo) { state.demoLogs[state.itemKey] = state.log; return; }
    if (state.props) {
      try {
        state.props.set('ac_log', serializeLog(state.log));
        state.props.saveAsync(function () {});
      } catch (e) { /* Protokoll ist nur eine Zugabe */ }
    }
  }

  function fmtZeit(t) {
    var d = new Date(t), n = new Date();
    var p2 = function (x) { return ('0' + x).slice(-2); };
    var hm = p2(d.getHours()) + ':' + p2(d.getMinutes());
    var tag0 = new Date(n.getFullYear(), n.getMonth(), n.getDate()).getTime();
    if (t >= tag0) return 'heute ' + hm;
    if (t >= tag0 - 86400000) return 'gestern ' + hm;
    return p2(d.getDate()) + '.' + p2(d.getMonth() + 1) + '. ' + hm;
  }

  // ---------- Anzeige: Diese Mail ----------
  function showEmpty() {
    state.item = null; state.analysis = null;
    $('subject').textContent = 'Keine Mail geöffnet';
    ['kWas', 'kWann', 'kAnhang', 'ownerLine', 'risks', 'log', 'checklist'].forEach(function (id) { $(id).innerHTML = ''; });
    ['eilig', 'sonderCard', 'kurz', 'logCard', 'replyCard', 'posBox', 'punkteBox'].forEach(function (id) { $(id).hidden = true; });
    $('kernCard').style.borderColor = '';
    renderStatus();
  }

  function render() {
    if (!state.analysis) { showEmpty(); return; }
    renderKern();
    renderPoly();
    renderSonder();
    renderNotiz();
    renderRisks();
    renderStatus();
    renderLog();
    renderReply();
    renderDetails();
  }

  function renderKern() {
    var a = state.analysis, k = a.kern;
    $('subject').textContent = state.subject || '(ohne Betreff)';
    $('eilig').hidden = !k.bisWann.eilig;

    // Bearbeiter-Zeile und Rahmen in der Farbe der Person
    var cur = currentStatus();
    var line = $('ownerLine');
    line.innerHTML = '';
    var col = '';
    if (cur) {
      var pe = person(cur.person);
      col = pe.farbe;
      var n = el('span', null, pe.name);
      n.style.color = pe.farbe;
      line.appendChild(n);
      line.appendChild(document.createTextNode(' · '));
      var s = el('span', null, cur.status);
      s.style.color = STATUS_FARBE[cur.status];
      line.appendChild(s);
    } else {
      line.appendChild(el('span', 'muted', 'Noch kein Bearbeiter'));
    }
    $('kernCard').style.borderColor = col;

    // Was
    var w = $('kWas');
    w.innerHTML = '';
    if (k.was.text) w.appendChild(document.createTextNode(k.was.text));
    k.was.fehlt.forEach(function (f, i) {
      var txt = /unvollständig/.test(f) ? f : f + ' fehlt';
      w.appendChild(el('span', 'bad', (k.was.text || i ? ' · ' : '') + txt));
    });
    if (!k.was.text && !k.was.fehlt.length) w.appendChild(el('span', 'bad', 'nicht erkennbar'));

    // Bis wann
    var wn = $('kWann');
    wn.innerHTML = '';
    if (k.bisWann.state === 'ok') {
      wn.appendChild(document.createTextNode(k.bisWann.text));
      if (!k.bisWann.uhrzeit) wn.appendChild(el('span', 'sub', 'Uhrzeit nicht genannt'));
    } else if (k.bisWann.state === 'unklar') {
      wn.appendChild(el('span', 'warn', k.bisWann.text || 'unklar'));
      wn.appendChild(el('span', 'sub', 'kein genaues Datum'));
    } else {
      wn.appendChild(el('span', 'bad', 'fehlt'));
    }

    // Anhang
    var an = k.anhang, ak = $('kAnhang');
    ak.innerHTML = '';
    if (an.state === 'ja') {
      ak.appendChild(el('span', 'ok', 'Ja'));
      if (an.namen.length) ak.appendChild(el('span', 'sub', an.namen.join(', ')));
    } else if (an.state === 'geliefert') {
      ak.appendChild(el('span', 'ok', 'Ja'));
      ak.appendChild(el('span', 'sub', 'Daten bereits geliefert'));
    } else if (an.state === 'folgt') {
      ak.appendChild(el('span', 'warn', 'Folgt noch'));
    } else if (an.state === 'teilweise') {
      ak.appendChild(el('span', 'bad', 'Teilweise'));
      ak.appendChild(el('span', 'sub', 'Es fehlen: ' + an.fehlend.join(', ')));
    } else {
      ak.appendChild(el('span', 'bad', 'Nein'));
      if (an.erwaehnt) ak.appendChild(el('span', 'sub', 'In der Mail erwähnt, aber nicht angehängt'));
    }
  }

  // ---------- Druckangaben, Sonderwünsche, Notiz ----------
  function zeigeListe(id, liste) {
    var d = $(id);
    d.innerHTML = '';
    if (liste && liste.length) d.appendChild(document.createTextNode(liste.slice(0, 3).join(', ')));
    else d.appendChild(el('span', 'none', 'nicht genannt'));
  }

  function kurzText(s) {
    return String(s).replace(/\s*(Seiten|Seiter)\b/i, ' S').replace(/Exemplare?/i, 'Ex').replace(/\s+/g, ' ').trim();
  }

  function renderPoly() {
    var a = state.analysis, pg = a.polygraf || {};
    zeigeListe('kSeiten', pg.seiten);
    zeigeListe('kDruck', pg.druck);
    zeigeListe('kProfil', pg.profil);
    if (pg.profilAbgeleitet) $('kProfil').appendChild(el('span', 'sub', 'vom Papier abgeleitet'));
    zeigeListe('kPapier', pg.papier);

    // Kurzfassung: nur das Wichtigste in Kurzform
    var teile = [];
    var k = a.kern;
    if (k.was.text) teile.push(kurzText(k.was.text));
    if (pg.seiten && pg.seiten[0]) teile.push(kurzText(pg.seiten[0]));
    if (pg.druck && pg.druck[0]) teile.push(kurzText(pg.druck[0]));
    if (pg.profil && pg.profil[0]) teile.push(kurzText(pg.profil[0]));
    if (pg.papier && pg.papier[0]) teile.push(kurzText(pg.papier[0]));
    if (k.bisWann.text) teile.push('bis ' + k.bisWann.text.replace(/^bis\s+/i, ''));
    var kz = $('kurz');
    kz.textContent = teile.join(' · ');
    kz.hidden = !teile.length;
  }

  function renderSonder() {
    var s = state.analysis.sonder || [];
    var card = $('sonderCard'), ul = $('sonderList');
    ul.innerHTML = '';
    card.hidden = !s.length;
    var titel = { format: 'Format', dringend: 'Dringend', zeit: 'Zeit' };
    s.forEach(function (x) {
      var li = el('li');
      li.appendChild(el('mark', null, x.text));
      li.appendChild(el('span', 'tag', (titel[x.typ] || '') + (x.hinweis && x.typ === 'zeit' ? ' (' + x.hinweis + ')' : x.typ === 'format' ? ' (kein DIN)' : '')));
      ul.appendChild(li);
    });
  }

  function renderNotiz() {
    $('notiz').value = state.notiz || '';
    $('notizMsg').textContent = '';
  }

  function saveNotiz() {
    if (!state.item) return;
    state.notiz = $('notiz').value.slice(0, 1500);
    if (state.demo) { $('notizMsg').textContent = 'Gespeichert (Demo).'; return; }
    if (!state.props) { $('notizMsg').textContent = 'Konnte nicht gespeichert werden.'; return; }
    try {
      state.props.set('ac_notiz', state.notiz);
      state.props.saveAsync(function (r) {
        $('notizMsg').textContent = r.status === Office.AsyncResultStatus.Succeeded ? 'Gespeichert.' : 'Konnte nicht gespeichert werden.';
      });
    } catch (e) { $('notizMsg').textContent = 'Konnte nicht gespeichert werden.'; }
  }

  function renderRisks() {
    var box = $('risks');
    box.innerHTML = '';
    (state.analysis.risiken || []).forEach(function (r) {
      var d = el('div', 'risk' + (r.typ !== 'anhang_fehlt' ? ' mild' : ''));
      d.appendChild(el('div', 'risk-title', (r.typ !== 'anhang_fehlt' ? '⚠ ' : '⛔ ') + r.titel));
      if (r.details && r.details.length) {
        var ul = el('ul');
        r.details.forEach(function (n) { ul.appendChild(el('li', null, n)); });
        d.appendChild(ul);
      }
      box.appendChild(d);
    });
  }

  function renderStatus() {
    var peopleBox = $('people');
    peopleBox.innerHTML = '';
    state.team.forEach(function (k) {
      var pe = person(k);
      var b = el('button', 'chip' + (k === state.person ? ' on' : ''), pe.name);
      b.type = 'button';
      b.title = pe.name;
      if (k === state.person) { b.style.borderColor = pe.farbe; b.style.color = pe.farbe; }
      b.onclick = function () { state.person = k; renderStatus(); };
      peopleBox.appendChild(b);
    });

    var cur = currentStatus();
    var suggest = state.analysis && state.analysis.vorschlag;
    var box = $('statuses');
    box.innerHTML = '';
    STATUS.forEach(function (s) {
      var b = el('button', 'status' + (cur && cur.status === s ? ' on' : '') + (!cur && suggest === s ? ' suggest' : ''), s);
      b.type = 'button';
      b.style.color = STATUS_FARBE[s];
      if (cur && cur.status === s) b.style.borderColor = STATUS_FARBE[s];
      b.title = (!cur && suggest === s) ? 'Vorschlag nach Mailtext' : '';
      b.onclick = function () { setStatus(s); };
      box.appendChild(b);
    });
    $('noStatus').hidden = !cur;
  }

  function renderLog() {
    var card = $('logCard'), ul = $('log');
    ul.innerHTML = '';
    if (!state.log.length) { card.hidden = true; return; }
    card.hidden = false;
    state.log.slice(0, 6).forEach(function (e) {
      var li = el('li');
      var n = el('b', null, e.p);
      n.style.color = person(e.p).farbe;
      li.appendChild(n);
      if (e.s === 'x') {
        li.appendChild(document.createTextNode(' hat den Status entfernt'));
      } else {
        li.appendChild(document.createTextNode(' hat auf '));
        var s = el('b', null, e.s);
        s.style.color = STATUS_FARBE[e.s];
        li.appendChild(s);
        li.appendChild(document.createTextNode(' gesetzt' + (e.a && e.a !== e.p ? ' (Bearbeiter ' + e.a + ')' : '')));
      }
      li.appendChild(el('span', 'zeit', fmtZeit(e.t)));
      ul.appendChild(li);
    });

    // Erinnerung: GZD wartet seit 48 Stunden
    var hint = $('gzdHint');
    var cur = currentStatus();
    var last = state.log.filter(function (e) { return e.s === 'GZD warten'; })[0];
    var std = last ? Math.floor((Date.now() - last.t) / 3600000) : 0;
    if (cur && cur.status === 'GZD warten' && last && std >= 48) {
      hint.hidden = false;
      hint.textContent = '⏰ GZD wartet seit ' + std + ' Std. Antwort vom Kunden prüfen.';
    } else {
      hint.hidden = true;
    }
  }

  // Mailverlauf: nur Nachrichten mit Auftragsinfo, Small Talk wird ausgeblendet
  function renderVerlauf() {
    var th = state.analysis.thread, box = $('verlaufBox');
    var msgs = th.messages;
    if (msgs.length < 2) { box.hidden = true; return; }
    box.hidden = false;
    var rel = msgs.map(function (m, i) { return m.score > 0 || i === th.wichtigste; });
    var versteckt = rel.filter(function (r) { return !r; }).length;
    $('verlaufInfo').textContent = msgs.length + ' Nachrichten' + (versteckt ? ', ' + versteckt + ' ohne Auftragsinfo ausgeblendet' : '');
    var ol = $('verlauf');
    ol.innerHTML = '';
    msgs.forEach(function (m, i) {
      if (!rel[i]) return;
      var li = el('li');
      var h = el('div', 'vhead', (m.von || 'Absender unbekannt') + (m.datum ? ' · ' + m.datum : ''));
      if (i === th.wichtigste) h.appendChild(el('span', 'star', '★ wichtigste Info'));
      li.appendChild(h);
      if (m.punkte.length) {
        var ul = el('ul');
        m.punkte.slice(0, 2).forEach(function (p) { ul.appendChild(el('li', null, p)); });
        li.appendChild(ul);
      }
      ol.appendChild(li);
    });
  }

  function activeMissing() {
    return state.analysis.missing.filter(function (r) {
      return state.detailsInReply || CORE.indexOf(r.key) !== -1;
    });
  }

  function renderReply() {
    var a = state.analysis;
    var miss = activeMissing();
    $('replyCard').hidden = a.missing.length === 0;
    var b = $('replyBtn');
    b.disabled = miss.length === 0;
    b.textContent = miss.length
      ? 'Rückfrage zu ' + miss.length + ' fehlenden Angabe' + (miss.length === 1 ? '' : 'n') + ' schreiben'
      : 'Kernangaben vollständig';
  }

  function renderDetails() {
    var a = state.analysis;
    var ul = $('checklist');
    ul.innerHTML = '';
    var ok = 0, total = 0;
    a.order.forEach(function (key) {
      var f = a.fields[key];
      var li = el('li', (f.state === 'fehlt' ? 'nichts' : f.state) + (f.required ? '' : ' optional'));
      var icon = f.state === 'ok' ? '✓' : (f.state === 'fehlt' ? '' : (f.state === 'folgt' ? '⏳' : '✗'));
      li.appendChild(el('div', 'icon', icon));
      var main = el('div');
      main.appendChild(el('div', 'flabel', f.label + (f.required ? '' : ' (Zusatz)')));
      var vals = el('div', 'fvals');
      f.values.slice(0, 5).forEach(function (v) {
        var s = el('span', null, v.text);
        s.title = v.context || '';
        vals.appendChild(s);
      });
      main.appendChild(vals);
      var hint = hintFor(f);
      if (hint) main.appendChild(el('div', 'fhint', hint));
      li.appendChild(main);
      ul.appendChild(li);
      if (f.required) { total++; if (f.state === 'ok') ok++; }
    });
    $('counter').textContent = ok + ' von ' + total + ' vollständig';

    // Positionen
    var pos = a.positionen || [];
    $('posBox').hidden = pos.length === 0;
    if (pos.length) {
      var unv = pos.filter(function (p) { return p.fehlt.length; }).length;
      $('posCount').textContent = pos.length + ' Positionen' + (unv ? ', ' + unv + ' unvollständig' : '');
      var pl = $('positionen');
      pl.innerHTML = '';
      pos.forEach(function (p, i) {
        var li = el('li');
        li.appendChild(el('div', 'pos-title', (i + 1) + '. ' + p.text));
        var row = el('div');
        [['Menge', p.menge], ['Format', p.format], ['Papier', p.papier], ['Farbe', p.farbe]].forEach(function (c) {
          row.appendChild(el('span', 'pos-cell ' + (c[1] ? 'ok' : 'fehlt'), c[0] + ': ' + (c[1] || 'fehlt')));
        });
        li.appendChild(row);
        pl.appendChild(li);
      });
    }

    // Wichtige Stellen
    $('punkteBox').hidden = a.punkte.length === 0;
    var pu = $('punkte');
    pu.innerHTML = '';
    a.punkte.forEach(function (p) { pu.appendChild(el('li', null, p)); });
  }

  function hintFor(f) {
    if (f.state === 'fehlt') return f.required ? 'Nicht in der Mail gefunden' : '';
    if (f.state === 'folgt') return 'Kunde schreibt, dass die Daten noch folgen';
    if (f.state === 'unklar') {
      if (f.key === 'termin') return f.eilig ? 'Nur "dringend", kein Datum' : 'Kein genaues Datum erkannt';
      if (f.key === 'daten') return 'Anhang erwähnt, aber keine Datei erkannt';
      return 'Nur ungenau erwähnt, bitte prüfen';
    }
    return '';
  }

  function note(msg, kind) {
    var n = $('note');
    n.textContent = msg;
    n.className = 'note' + (kind ? ' ' + kind : '');
    n.hidden = !msg;
  }

  // ---------- Status setzen ----------
  function presetColor(k) {
    try { return Office.MailboxEnums.CategoryColor['Preset' + person(k).preset]; } catch (e) { return undefined; }
  }

  function setStatus(status) {
    if (!state.item) return;
    var p = state.person || state.me || state.team[0];
    var actor = state.me || p;
    var cat = p + SEP + status;

    if (state.demo) {
      state.categories = state.categories.filter(function (c) { return !STATUS_RE.test(c); }).concat([cat]);
      state.demoCats[state.itemKey] = state.categories;
      addLog(actor, status, p);
      afterSet(status);
      return;
    }

    ensureMaster(cat, p, function () {
      var old = state.categories.filter(function (c) { return STATUS_RE.test(c) && c !== cat; });
      var add = function () {
        state.item.categories.addAsync([cat], function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            note('Konnte Kategorie nicht setzen: ' + (r.error && r.error.message), 'err');
            return;
          }
          addLog(actor, status, p);
          readCategories(function () { afterSet(status); });
        });
      };
      if (old.length) state.item.categories.removeAsync(old, function () { add(); });
      else add();
    });
  }

  function removeStatus() {
    if (!state.item) return;
    var old = state.categories.filter(function (c) { return STATUS_RE.test(c); });
    if (!old.length) return;
    var actor = state.me || state.person;
    if (state.demo) {
      state.categories = state.categories.filter(function (c) { return !STATUS_RE.test(c); });
      state.demoCats[state.itemKey] = state.categories;
      addLog(actor, 'x', '');
      afterSet('x');
      return;
    }
    state.item.categories.removeAsync(old, function () {
      addLog(actor, 'x', '');
      readCategories(function () { afterSet('x'); });
    });
  }

  function ensureMaster(cat, personKuerzel, done) {
    var mc = Office.context.mailbox.masterCategories;
    if (!mc) { done(); return; }
    mc.getAsync(function (r) {
      var exists = r.status === Office.AsyncResultStatus.Succeeded &&
        (r.value || []).some(function (c) { return c.displayName === cat; });
      if (exists) { done(); return; }
      var entry = { displayName: cat };
      var col = presetColor(personKuerzel);
      if (col) entry.color = col;
      mc.addAsync([entry], function () { done(); });
    });
  }

  function afterSet(status) {
    syncPersonFromCategories();
    renderKern();
    renderStatus();
    renderLog();
    if (status === 'x') note('Status entfernt.', 'good');
    else if (status === 'Erledigt') note('Als erledigt markiert. Der Ablage-Flow verschiebt die Mail später in den Ordner "Erledigt".', 'good');
    else note('Status gesetzt: ' + state.person + ' · ' + status, 'good');
    setTimeout(function () { note(''); }, 4000);
  }

  // ---------- Rückfrage ----------
  function replyData() {
    var a = state.analysis;
    return { fields: a.fields, missing: activeMissing(), thread: a.thread };
  }

  function writeReply() {
    if (!state.analysis) return;
    var data = replyData();
    if (!data.missing.length) return;
    if (state.demo) { alert(Auftrag.rueckfrage(data)); return; }
    var item = state.item;
    var an = (item.from && item.from.emailAddress) || '';
    var betreff = item.subject || '';
    if (!/^\s*(AW|RE)\s*:/i.test(betreff)) betreff = 'AW: ' + betreff.replace(/^\s*(WG|FW|Fwd)\s*:\s*/i, '');
    var form = { subject: betreff, htmlBody: Auftrag.rueckfrageHtml(data) };
    if (an) form.toRecipients = [an];
    try {
      Office.context.mailbox.displayNewMessageForm(form);
    } catch (e) {
      kopieren();
    }
  }

  function kopieren() {
    if (!state.analysis) return;
    var txt = Auftrag.rueckfrage(replyData());
    var ta = document.createElement('textarea');
    ta.value = txt;
    ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { }
    document.body.removeChild(ta);
    if (ok) {
      note('Rückfragetext kopiert. In eine Antwort einfügen mit Strg+V.', 'good');
    } else {
      $('copyBox').hidden = false;
      $('copyText').value = txt;
      $('copyText').select();
      note('Automatisches Kopieren geht hier nicht. Text unten markieren und mit Strg+C kopieren.', 'err');
    }
  }

  // ---------- Übersicht ----------
  // Datenquelle: im Demo-Modus die Testmails. In Outlook braucht die Mail-Liste Microsoft Graph
  // (noch nicht eingerichtet), dort erscheint ein Hinweis.
  function buildDemoOverview() {
    var tm = window.TESTMAILS || [];
    var fake = [null, { p: 'THe', s: 'In Bearbeitung' }, { p: 'SRu', s: 'Daten folgen' }, null,
      { p: 'FSt', s: 'GZD warten' }, { p: 'JRe', s: 'Warten' }, { p: 'THe', s: 'Erledigt' }, { p: 'SRu', s: 'In Bearbeitung' }];
    state.overview.items = tm.map(function (m, i) {
      var an = Auftrag.analyse(m);
      return {
        idx: i,
        subject: m.subject || '(ohne Betreff)',
        from: m.from || '',
        time: Date.now() - i * 5400000,
        kern: an.kern,
        urgent: an.kern.bisWann.eilig,
        risiko: an.risiken.some(function (r) { return r.typ === 'anhang_fehlt'; }),
        status: fake[i % fake.length]
      };
    });
  }

  function countBy(items, fn) {
    var c = {};
    items.forEach(function (it) { var k = fn(it); c[k] = (c[k] || 0) + 1; });
    return c;
  }

  function chipRow(box, entries, current, onPick, colorFn) {
    box.innerHTML = '';
    entries.forEach(function (e) {
      var b = el('button', 'chip' + (current === e.key ? ' on' : ''), e.label);
      b.type = 'button';
      var col = colorFn ? colorFn(e.key) : '';
      if (current === e.key && col) { b.style.borderColor = col; b.style.color = col; }
      else if (col) { b.style.color = col; }
      b.onclick = function () { onPick(e.key); };
      box.appendChild(b);
    });
  }

  function renderOverview() {
    var o = state.overview;
    var all = o.items;
    var cMa = countBy(all, function (it) { return it.status ? it.status.p : 'ohne'; });
    var cSt = countBy(all, function (it) { return it.status ? it.status.s : 'ohne'; });

    var maEntries = [{ key: 'alle', label: 'Alle ' + all.length }]
      .concat(state.team.map(function (k) { return { key: k, label: person(k).name + ' ' + (cMa[k] || 0) }; }))
      .concat([{ key: 'ohne', label: 'Ohne ' + (cMa.ohne || 0) }]);
    chipRow($('overMa'), maEntries, o.ma, function (k) { o.ma = k; renderOverview(); },
      function (k) { return (k === 'alle' || k === 'ohne') ? '' : person(k).farbe; });

    var stEntries = [{ key: 'alle', label: 'Alle' }]
      .concat(STATUS.map(function (s) { return { key: s, label: s + ' ' + (cSt[s] || 0) }; }))
      .concat([{ key: 'ohne', label: 'Ohne ' + (cSt.ohne || 0) }]);
    chipRow($('overSt'), stEntries, o.st, function (k) { o.st = k; renderOverview(); },
      function (k) { return STATUS_FARBE[k] || ''; });

    $('sortWichtig').className = 'chip' + (o.sort === 'wichtig' ? ' on' : '');
    $('sortNeu').className = 'chip' + (o.sort === 'neu' ? ' on' : '');

    var q = o.q.trim().toLowerCase();
    var list = all.filter(function (it) {
      if (o.ma === 'ohne' && it.status) return false;
      if (o.ma !== 'alle' && o.ma !== 'ohne' && !(it.status && it.status.p === o.ma)) return false;
      if (o.st === 'ohne' && it.status) return false;
      if (o.st !== 'alle' && o.st !== 'ohne' && !(it.status && it.status.s === o.st)) return false;
      if (q) {
        var hay = (it.subject + ' ' + it.from + ' ' + it.kern.was.text + ' ' + it.kern.bisWann.text).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });

    var score = function (it) {
      if (it.status && it.status.s === 'Erledigt') return -1;
      return (it.urgent ? 4 : 0) + (it.risiko ? 3 : 0) + (it.kern.bisWann.state === 'fehlt' ? 1 : 0);
    };
    list.sort(function (a, b) {
      if (o.sort === 'wichtig') { var d = score(b) - score(a); if (d) return d; }
      return b.time - a.time;
    });

    $('overCount').textContent = list.length + ' von ' + all.length + ' Mails';
    var ul = $('overList');
    ul.innerHTML = '';
    list.forEach(function (it) {
      var li = el('li', 'orow');
      li.style.borderLeftColor = it.status ? person(it.status.p).farbe : '';
      var r1 = el('div', 'r1');
      r1.appendChild(el('span', 'subj', it.subject));
      r1.appendChild(el('span', 'when', fmtZeit(it.time)));
      li.appendChild(r1);

      var was = it.kern.was.text || 'nicht erkennbar';
      var r2 = el('div', 'r2', was);
      it.kern.was.fehlt.forEach(function (f) { r2.appendChild(el('span', 'bad', ' · ' + (/unvollständig/.test(f) ? f : f + ' fehlt'))); });
      li.appendChild(r2);

      var r3 = el('div', 'r3');
      if (it.urgent) r3.appendChild(el('span', 'eil', 'Eilig'));
      var w = it.kern.bisWann;
      r3.appendChild(el('span', w.state === 'fehlt' ? 'bad' : (w.state === 'unklar' ? 'okk' : 'okk'),
        'Bis: ' + (w.state === 'fehlt' ? 'fehlt' : w.text)));
      var an = it.kern.anhang.state;
      var anOk = an === 'ja' || an === 'geliefert';
      r3.appendChild(el('span', anOk ? 'okk' : 'bad',
        'Anhang: ' + (anOk ? 'ja' : (an === 'folgt' ? 'folgt' : (an === 'teilweise' ? 'teilweise' : 'nein')))));
      if (it.status) {
        var st = el('span', null, it.status.p + ' · ' + it.status.s);
        st.style.color = STATUS_FARBE[it.status.s];
        st.style.fontWeight = '600';
        r3.appendChild(st);
      }
      li.appendChild(r3);

      li.onclick = function () { openFromOverview(it); };
      ul.appendChild(li);
    });
  }

  function openFromOverview(it) {
    if (state.demo) {
      runDemo(window.TESTMAILS[it.idx]);
      showTab('mail');
    } else if (it.id && Office.context.mailbox.displayMessageForm) {
      Office.context.mailbox.displayMessageForm(it.id);
    }
  }

  function showTab(which) {
    var mail = which === 'mail';
    $('viewMail').hidden = !mail;
    $('viewOver').hidden = mail;
    $('tabMail').className = 'tab' + (mail ? ' on' : '');
    $('tabOver').className = 'tab' + (mail ? '' : ' on');
    if (!mail) renderOverview();
  }

  function setupOverview() {
    var notice = $('overNotice');
    if (state.demo) {
      buildDemoOverview();
      notice.hidden = false;
      notice.textContent = 'Vorschau mit Testdaten. In Outlook zeigt dieser Reiter später alle Mails des Postfachs.';
      $('overBar').hidden = false;
    } else {
      notice.hidden = false;
      notice.textContent = 'Die Übersicht über alle Mails braucht eine Verbindung zu Microsoft 365 (Graph) und ist noch nicht eingerichtet. Bis dahin: Kategorien in der linken Leiste anklicken, z. B. "SRu · Warten".';
      $('overBar').hidden = true;
      $('overList').innerHTML = '';
    }
  }

  // ---------- Start ----------
  window.AuftragUI = { htmlZuText: htmlZuText, parseLog: parseLog, serializeLog: serializeLog }; // für Tests

  function bind() {
    $('saveSettings').onclick = saveSettings;
    $('replyBtn').onclick = writeReply;
    $('copyBtn').onclick = kopieren;
    $('noStatus').onclick = removeStatus;
    $('notizSave').onclick = saveNotiz;
    $('tabMail').onclick = function () { showTab('mail'); };
    $('tabOver').onclick = function () { showTab('over'); };
    $('detailsInReply').onchange = function () { state.detailsInReply = this.checked; if (state.analysis) renderReply(); };
    $('overSearch').oninput = function () { state.overview.q = this.value; renderOverview(); };
    $('sortWichtig').onclick = function () { state.overview.sort = 'wichtig'; renderOverview(); };
    $('sortNeu').onclick = function () { state.overview.sort = 'neu'; renderOverview(); };
  }

  function runDemo(mail) {
    state.item = { demo: true };
    state.itemKey = 'demo:' + (mail.subject || '') + ':' + (mail.text || '').length;
    state.subject = mail.subject || '';
    state.categories = state.demoCats[state.itemKey] || [];
    state.log = state.demoLogs[state.itemKey] || [];
    state.analysis = Auftrag.analyse(mail);
    syncPersonFromCategories();
    note('');
    render();
  }

  function startDemo() {
    state.demo = true;
    $('demo').hidden = false;
    loadSettings();
    var sel = $('demoSelect');
    (window.TESTMAILS || []).forEach(function (t, i) {
      var o = el('option', null, t.name);
      o.value = i;
      sel.appendChild(o);
    });
    sel.onchange = function () { runDemo(window.TESTMAILS[sel.value]); };
    $('demoGo').onclick = function () {
      var t = $('demoText').value.trim();
      if (t) runDemo({ subject: 'Eigene Testmail', text: t, attachments: [] });
      else runDemo(window.TESTMAILS[sel.value]);
    };
    setupOverview();
    if (window.TESTMAILS && window.TESTMAILS.length) runDemo(window.TESTMAILS[0]);
  }

  function start() {
    bind();
    var haveMailbox = typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.item !== undefined;
    if (haveMailbox) {
      loadSettings();
      setupOverview();
      try {
        Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { note(''); loadItem(); });
      } catch (e) { /* ältere Clients */ }
      loadItem();
    } else {
      startDemo();
    }
  }

  if (typeof Office !== 'undefined' && Office.onReady) {
    var started = false;
    var go = function () { if (!started) { started = true; start(); } };
    Office.onReady(go);
    setTimeout(go, 2500);
  } else {
    window.addEventListener('load', start);
  }
})();
