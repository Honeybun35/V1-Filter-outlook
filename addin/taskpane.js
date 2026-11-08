/* Seitenleiste "Auftrags-Check" für Outlook (Office.js, Mailbox 1.8) */
(function () {
  'use strict';

  var STATUS = ['Warten', 'Daten folgen', 'In Bearbeitung', 'GZD warten', 'Erledigt'];
  var SEP = ' · ';
  var STATUS_RE = new RegExp('^(.+?)' + SEP.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(' + STATUS.join('|') + ')$');
  var TEAM = window.TEAM || [];
  var DEFAULT_TEAM = TEAM.map(function (t) { return t.kuerzel; });

  function person(k) {
    for (var i = 0; i < TEAM.length; i++) if (TEAM[i].kuerzel === k) return TEAM[i];
    return { kuerzel: k, name: k, farbe: '#6b7280', preset: 12 };
  }

  var state = {
    demo: false,
    item: null,
    analysis: null,
    categories: [],      // Anzeigenamen der Kategorien der aktuellen Mail
    me: '',
    team: DEFAULT_TEAM.slice(),
    person: '',
    na: {},              // pro Mail: Felder, die als "nicht nötig" markiert sind
    itemKey: ''
  };

  var $ = function (id) { return document.getElementById(id); };

  // ---------- Einstellungen ----------
  function loadSettings() {
    var me = '';
    try {
      if (!state.demo && Office.context.roamingSettings) {
        me = Office.context.roamingSettings.get('me') || '';
      } else {
        me = localStorage.getItem('ac_me') || '';
      }
    } catch (e) { /* ignorieren */ }
    state.me = me;
    state.team = DEFAULT_TEAM.slice();
    state.person = me || state.team[0];
    var sel = $('meSelect');
    sel.innerHTML = '<option value="">– bitte wählen –</option>';
    TEAM.forEach(function (t) {
      var o = document.createElement('option');
      o.value = t.kuerzel; o.textContent = t.name + ' (' + t.kuerzel + ')';
      sel.appendChild(o);
    });
    sel.value = me;
  }

  function saveSettings() {
    var me = $('meSelect').value;
    state.me = me;
    state.person = me || state.team[0];
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
    syncPersonFromCategories();
    render();
  }

  // ---------- Mail laden ----------
  function loadItem() {
    var item = Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { showEmpty(); return; }
    state.item = item;
    state.itemKey = item.itemId || item.conversationId || item.subject || '';
    var subject = item.subject || '';
    var attachments = (item.attachments || []).map(function (a) { return a.name; });

    item.body.getAsync(Office.CoercionType.Text, function (res) {
      var text = (res.status === Office.AsyncResultStatus.Succeeded) ? res.value : '';
      readCategories(function () {
        var from = (item.from && (item.from.displayName || item.from.emailAddress)) || '';
        state.analysis = Auftrag.analyse({ subject: subject, text: text, attachments: attachments, from: from });
        state.subject = subject;
        syncPersonFromCategories();
        render();
      });
    });
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

  function showEmpty() {
    state.item = null; state.analysis = null;
    $('subject').textContent = 'Keine Mail geöffnet';
    $('checklist').innerHTML = '';
    $('counter').textContent = '';
    $('replyBtn').hidden = true; $('replyHint').hidden = true;
    $('punkteBox').hidden = true; $('urgent').hidden = true;
    $('risks').innerHTML = ''; $('verlaufBox').hidden = true; $('posBox').hidden = true;
    renderStatus();
  }

  // ---------- Anzeige ----------
  function render() {
    var a = state.analysis;
    if (!a) { showEmpty(); return; }
    $('subject').textContent = state.subject || '(ohne Betreff)';
    $('urgent').hidden = !a.urgent;
    renderRisks();
    renderStatus();
    renderChecklist();
    renderPunkte();
    renderPositionen();
    renderVerlauf();
  }

  // Einzelne Druckarbeiten der Mail, je mit Menge / Format / Papier / Farbe
  function renderPositionen() {
    var pos = state.analysis.positionen || [];
    var box = $('posBox');
    box.hidden = pos.length === 0;
    if (!pos.length) return;
    var unvollstaendig = pos.filter(function (p) { return p.fehlt.length; }).length;
    $('posCount').textContent = pos.length + ' Positionen' + (unvollstaendig ? ', ' + unvollstaendig + ' unvollständig' : '');
    var ul = $('positionen');
    ul.innerHTML = '';
    pos.forEach(function (p, i) {
      var li = document.createElement('li');
      var h = document.createElement('div');
      h.className = 'pos-title';
      h.textContent = (i + 1) + '. ' + p.text;
      li.appendChild(h);
      var row = document.createElement('div');
      row.className = 'pos-row';
      [['Menge', p.menge], ['Format', p.format], ['Papier', p.papier], ['Farbe', p.farbe]].forEach(function (c) {
        var s = document.createElement('span');
        s.className = 'pos-cell ' + (c[1] ? 'ok' : 'fehlt');
        s.textContent = c[0] + ': ' + (c[1] || 'fehlt');
        row.appendChild(s);
      });
      li.appendChild(row);
      ul.appendChild(li);
    });
  }

  // Roter Warnblock: fehlende Anhänge, Weiterleitung ohne Anhang
  function renderRisks() {
    var box = $('risks');
    box.innerHTML = '';
    (state.analysis.risiken || []).forEach(function (r) {
      var d = document.createElement('div');
      d.className = 'risk' + (r.typ !== 'anhang_fehlt' ? ' mild' : '');
      var t = document.createElement('div');
      t.className = 'risk-title';
      t.textContent = (r.typ !== 'anhang_fehlt' ? '⚠ ' : '⛔ ') + r.titel;
      d.appendChild(t);
      if (r.details && r.details.length) {
        var ul = document.createElement('ul');
        r.details.forEach(function (n) {
          var li = document.createElement('li'); li.textContent = n; ul.appendChild(li);
        });
        d.appendChild(ul);
      }
      box.appendChild(d);
    });
  }

  // Zusammenfassung des Mailverlaufs
  function renderVerlauf() {
    var th = state.analysis.thread;
    var box = $('verlaufBox');
    var msgs = th && th.messages ? th.messages.filter(function (m) { return m.punkte.length || m.von; }) : [];
    if (msgs.length < 2) { box.hidden = true; return; }
    box.hidden = false;
    $('verlaufCount').textContent = msgs.length + ' Nachrichten';
    var ol = $('verlauf');
    ol.innerHTML = '';
    msgs.forEach(function (m, i) {
      var li = document.createElement('li');
      var h = document.createElement('div');
      h.className = 'vhead';
      h.textContent = (i === 0 ? 'Aktuell' : 'Davor') + (m.von ? ' · ' + m.von : '') + (m.datum ? ' · ' + m.datum : '');
      li.appendChild(h);
      if (m.punkte.length) {
        var ul = document.createElement('ul');
        m.punkte.forEach(function (p) {
          var l2 = document.createElement('li'); l2.textContent = p; ul.appendChild(l2);
        });
        li.appendChild(ul);
      }
      ol.appendChild(li);
    });
  }

  function renderStatus() {
    var peopleBox = $('people');
    peopleBox.innerHTML = '';
    state.team.forEach(function (p) {
      var pe = person(p);
      var b = document.createElement('button');
      b.className = 'chip' + (p === state.person ? ' on' : '');
      b.textContent = p;
      b.title = pe.name;
      b.style.borderColor = pe.farbe;
      if (p === state.person) { b.style.background = pe.farbe; b.style.color = '#fff'; }
      b.onclick = function () { state.person = p; renderStatus(); };
      peopleBox.appendChild(b);
    });

    var cur = currentStatus();
    // Farbband: wem gehört diese Mail?
    var band = $('ownerBand');
    if (cur) {
      var owner = person(cur.person);
      band.hidden = false;
      band.style.background = owner.farbe;
      band.textContent = (state.me && cur.person !== state.me ? 'Gehört ' : 'Bearbeiter: ') + owner.name;
    } else {
      band.hidden = true;
    }
    var suggest = state.analysis && state.analysis.vorschlag;
    var box = $('statuses');
    box.innerHTML = '';
    STATUS.forEach(function (s) {
      var b = document.createElement('button');
      b.className = 'status' + (cur && cur.status === s ? ' on' : '') + (!cur && suggest === s ? ' suggest' : '');
      b.setAttribute('data-s', s);
      b.textContent = s;
      b.title = (!cur && suggest === s) ? 'Vorschlag nach Mailtext' : '';
      b.onclick = function () { setStatus(s); };
      box.appendChild(b);
    });
    $('current').textContent = cur ? ('Aktuell: ' + cur.person + ' · ' + cur.status) : 'Noch kein Status gesetzt.';
    if (!state.item) {
      $('current').textContent = 'Öffne eine Mail, um den Status zu setzen.';
    }
  }

  function note(msg, kind) {
    var n = $('note');
    n.textContent = msg;
    n.className = 'note' + (kind ? ' ' + kind : '');
    n.hidden = !msg;
  }

  function activeMissing() {
    var na = state.na[state.itemKey] || {};
    return state.analysis.missing.filter(function (r) { return !na[r.key]; });
  }

  function renderChecklist() {
    var a = state.analysis;
    var na = state.na[state.itemKey] || {};
    var ul = $('checklist');
    ul.innerHTML = '';
    var ok = 0, total = 0;

    a.order.forEach(function (key) {
      var f = a.fields[key];
      var isNa = !!na[key];
      var li = document.createElement('li');
      var cls = f.state + (f.required ? '' : ' optional') + (isNa ? ' na' : '');
      li.className = cls;

      var icon = document.createElement('div');
      icon.className = 'icon';
      icon.textContent = f.state === 'ok' ? '✓' : (f.state === 'fehlt' ? '✗' : (f.state === 'folgt' ? '⏳' : '!'));

      var main = document.createElement('div');
      var lab = document.createElement('div');
      lab.className = 'flabel';
      lab.textContent = f.label + (f.required ? '' : ' (Zusatz)');
      main.appendChild(lab);

      var vals = document.createElement('div');
      vals.className = 'fvals';
      f.values.slice(0, 5).forEach(function (v) {
        var s = document.createElement('span');
        s.textContent = v.text;
        s.title = v.context || '';
        vals.appendChild(s);
      });
      if (f.values.length > 5) {
        var more = document.createElement('span');
        more.textContent = '+' + (f.values.length - 5);
        vals.appendChild(more);
      }
      main.appendChild(vals);

      var hint = hintFor(f);
      if (hint) {
        var h = document.createElement('div');
        h.className = 'fhint';
        h.textContent = hint;
        main.appendChild(h);
      }

      var side = document.createElement('div');
      if (f.required && f.state !== 'ok') {
        var skip = document.createElement('button');
        skip.className = 'skip';
        skip.textContent = isNa ? 'doch nötig' : 'nicht nötig';
        skip.onclick = function () {
          state.na[state.itemKey] = state.na[state.itemKey] || {};
          state.na[state.itemKey][key] = !isNa;
          renderChecklist();
        };
        side.appendChild(skip);
      }

      li.appendChild(icon); li.appendChild(main); li.appendChild(side);
      ul.appendChild(li);

      if (f.required && !isNa) { total++; if (f.state === 'ok') ok++; }
    });

    $('counter').textContent = ok + ' von ' + total + ' vollständig';
    var miss = activeMissing();
    $('replyBtn').hidden = miss.length === 0;
    $('replyHint').hidden = miss.length === 0;
    $('replyBtn').textContent = 'Rückfrage zu ' + miss.length + ' fehlenden Angabe' + (miss.length === 1 ? '' : 'n') + ' schreiben';
  }

  function hintFor(f) {
    if (f.state === 'fehlt') return f.required ? 'Nicht in der Mail gefunden' : '';
    if (f.state === 'folgt') return 'Kunde schreibt, dass die Daten noch folgen';
    if (f.state === 'unklar') {
      if (f.key === 'termin') return f.eilig ? 'Nur "dringend", kein Datum' : 'Kein genaues Datum erkannt';
      if (f.key === 'daten') return 'Anhang erwähnt, aber keine Datei erkannt';
      return 'Nur ungenau erwähnt, bitte prüfen';
    }
    if (f.key === 'termin' && !f.hasUhrzeit) return 'Uhrzeit nicht genannt';
    return '';
  }

  function renderPunkte() {
    var a = state.analysis;
    var ul = $('punkte');
    ul.innerHTML = '';
    a.punkte.forEach(function (p) {
      var li = document.createElement('li');
      li.textContent = p;
      ul.appendChild(li);
    });
    $('punkteBox').hidden = a.punkte.length === 0;
  }

  // ---------- Status setzen ----------
  // Die Kategorie-Farbe in Outlook gehört zur Person (nicht zum Status)
  function presetColor(personKuerzel) {
    try {
      return Office.MailboxEnums.CategoryColor['Preset' + person(personKuerzel).preset];
    } catch (e) { return undefined; }
  }

  function setStatus(status) {
    if (!state.item) return;
    var person = state.person || state.me || state.team[0];
    var cat = person + SEP + status;

    if (state.demo) {
      state.categories = state.categories.filter(function (c) { return !STATUS_RE.test(c); }).concat([cat]);
      afterSet(status);
      return;
    }

    ensureMaster(cat, person, function () {
      var old = state.categories.filter(function (c) { return STATUS_RE.test(c) && c !== cat; });
      var add = function () {
        state.item.categories.addAsync([cat], function (r) {
          if (r.status !== Office.AsyncResultStatus.Succeeded) {
            note('Konnte Kategorie nicht setzen: ' + (r.error && r.error.message), 'err');
            return;
          }
          readCategories(function () { afterSet(status); });
        });
      };
      if (old.length) {
        state.item.categories.removeAsync(old, function () { add(); });
      } else {
        add();
      }
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
    renderStatus();
    if (status === 'Erledigt') {
      note('Als erledigt markiert. Der Ablage-Flow schiebt die Mail beim nächsten Durchlauf in den Ordner "Erledigt".', 'good');
    } else {
      note('Status gesetzt: ' + state.person + ' · ' + status, 'good');
      setTimeout(function () { note(''); }, 3500);
    }
  }

  // ---------- Rückfrage ----------
  function writeReply() {
    var a = state.analysis;
    if (!a) return;
    var miss = activeMissing();
    var copy = { fields: a.fields, missing: miss, thread: a.thread };
    if (state.demo) {
      alert(Auftrag.rueckfrage(copy));
      return;
    }
    state.item.displayReplyForm(Auftrag.rueckfrageHtml(copy));
  }

  // ---------- Start ----------
  function bind() {
    $('saveSettings').onclick = saveSettings;
    $('replyBtn').onclick = writeReply;
  }

  function startDemo() {
    state.demo = true;
    $('demo').hidden = false;
    loadSettings();
    var sel = $('demoSelect');
    (window.TESTMAILS || []).forEach(function (t, i) {
      var o = document.createElement('option');
      o.value = i; o.textContent = t.name;
      sel.appendChild(o);
    });
    function run(mail) {
      state.item = { demo: true };
      state.itemKey = 'demo:' + (mail.subject || '') + ':' + (mail.text || '').length;
      state.subject = mail.subject || '';
      state.categories = [];
      state.analysis = Auftrag.analyse(mail);
      syncPersonFromCategories();
      note('');
      render();
    }
    sel.onchange = function () { run(window.TESTMAILS[sel.value]); };
    $('demoGo').onclick = function () {
      var t = $('demoText').value.trim();
      if (t) run({ subject: 'Eigene Testmail', text: t, attachments: [] });
      else run(window.TESTMAILS[sel.value]);
    };
    if (window.TESTMAILS && window.TESTMAILS.length) run(window.TESTMAILS[0]);
  }

  function start() {
    bind();
    var haveMailbox = typeof Office !== 'undefined' && Office.context && Office.context.mailbox && Office.context.mailbox.item !== undefined;
    if (haveMailbox) {
      loadSettings();
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
    // Falls office.js zwar geladen ist, aber nicht in Outlook läuft
    setTimeout(go, 2500);
  } else {
    window.addEventListener('load', start);
  }
})();
