/* Läuft automatisch beim Senden einer Mail (Ereignis OnMessageSend).
 * Erkennt "Gut zum Druck" (GzD) und setzt auf der gesendeten Mail die Kategorie
 * "<Kürzel> · GZD warten". Die Mail wird nie blockiert oder verändert, nur kategorisiert. */

var GZD_TEXT = /(gut\s+zum\s+druck|\bgzd\b|g\.\s?z\.\s?d\.|druckfreigabe|korrekturabzug)/i;
var GZD_DATEI = /(gzd|gut[\s_\-]?zum[\s_\-]?druck|korrektur|freigabe)/i;
var TEAM_PRESET = { SRu: 8, THe: 7, FSt: 4, JRe: 1 };

function onMessageSendHandler(event) {
  try {
    var item = Office.context.mailbox.item;
    item.subject.getAsync(function (sr) {
      var subject = (sr.status === Office.AsyncResultStatus.Succeeded && sr.value) || '';
      item.body.getAsync(Office.CoercionType.Text, function (br) {
        var body = (br.status === Office.AsyncResultStatus.Succeeded && br.value) || '';
        // Nur den neu geschriebenen Teil prüfen, nicht den zitierten Verlauf darunter
        var neu = body.split(/\r?\n[ \t]*(?:Von|From)\s*:/i)[0].slice(0, 3000);
        item.getAttachmentsAsync(function (ar) {
          var names = (ar.status === Office.AsyncResultStatus.Succeeded && ar.value || []).map(function (a) { return a.name; });
          var istGzd = GZD_TEXT.test(subject) || GZD_TEXT.test(neu) ||
            names.some(function (n) { return GZD_DATEI.test(n); });
          if (!istGzd) { event.completed({ allowEvent: true }); return; }
          kategorieSetzen(item, function () { event.completed({ allowEvent: true }); });
        });
      });
    });
  } catch (e) {
    // Egal was passiert: die Mail muss rausgehen
    event.completed({ allowEvent: true });
  }
}

function kategorieSetzen(item, done) {
  var me = '';
  try { me = Office.context.roamingSettings.get('me') || ''; } catch (e) { }
  var cat = (me ? me + ' · ' : '') + '4 GZD warten';
  var mc = Office.context.mailbox.masterCategories;
  var add = function () {
    item.categories.addAsync([cat], function () { done(); });
  };
  if (!mc) { add(); return; }
  mc.getAsync(function (r) {
    var exists = r.status === Office.AsyncResultStatus.Succeeded &&
      (r.value || []).some(function (c) { return c.displayName === cat; });
    if (exists) { add(); return; }
    var entry = { displayName: cat };
    try {
      var p = TEAM_PRESET[me];
      if (p !== undefined) entry.color = Office.MailboxEnums.CategoryColor['Preset' + p];
    } catch (e) { }
    mc.addAsync([entry], function () { add(); });
  });
}

Office.onReady(function () { });
Office.actions.associate('onMessageSendHandler', onMessageSendHandler);
