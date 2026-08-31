/**
 * People Systems Snapshot — Apps Script (bilingual EN / ES)
 * --------------------------------------------------------
 * Drop-in replacement for the current Code.gs. Keeps your existing design,
 * sheet logging, and 5-day follow-up. Adds Spanish.
 *
 * How language is decided:
 *   - The Spanish page (/es/people_systems_snapshot.html) POSTs  lang=es
 *   - The English page POSTs no lang field  ->  treated as "en"
 *   - The language is stored in the sheet (column G) so the follow-up
 *     email also goes out in the right language.
 *
 * SHEET COLUMNS (unchanged except G):
 *   A timestamp | B name | C email | D pattern | E result | F followUpSent | G lang
 *   Existing rows with an empty G are treated as English.
 *
 * Deploy: Apps Script editor ▸ paste over Code.gs ▸ Deploy ▸ Manage deployments
 *         ▸ edit existing ▸ Version: New version ▸ Deploy.  URL stays the same.
 */

// ─────────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var data = e.parameter;
    data.signals = JSON.parse(data.signals || '[]');
    data.lang = normalizeLang(data.lang);

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.appendRow([
      new Date(),
      data.name || '',
      data.email || '',
      data.pattern || '',
      data.result || '',
      '',            // F — followUpSent, set later by sendFollowUps()
      data.lang      // G — lang
    ]);

    if (data.email) {
      sendResultsEmail(data);
    }

    return ContentService
      .createTextOutput('ok')
      .setMimeType(ContentService.MimeType.TEXT);

  } catch(err) {
    return ContentService
      .createTextOutput('error: ' + err.toString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function normalizeLang(v) {
  return (String(v || 'en').toLowerCase().indexOf('es') === 0) ? 'es' : 'en';
}

// "Pattern: The Founder's Burden" / "Patrón: La carga de quien fundó" -> name only
function stripPatternPrefix(tag) {
  return String(tag || '').replace(/^\s*(Pattern|Patrón)\s*:\s*/i, '');
}


// ─────────────────────────────────────────────────────────────────────
// RESULTS EMAIL
// ─────────────────────────────────────────────────────────────────────
var RESULT_COPY = {
  en: {
    subject:  'Your People Systems Snapshot results',
    eyebrow:  'People Systems Snapshot',
    kicker:   'Your Primary Pattern',
    greeting: function (name) { return 'Hi ' + name + ','; },
    signals:  'Signals to watch for',
    priority: "What I'd prioritize first",
    ctaTitle: 'What happens next',
    ctaBody:  'A 30-minute conversation can turn this snapshot into a clear picture of what to prioritize.',
    ctaReach: 'Reply to this email, or reach Daniela directly at'
  },
  es: {
    subject:  'Los resultados de tu People Systems Snapshot',
    eyebrow:  'People Systems Snapshot',
    kicker:   'Tu patrón principal',
    greeting: function (name) { return 'Hola ' + name + ','; },
    signals:  'Señales a las que prestar atención',
    priority: 'Qué priorizaría primero',
    ctaTitle: 'Qué sigue',
    ctaBody:  'Una conversación de 30 minutos puede convertir este snapshot en una imagen clara de qué priorizar.',
    ctaReach: 'Responde a este correo o escribe a Daniela directamente a'
  }
};

function sendResultsEmail(data) {
  var lang = normalizeLang(data.lang);
  var t = RESULT_COPY[lang];
  var name = data.name || '';
  var signals = data.signals || [];

  var signalsHtml = '';
  for (var i = 0; i < signals.length; i++) {
    signalsHtml += '<li style="margin-bottom:8px;">&rarr; ' + signals[i] + '</li>';
  }

  var patternName = stripPatternPrefix(data.pattern);

  var top =
    '<div style="font-family:Georgia,serif;max-width:600px;margin:0 auto;color:#333333;">' +
    '<div style="background:#0f4c5c;padding:36px 32px;text-align:center;border-radius:10px 10px 0 0;">' +
    '<p style="color:rgba(255,255,255,0.55);font-family:Helvetica,sans-serif;font-size:11px;font-weight:600;letter-spacing:0.14em;text-transform:uppercase;margin:0 0 10px 0;">' + t.eyebrow + '</p>' +
    '<p style="color:rgba(255,255,255,0.75);font-family:Georgia,serif;font-style:italic;font-size:15px;margin:0 0 16px 0;">' + t.kicker + '</p>' +
    '<h1 style="color:#ffffff;font-size:24px;font-weight:normal;line-height:1.3;margin:0 0 14px 0;">' + patternName + '</h1>' +
    '<p style="color:rgba(255,255,255,0.85);font-family:Helvetica,sans-serif;font-size:15px;line-height:1.6;margin:0;">' + (data.result || '') + '</p>' +
    '</div>';

  var body =
    '<div style="background:#FDF9F4;border:1px solid #e5e0d8;border-top:none;border-radius:0 0 10px 10px;padding:36px 32px;">' +
    '<p style="font-size:14px;color:#6b7280;font-family:Helvetica,sans-serif;margin:0 0 20px 0;">' + t.greeting(name) + '</p>' +
    '<p style="font-size:15px;font-family:Helvetica,sans-serif;color:#333333;line-height:1.75;margin:0 0 28px 0;">' + (data.description || '') + '</p>' +
    '<p style="font-size:11px;font-family:Helvetica,sans-serif;letter-spacing:0.1em;text-transform:uppercase;color:#cb6b4b;margin:0 0 12px 0;">' + t.signals + '</p>' +
    '<ul style="list-style:none;padding:0;margin:0 0 28px 0;font-family:Helvetica,sans-serif;font-size:14px;color:#333333;line-height:1.6;">' + signalsHtml + '</ul>' +
    '<p style="font-size:11px;font-family:Helvetica,sans-serif;letter-spacing:0.1em;text-transform:uppercase;color:#cb6b4b;margin:0 0 12px 0;">' + t.priority + '</p>' +
    '<p style="font-size:15px;font-family:Helvetica,sans-serif;color:#333333;line-height:1.75;margin:0 0 32px 0;">' + (data.help || '') + '</p>';

  var cta =
    '<div style="background:#e8f2f5;border-radius:8px;padding:24px;text-align:center;">' +
    '<p style="font-size:15px;color:#0f4c5c;margin:0 0 8px 0;font-family:Georgia,serif;">' + t.ctaTitle + '</p>' +
    '<p style="font-size:13px;color:#6b7280;font-family:Helvetica,sans-serif;margin:0 0 16px 0;">' + t.ctaBody + '</p>' +
    '<p style="font-size:14px;color:#0f4c5c;font-family:Helvetica,sans-serif;margin:0;">' + t.ctaReach + '<br>' +
    '<a href="mailto:daniela@danielareinoso.co" style="color:#cb6b4b;">daniela@danielareinoso.co</a></p>' +
    '</div>';

  var footer =
    '<p style="font-size:11px;color:#9ca3af;font-family:Helvetica,sans-serif;text-align:center;margin:28px 0 0 0;">Daniela Reinoso &middot; danielareinoso.co</p>' +
    '</div></div>';

  var htmlBody = top + body + cta + footer;

  MailApp.sendEmail({
    to: data.email,
    subject: t.subject,
    htmlBody: htmlBody,
    replyTo: 'daniela@danielareinoso.co',
    name: 'Daniela Reinoso'
  });
}


// ─────────────────────────────────────────────────────────────────────
// 5-DAY FOLLOW-UP
// ─────────────────────────────────────────────────────────────────────
function sendFollowUps() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var now = new Date();

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[0]);
    var name = row[1];
    var email = row[2];
    var followUpSent = row[5];      // column F
    var lang = normalizeLang(row[6]); // column G (blank -> "en")

    if (!email || followUpSent) continue;

    var daysSince = (now - timestamp) / (1000 * 60 * 60 * 24);
    if (daysSince < 5) continue;

    var firstName = String(name || '').split(' ')[0];
    sendFollowUpEmail(email, firstName, lang);
    sheet.getRange(i + 1, 6).setValue('sent');
  }
}

var FOLLOWUP_COPY = {
  en: {
    subject: 'One more thought on your Snapshot',
    lines: function (firstName) {
      return [
        'Hi ' + firstName + ',',
        'I wanted to share one additional thought about your Snapshot result.',
        'The goal is not to "fix HR" all at once.',
        'In most growing teams, the fastest progress comes from finding the one people-system bottleneck creating the most drag right now.',
        'For some organizations, that\'s unclear roles. For others, it\'s inconsistent onboarding, decision-making, manager support, or performance conversations that happen too late.',
        'Your Snapshot is a starting point, not a diagnosis.',
        'If it would be useful, I\'d be happy to help you interpret what your result might mean in your actual context.',
        'You can reply here or reach me directly at <a href="mailto:daniela@danielareinoso.co" style="color:#cb6b4b;">daniela@danielareinoso.co</a>.',
        'Either way, I hope the Snapshot gave you a clearer way to name what may be happening underneath the surface.'
      ];
    }
  },
  es: {
    subject: 'Una idea más sobre tu Snapshot',
    lines: function (firstName) {
      return [
        'Hola ' + firstName + ',',
        'Quería compartir una idea más sobre el resultado de tu Snapshot.',
        'El objetivo no es "arreglar RR. HH." de una sola vez.',
        'En la mayoría de los equipos en crecimiento, el avance más rápido viene de encontrar el único cuello de botella en los sistemas de personas que genera más fricción ahora mismo.',
        'Para algunas organizaciones son los roles poco claros. Para otras, una incorporación inconsistente, la toma de decisiones, el apoyo a las jefaturas o conversaciones de desempeño que llegan demasiado tarde.',
        'Tu Snapshot es un punto de partida, no un diagnóstico.',
        'Si te resulta útil, con gusto te ayudo a interpretar qué podría significar tu resultado en tu contexto real.',
        'Puedes responder aquí o escribirme directamente a <a href="mailto:daniela@danielareinoso.co" style="color:#cb6b4b;">daniela@danielareinoso.co</a>.',
        'En cualquier caso, espero que el Snapshot te haya dado una forma más clara de nombrar lo que puede estar pasando bajo la superficie.'
      ];
    }
  }
};

function sendFollowUpEmail(email, firstName, lang) {
  var t = FOLLOWUP_COPY[normalizeLang(lang)];
  var paras = t.lines(firstName).map(function (line) {
    return '<p>' + line + '</p>';
  }).join('');

  var body =
    '<div style="font-family:Helvetica,sans-serif;max-width:600px;margin:0 auto;color:#333;line-height:1.75;font-size:15px;">' +
    paras +
    '<p style="margin-top:32px;color:#9ca3af;font-size:12px;">Daniela Reinoso &middot; danielareinoso.co</p>' +
    '</div>';

  MailApp.sendEmail({
    to: email,
    subject: t.subject,
    htmlBody: body,
    replyTo: 'daniela@danielareinoso.co',
    name: 'Daniela Reinoso'
  });
}
