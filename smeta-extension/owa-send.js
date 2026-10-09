// «📨 Отправить» — письмо уходит прямо из панели, с вашей рабочей почты (окажется в «Отправленных»).
// Делаем ровно то, что делает сам Outlook Web App (Exchange 2019): черновик (CreateItem) → вложения
// (CreateAttachmentFromLocalFile) → отправка (UpdateItem с SendAndSaveCopy). Запросы идут со страницы Outlook,
// от имени того, кто в ней вошёл; нет открытой вкладки — откроем в фоне.
import { OWA_ORIGIN } from './outlook.js';
import { fixBold } from './richtext.js';

function waitComplete(tabId, timeoutMs = 30000) {
  return new Promise((resolve) => {
    const timer = setTimeout(done, timeoutMs);
    function done() { clearTimeout(timer); chrome.tabs.onUpdated.removeListener(on); resolve(); }
    function on(id, info) { if (id === tabId && info.status === 'complete') done(); }
    chrome.tabs.onUpdated.addListener(on);
    chrome.tabs.get(tabId).then((t) => { if (t.status === 'complete') done(); }).catch(done);
  });
}

async function owaTab() {
  const tabs = await chrome.tabs.query({ url: `${OWA_ORIGIN}/owa/*` });
  const ready = tabs.find((t) => t.status === 'complete') || tabs[0];
  if (ready) { if (ready.status !== 'complete') await waitComplete(ready.id); return ready; }
  const tab = await chrome.tabs.create({ url: `${OWA_ORIGIN}/owa/`, active: false });
  await waitComplete(tab.id);
  return tab;
}

const escHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/** Текст письма → HTML как у Outlook: шрифт Calibri 12pt, абзацы, ссылки кликабельные, **жирный**. */
export function textToHtml(text) {
  const lines = fixBold(text).split('\n').map((line) => {
    const html = escHtml(line)
      .replace(/(https?:\/\/[^\s<*]+)/g, (u) => `<a href="${u}">${u}</a>`)
      .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>'); // **жирный**
    return `<p>${html || '<br>'}</p>`;
  }).join('');
  return '<div style="font-size:12pt;color:#000000;font-family:Calibri,Helvetica,sans-serif;" dir="ltr">' + lines + '<p><br></p></div>';
}

/**
 * Отправить письмо. msg: { to: [], cc: [], subject, text, files: [File], signatureHtml }.
 * Возвращает { ok: true } или бросает ошибку с понятным текстом.
 */
export async function sendMail(msg) {
  const files = await Promise.all((msg.files || []).map(async (f) => {
    const bytes = new Uint8Array(await f.arrayBuffer());
    let s = '';
    for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return { name: f.name, type: f.type || 'application/octet-stream', size: f.size, b64: btoa(s) };
  }));
  const tab = await owaTab();
  let res;
  try {
    [res] = await chrome.scripting.executeScript({
      target: { tabId: tab.id }, world: 'MAIN', func: owaSendInPage,
      args: [{ to: msg.to, cc: msg.cc || [], subject: msg.subject, html: textToHtml(msg.text), files, signatureHtml: msg.signatureHtml || '' }],
    });
  } catch (e) {
    throw new Error(/error page|Cannot access/i.test(e.message) ? 'Outlook не открылся — проверьте сеть/VPN и вход в почту' : e.message);
  }
  const r = res && res.result;
  if (!r) throw new Error('Outlook не ответил');
  if (r.error) throw new Error(r.error);
  return r;
}

/** Выполняется на странице Outlook Web App (mail.rwb.ru/owa): там вход и X-OWA-CANARY. */
async function owaSendInPage(msg) {
  const canary = (document.cookie.match(/(?:^|;\s*)X-OWA-CANARY=([^;]*)/i) || [])[1];
  if (!canary) return { error: 'Не вижу входа в Outlook — откройте почту, войдите и попробуйте ещё раз' };
  let seq = -900;
  const header = (ver) => ({
    __type: 'JsonRequestHeaders:#Exchange', RequestServerVersion: ver,
    TimeZoneContext: { __type: 'TimeZoneContext:#Exchange', TimeZoneDefinition: { __type: 'TimeZoneDefinitionType:#Exchange', Id: 'Russian Standard Time' } },
  });
  async function call(action, body, actionName = action) {
    const id = seq--;
    const r = await fetch(`/owa/service.svc?action=${action}&ID=${id}&AC=1`, {
      method: 'POST',
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json; charset=utf-8', Action: action, 'X-OWA-CANARY': decodeURIComponent(canary),
        'X-OWA-ActionName': actionName, 'X-OWA-ActionId': String(id), 'X-Requested-With': 'XMLHttpRequest',
      },
      body: JSON.stringify(body),
    });
    if (r.status === 440 || r.status === 401) throw new Error('Outlook просит войти заново — откройте почту и войдите');
    let j = null;
    try { j = await r.json(); } catch { /* ниже скажем код */ }
    const m = j && j.Body && j.Body.ResponseMessages && j.Body.ResponseMessages.Items && j.Body.ResponseMessages.Items[0];
    if (!r.ok || !m) throw new Error(`Outlook не принял письмо (${action}, код ${r.status})`);
    if (m.ResponseClass !== 'Success') throw new Error(`Outlook: ${m.MessageText || m.ResponseCode}`);
    return m;
  }
  const rcpt = (list) => list.map((e) => ({ Name: e, EmailAddress: e, RoutingType: 'SMTP', MailboxType: 'OneOff' }));

  // Подпись — сохранённая в расширении; нет — та же, что Outlook ставит в новые письма (если найдём)
  let signature = msg.signatureHtml ? `<div id="Signature">${msg.signatureHtml}</div>` : '';
  if (!signature) try {
    const r = await fetch('/owa/service.svc?action=GetOwaUserConfiguration&AC=1', {
      method: 'POST', credentials: 'include',
      headers: { 'Content-Type': 'application/json; charset=utf-8', Action: 'GetOwaUserConfiguration', 'X-OWA-CANARY': decodeURIComponent(canary) },
      body: '{}',
    });
    const j = await r.json();
    const o = (j && (j.UserOptions || (j.Body && j.Body.UserOptions))) || {};
    if (o.SignatureHtml && o.AutoAddSignature !== false) signature = `<div id="Signature">${o.SignatureHtml}</div>`;
  } catch { /* без подписи */ }

  const html = '<html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><style type="text/css" style="display:none;"><!-- P {margin-top:0;margin-bottom:0;} --></style></head>' +
    `<body dir="ltr"><div id="divtagdefaultwrapper" style="font-size:12pt;color:#000000;font-family:Calibri,Helvetica,sans-serif;" dir="ltr">${msg.html}${signature}</div></body></html>`;

  try {
    // 1. Черновик
    const created = await call('CreateItem', {
      __type: 'CreateItemJsonRequest:#Exchange', Header: header('V2015_10_15'),
      Body: {
        __type: 'CreateItemRequest:#Exchange',
        Items: [{
          __type: 'Message:#Exchange', Subject: msg.subject,
          Body: { __type: 'BodyContentType:#Exchange', BodyType: 'HTML', Value: html },
          Importance: 'Normal', From: null, ToRecipients: rcpt(msg.to), CcRecipients: rcpt(msg.cc), BccRecipients: [],
          Sensitivity: 'Normal', IsDeliveryReceiptRequested: false, IsReadReceiptRequested: false,
        }],
        ClientSupportsIrm: true, OutboundCharset: 'AutoDetect', MessageDisposition: 'SaveOnly', ComposeOperation: 'newMail',
      },
    });
    const item = created.Items && created.Items[0] && created.Items[0].ItemId;
    if (!item || !item.Id) throw new Error('Outlook не создал черновик');
    let changeKey = item.ChangeKey;

    // 2. Вложения — по одному, у черновика каждый раз новый ChangeKey
    for (const f of msg.files) {
      const att = await call('CreateAttachmentFromLocalFile', {
        __type: 'CreateAttachmentJsonRequest:#Exchange', Header: header('Exchange2013'),
        Body: {
          __type: 'CreateAttachmentRequest:#Exchange',
          ParentItemId: { __type: 'ItemId:#Exchange', Id: item.Id, ChangeKey: changeKey },
          Attachments: [{ __type: 'FileAttachment:#Exchange', Content: f.b64, IsContactPhoto: false, ContentType: f.type, IsInline: false, Name: f.name, Size: f.size }],
          RequireImageType: false, IncludeContentIdInResponse: false, ClientSupportsIrm: true, CancellationId: null,
        },
      });
      const a = att.Attachments && att.Attachments[0] && att.Attachments[0].AttachmentId;
      if (a && a.RootItemChangeKey) changeKey = a.RootItemChangeKey;
    }

    // 3. Отправить и сохранить копию в «Отправленных»
    await call('UpdateItem', {
      __type: 'UpdateItemJsonRequest:#Exchange', Header: header('Exchange2015'),
      Body: {
        __type: 'UpdateItemRequest:#Exchange',
        ItemChanges: [{
          __type: 'ItemChange:#Exchange',
          Updates: [{ __type: 'SetItemField:#Exchange', Path: { __type: 'PropertyUri:#Exchange', FieldURI: 'Subject' }, Item: { __type: 'Message:#Exchange', Subject: msg.subject } }],
          ItemId: { __type: 'ItemId:#Exchange', Id: item.Id, ChangeKey: changeKey },
        }],
        ConflictResolution: 'AlwaysOverwrite', ClientSupportsIrm: true, SendCalendarInvitationsOrCancellations: 'SendToNone',
        MessageDisposition: 'SendAndSaveCopy', SuppressReadReceipts: false, ComposeOperation: 'newMail',
        OutboundCharset: 'AutoDetect', PromoteInlineAttachments: false, SendOnNotFoundError: true,
      },
    }, 'UpdateMessageForComposeSend'); // так Outlook называет «Отправить»
    return { ok: true, signature: Boolean(signature) };
  } catch (e) {
    return { error: e.message };
  }
}
