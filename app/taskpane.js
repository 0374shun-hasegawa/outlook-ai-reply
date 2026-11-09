/*
 * 作業ウィンドウ（AI返信）。Office.js と Claude API をつなぐ。
 * API キーは roamingSettings から読むだけで、画面にもログにも出さない。
 */
(function () {
  'use strict';

  var Core = window.ReplyCore;
  var Store = window.ReplyStore;
  var DEFAULT_TIMEOUT_MS = 60000;

  function $(id) { return document.getElementById(id); }

  var el = {
    purpose: $('purpose'),
    create: $('create'),
    status: $('status'),
    resultBlock: $('resultBlock'),
    result: $('result'),
    platformNote: $('platformNote'),
    replyAll: $('replyAll'),
    replyAllLabel: $('replyAllLabel'),
    openReply: $('openReply'),
    copy: $('copy')
  };

  var lastReply = '';

  function setStatus(message, kind) {
    el.status.textContent = message;
    el.status.className = 'status' + (kind ? ' ' + kind : '');
  }

  // 設定画面へのリンク付きの状態表示（DOM を組んで textContent で入れる）。
  function setStatusWithSettingsLink(message) {
    el.status.textContent = '';
    el.status.className = 'status error';
    el.status.appendChild(document.createTextNode(message + ' '));
    var a = document.createElement('a');
    a.href = 'settings.html';
    a.className = 'link';
    a.textContent = '設定画面を開く';
    el.status.appendChild(a);
  }

  function setBusy(on) {
    el.create.disabled = on;
    el.create.textContent = on ? '作成中…' : '作成';
    el.create.setAttribute('aria-busy', on ? 'true' : 'false');
  }

  function readBody(item) {
    return new Promise(function (resolve) {
      try {
        item.body.getAsync(Office.CoercionType.Text, function (res) {
          resolve(res && res.status === Office.AsyncResultStatus.Succeeded ? (res.value || '') : '');
        });
      } catch (err) {
        resolve('');
      }
    });
  }

  function readMail(item) {
    return readBody(item).then(function (body) {
      return {
        subject: item.subject || '',
        from: item.from || null,
        to: item.to || [],
        cc: item.cc || [],
        body: body
      };
    });
  }

  function onCreate() {
    var purpose = el.purpose.value.trim();
    if (!purpose) {
      setStatus('趣旨を入力してください。', 'error');
      return;
    }
    var roaming = Office.context.roamingSettings;
    var settings = Store.readAll(roaming);
    if (!settings.apiKey) {
      setStatusWithSettingsLink('APIキーが未設定です。');
      return;
    }
    var item = Office.context.mailbox.item;
    var profile = Office.context.mailbox.userProfile || {};
    var timeoutMs = window.AI_REPLY_TIMEOUT_MS || DEFAULT_TIMEOUT_MS;

    setBusy(true);
    setStatus('作成しています…', '');

    readMail(item)
      .then(function (mail) {
        var requestBody = Core.buildRequestBody({
          model: settings.model,
          purpose: purpose,
          mail: mail,
          myAddress: profile.emailAddress || '',
          styleSample: settings.styleSample,
          extraInstructions: settings.extraInstructions
        });
        return Core.callClaude({
          apiKey: settings.apiKey,
          requestBody: requestBody,
          fetchImpl: window.fetch.bind(window),
          timeoutMs: timeoutMs
        });
      })
      .then(function (res) {
        if (res.ok) {
          lastReply = res.text;
          el.result.value = res.text;
          el.resultBlock.hidden = false;
          setStatus('作成しました。内容を確認して、返信画面を開くか、コピーしてください。', 'ok');
        } else {
          var code = typeof res.status === 'number' ? '（ステータス ' + res.status + '）' : '';
          setStatus(res.message + code, 'error');
        }
      })
      .catch(function () {
        setStatus('想定外のエラーが起きました。もう一度お試しください。', 'error');
      })
      .then(function () {
        setBusy(false);
      });
  }

  function onOpenReply() {
    if (!lastReply) return;
    var html = Core.buildReplyHtmlBody(lastReply);
    var item = Office.context.mailbox.item;
    try {
      if (el.replyAll.checked) {
        item.displayReplyAllForm({ htmlBody: html });
      } else {
        item.displayReplyForm({ htmlBody: html });
      }
      setStatus('返信画面を開きました。', 'ok');
    } catch (err) {
      setStatus('返信画面を開けませんでした。「コピー」で文面をコピーして、返信に貼り付けてください。', 'error');
    }
  }

  function copyText(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }
    return new Promise(function (resolve, reject) {
      el.result.select();
      var ok = document.execCommand && document.execCommand('copy');
      if (ok) resolve(); else reject(new Error('copy failed'));
    });
  }

  function onCopy() {
    if (!lastReply) return;
    copyText(lastReply)
      .then(function () {
        setStatus('コピーしました。返信画面に貼り付けてください。', 'ok');
      })
      .catch(function () {
        el.result.select();
        setStatus('自動でコピーできませんでした。文面を選んでコピーしてください。', 'error');
      });
  }

  function onReplyAllChanged() {
    Store.writeSome(Office.context.roamingSettings, { replyAll: el.replyAll.checked })
      .catch(function () {
        setStatus('「全員に返信」の選択を保存できませんでした。', 'error');
      });
  }

  // iPhone・Android の Outlook では返信画面を開く API が使えないため、コピーを主経路にする。
  function applyPlatform() {
    var platform = Office.context.platform;
    if (platform !== 'iOS' && platform !== 'Android') return;
    el.openReply.hidden = true;
    el.replyAllLabel.hidden = true;
    el.copy.className = 'btn primary';
    el.platformNote.hidden = false;
    el.platformNote.textContent =
      'iPhone・Android の Outlook では返信画面を直接開けません。「コピー」で文面をコピーし、返信画面に貼り付けてください。';
  }

  function init() {
    var settings = Store.readAll(Office.context.roamingSettings);
    el.replyAll.checked = settings.replyAll;
    applyPlatform();

    el.create.addEventListener('click', onCreate);
    el.openReply.addEventListener('click', onOpenReply);
    el.copy.addEventListener('click', onCopy);
    el.replyAll.addEventListener('change', onReplyAllChanged);
  }

  Office.onReady(function () {
    init();
  });
})();
