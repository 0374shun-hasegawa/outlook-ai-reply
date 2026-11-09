/*
 * 作業ウィンドウ（AI返信）。Office.js と Claude API をつなぐ。
 * API キーは roamingSettings から読むだけで、画面にもログにも出さない。
 */
(function () {
  'use strict';

  var Core = window.ReplyCore;
  var Store = window.ReplyStore;
  var Pending = window.PendingReply;
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
  // 元メールがテキスト形式か（作成時に調べる）。返信画面への入れ方を変えるのに使う。
  var lastMailIsPlainText = false;
  // 保留中の返信文の保存（作成直後に始める）。null は保存していない、Promise は保存中または保存済み。
  var pendingSaved = null;

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

  function readBodyHtml(item) {
    return new Promise(function (resolve) {
      try {
        item.body.getAsync(Office.CoercionType.Html, function (res) {
          resolve(res && res.status === Office.AsyncResultStatus.Succeeded ? (res.value || '') : '');
        });
      } catch (err) {
        resolve('');
      }
    });
  }

  function readMail(item) {
    return Promise.all([readBody(item), readBodyHtml(item)]).then(function (bodies) {
      var body = bodies[0];
      lastMailIsPlainText = Core.isPlainTextMailHtml(bodies[1]);
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
    // 設定画面で保存した直後でも最新の値を使うため、作成のたびに読み直す。
    var settings = Store.readLatest(Office.context.roamingSettings);
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
          // 返信画面へ自動で入れるのは、htmlBody が使えないテキスト形式のメールと、
          // 返信画面を開く API が無い iPhone・Android。HTML 形式は htmlBody で直接渡す（二重に入れない）。
          if (lastMailIsPlainText || isMobile()) {
            savePending(item, res.text);
          } else {
            pendingSaved = null;
          }
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

  // 返信画面が開いたときに差し込む返信文を roamingSettings に残す（launchevent.js が読み出して消す）。
  function savePending(item, text) {
    var record = Pending.create(text, item.conversationId, el.replyAll.checked, Date.now());
    var values = {};
    values[Pending.KEY] = record;
    var state = { done: false, ok: false };
    pendingSaved = Store.writeSome(Office.context.roamingSettings, values)
      .then(function () { state.ok = true; }, function () { state.ok = false; })
      .then(function () { state.done = true; return state.ok; });
    pendingSaved.state = state;
  }

  function isMobile() {
    var platform = Office.context.platform;
    return platform === 'iOS' || platform === 'Android';
  }

  function displayReply(formData) {
    var item = Office.context.mailbox.item;
    if (el.replyAll.checked) {
      item.displayReplyAllForm(formData);
    } else {
      item.displayReplyForm(formData);
    }
  }

  function openFailed() {
    setStatus('返信画面を開けませんでした。「コピー」で文面をコピーして、返信に貼り付けてください。', 'error');
  }

  // テキスト形式のメールは htmlBody のタグが文字で出るため本文を渡さず、開いた返信画面に
  // イベント処理が差し込む。保留の保存が終わる前に開くと差し込めないので、終わるのを待つ。
  function openReplyForPending(saving) {
    function open(ok) {
      if (!ok) {
        setStatus('返信文を返信画面へ渡す準備に失敗しました。「コピー」で文面をコピーして、返信に貼り付けてください。', 'error');
        return;
      }
      try {
        displayReply('');
        setStatus('返信画面を開きました。本文の先頭に返信文が入ります（入らなければ「コピー」で貼り付けてください）。', 'ok');
      } catch (err) {
        openFailed();
      }
    }
    if (saving.state.done) {
      open(saving.state.ok);
    } else {
      saving.then(open);
    }
  }

  function onOpenReply() {
    if (!lastReply) return;
    if (pendingSaved) {
      openReplyForPending(pendingSaved);
      return;
    }
    try {
      displayReply({ htmlBody: Core.buildReplyHtmlBody(lastReply) });
      setStatus('返信画面を開きました。', 'ok');
    } catch (err) {
      openFailed();
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

  // iPhone・Android の Outlook では返信画面を開く API が使えない。返信ボタンで開いた返信画面に
  // イベント処理（launchevent.js）が差し込むので、それを案内し、入らないときのためにコピーを残す。
  function applyPlatform() {
    var platform = Office.context.platform;
    if (platform !== 'iOS' && platform !== 'Android') return;
    el.openReply.hidden = true;
    el.replyAllLabel.hidden = true;
    el.copy.className = 'btn primary';
    el.platformNote.hidden = false;
    el.platformNote.textContent =
      '返信ボタンを押すと本文に自動で入ります（入らなければコピーして貼り付け）。';
  }

  function init() {
    var settings = Store.readLatest(Office.context.roamingSettings);
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
