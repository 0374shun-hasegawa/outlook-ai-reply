/*
 * OnNewMessageCompose のイベント処理。返信画面が開いたら、作業ウィンドウが残した保留中の返信文を
 * 本文の先頭に差し込む。期限切れ・別の会話・保留なしのときは何もしない。
 * 元メールがテキスト形式だと返信画面もテキスト編集になり、HTML を渡すとタグが文字で出るため、
 * 本文の形式を見て入れ方を変える。
 */
(function (root) {
  'use strict';

  // Outlook はイベント処理を 60 秒で打ち切る。その前に必ず completed を呼ぶ。
  const GUARD_MS = 50 * 1000;

  function isMobile(Office) {
    const p = Office.context && Office.context.platform;
    return p === 'iOS' || p === 'Android';
  }

  function callAsync(fn, args) {
    return new Promise(function (resolve) {
      try {
        fn.apply(null, args.concat([function (res) { resolve(res || {}); }]));
      } catch (err) {
        resolve({ status: 'failed', error: err });
      }
    });
  }

  function succeeded(Office, res) {
    return res && res.status === Office.AsyncResultStatus.Succeeded;
  }

  function readPending(roaming, key) {
    const v = roaming.get(key);
    if (typeof v === 'string') {
      try { return JSON.parse(v); } catch (e) { return null; }
    }
    return v || null;
  }

  function removePending(roaming, key) {
    roaming.remove(key);
    return new Promise(function (resolve) {
      try {
        roaming.saveAsync(function () { resolve(); });
      } catch (e) {
        resolve();
      }
    });
  }

  function bodyType(Office, body) {
    if (typeof body.getTypeAsync !== 'function') return Promise.resolve('text');
    return callAsync(body.getTypeAsync.bind(body), []).then(function (res) {
      if (!succeeded(Office, res)) return 'text';
      return res.value === Office.CoercionType.Html ? 'html' : 'text';
    });
  }

  // 使える差し込み方法を順に試す。iPhone・Android では prependAsync が使えないため最初から外す。
  function insertionMethods(Office, body) {
    const names = isMobile(Office)
      ? ['setSelectedDataAsync', 'setSignatureAsync']
      : ['prependAsync', 'setSelectedDataAsync', 'setSignatureAsync'];
    return names.filter(function (n) { return typeof body[n] === 'function'; });
  }

  function tryInsert(Office, body, content, coercionType) {
    const methods = insertionMethods(Office, body);
    let i = 0;
    function next() {
      if (i >= methods.length) return Promise.resolve(null);
      const name = methods[i++];
      return callAsync(body[name].bind(body), [content, { coercionType: coercionType }]).then(function (res) {
        return succeeded(Office, res) ? name : next();
      });
    }
    return next();
  }

  // 戻り値: 'inserted:<方法>' / 'none' / 'expired' / 'mismatch' / 'failed'（テストと確認用）
  function handle(Office, PendingReply, nowMs) {
    const roaming = Office.context.roamingSettings;
    const item = Office.context.mailbox.item;
    const pending = readPending(roaming, PendingReply.KEY);
    if (!pending) return Promise.resolve('none');
    const conversationId = item && item.conversationId;
    if (!PendingReply.isApplicable(pending, conversationId, nowMs)) {
      const expired = typeof pending.createdAt !== 'number' || nowMs - pending.createdAt > PendingReply.MAX_AGE_MS;
      if (expired) return removePending(roaming, PendingReply.KEY).then(function () { return 'expired'; });
      return Promise.resolve('mismatch');
    }
    return bodyType(Office, item.body).then(function (type) {
      const coercion = type === 'html' ? Office.CoercionType.Html : Office.CoercionType.Text;
      return tryInsert(Office, item.body, PendingReply.contentFor(type, pending.text), coercion);
    }).then(function (method) {
      if (!method) return 'failed';
      return removePending(roaming, PendingReply.KEY).then(function () { return 'inserted:' + method; });
    });
  }

  function onNewMessageComposeHandler(event) {
    const Office = root.Office;
    let done = false;
    function complete() {
      if (done) return;
      done = true;
      try { event.completed(); } catch (e) { /* 何もしない */ }
    }
    const guard = setTimeout(complete, GUARD_MS);
    Promise.resolve()
      .then(function () { return handle(Office, root.PendingReply, Date.now()); })
      .then(function (result) { root.AI_REPLY_LAST_LAUNCH = result; })
      .catch(function () { root.AI_REPLY_LAST_LAUNCH = 'error'; })
      .then(function () {
        clearTimeout(guard);
        complete();
      });
  }

  root.onNewMessageComposeHandler = onNewMessageComposeHandler;
  root.AiReplyLaunch = { handle: handle, onNewMessageComposeHandler: onNewMessageComposeHandler };

  if (root.Office && typeof root.Office.onReady === 'function') {
    root.Office.onReady(function () {});
  }
  if (root.Office && root.Office.actions && typeof root.Office.actions.associate === 'function') {
    root.Office.actions.associate('onNewMessageComposeHandler', onNewMessageComposeHandler);
  }
})(typeof window !== 'undefined' ? window : globalThis);
