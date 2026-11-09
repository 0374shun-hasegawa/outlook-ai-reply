/*
 * 保留中の返信文（閲覧画面で作った返信文を、返信画面が開いたときに差し込むための受け渡し）。
 * 作業ウィンドウが roamingSettings に保存し、OnNewMessageCompose のイベント処理が読み出して消す。
 * DOM・Office.js に依存しない（ブラウザでは window.PendingReply、node では require()）。
 */
(function (root) {
  'use strict';

  const KEY = 'pendingReply';
  // この時間を過ぎた保留は差し込まない（作ってから返信画面を開くまでの猶予）。
  const MAX_AGE_MS = 10 * 60 * 1000;

  function create(text, conversationId, replyAll, nowMs) {
    return {
      text: String(text),
      conversationId: conversationId || '',
      createdAt: nowMs,
      replyAll: !!replyAll
    };
  }

  // 保留を差し込んでよいか。期限内で、作成中のメールが同じ会話の返信であるときだけ true。
  function isApplicable(pending, composeConversationId, nowMs) {
    if (!pending || typeof pending !== 'object') return false;
    if (typeof pending.text !== 'string' || pending.text === '') return false;
    if (typeof pending.createdAt !== 'number') return false;
    const age = nowMs - pending.createdAt;
    if (age < 0 || age > MAX_AGE_MS) return false;
    if (!pending.conversationId || !composeConversationId) return false;
    return pending.conversationId === composeConversationId;
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // 本文の形式（'text' または 'html'）に合わせた差し込み内容。
  // 後ろに空行を 1 つ足し、元の本文（引用）とくっつかないようにする。
  function contentFor(bodyType, text) {
    const normalized = String(text).replace(/\r\n|\r/g, '\n');
    if (bodyType === 'html') {
      return escapeHtml(normalized).replace(/\n/g, '<br>') + '<br><br>';
    }
    return normalized + '\n\n';
  }

  const api = {
    KEY: KEY,
    MAX_AGE_MS: MAX_AGE_MS,
    create: create,
    isApplicable: isApplicable,
    contentFor: contentFor
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    root.PendingReply = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
