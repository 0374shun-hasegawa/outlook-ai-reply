/*
 * AI返信アドインの純粋ロジック（DOM・Office.js に依存しない）。
 * ブラウザでは window.ReplyCore、node（テスト）では require() で使う。
 */
(function (root) {
  'use strict';

  // Outlook の roamingSettings 合計上限（約32KB）。
  const ROAMING_LIMIT_BYTES = 32 * 1024;
  const DEFAULT_MODEL = 'claude-sonnet-5-5';
  const API_URL = 'https://api.anthropic.com/v1/messages';
  const API_VERSION = '2023-06-01';
  const MAX_TOKENS = 2048;

  // 指示文は、実際に使って効果を確かめている社内の返信文テンプレートに合わせる（設計意図参照）。
  // 返信は元メールへの返信なので、「新しいメール」は「返信メールの本文」と言い換えている。
  function buildSenderText(input) {
    const name = (input.myName || '').trim();
    const addr = (input.myAddress || '').trim() || '（不明）';
    return name ? name + ' <' + addr + '>' : '<' + addr + '>';
  }

  function buildInstructionBody(sender, hasStyle) {
    const lines = [
      '以下のメール履歴を参考にして、元のメールへの返信メールの本文を作成してください。',
      '依頼の目的は、私がメールの返信を楽にやれるようにしたいからです。',
      '',
      '# 入力の構成',
      '- 「過去の経緯確認用のメール履歴」: 私が開いている元のメール（本文・件名・差出人・宛先）です。'
    ];
    if (hasStyle) {
      lines.push('- 「文体参考用メール履歴」: 私が過去に送ったメールです。');
    }
    lines.push('- 「新規メールの内容」: 返信メールの本文に盛り込む趣旨です。');
    lines.push('', '# 指示');
    if (hasStyle) {
      lines.push(
        '1. 上記の文体参考用メール履歴から、私（送信者：' + sender + '）の文体、挨拶、自己紹介、会社名の表記方法を正確に模倣してください。',
        '2. 特に以下の点に注意してください：',
        '   - 宛名の書き方（例：〇〇様）',
        '   - 自己紹介の仕方（例：お世話になっておりますテクノアの〇〇です）',
        '   - 会社名の表記（例：株式会社〇〇）',
        '   - 文末表現や敬語の使い方',
        '3. 過去の経緯確認用のメール履歴を参照し、これまでのやり取りの文脈や背景を理解してください。新規メールの内容がこの経緯と整合性を持つようにしてください。',
        '4. 「新規メールの内容」に記載された情報を使用して、返信メールの本文を作成してください。',
        '5. 文体参考用メール履歴にある私のメールの文体や構成を忠実に再現しつつ、新規メールの内容を盛り込んでください。',
        '6. 履歴にない表現や言い回しを使用する場合は、できるだけ控えめにし、全体的な調子を維持してください。'
      );
    } else {
      lines.push(
        '1. 私（送信者：' + sender + '）の返信として、文体参考が無いので一般的で丁寧なビジネスメールの文体で書いてください。',
        '2. 特に以下の点に注意してください：',
        '   - 宛名の書き方（例：〇〇様）',
        '   - 文末表現や敬語の使い方',
        '3. 過去の経緯確認用のメール履歴を参照し、これまでのやり取りの文脈や背景を理解してください。新規メールの内容がこの経緯と整合性を持つようにしてください。',
        '4. 「新規メールの内容」に記載された情報を使用して、返信メールの本文を作成してください。'
      );
    }
    lines.push(
      '',
      '# 守ること',
      '- 宛名・肩書き・会社名は、元メールの差出人表示名と署名に書かれているものだけを使い、推測で足さないでください（書かれていない肩書きは付けない）。',
      '- URL・日時・数字はそのまま書いてください。',
      '- 件名、前置き、解説、「以下が返信文です」などの説明は出さず、返信メールの本文だけを出力してください。'
    );
    return lines.join('\n');
  }

  // 文体参考あり・送信者不明の既定の形（定数参照用）
  const SYSTEM_PROMPT_BASE = buildInstructionBody('（不明）', true);

  // options: { sender, hasStyle }（hasStyle 省略時は文体参考ありとして扱う）
  function buildSystemPrompt(extraInstructions, options) {
    const opts = options || {};
    const hasStyle = opts.hasStyle !== false;
    const base = buildInstructionBody(opts.sender || '（不明）', hasStyle);
    const extra = (extraInstructions || '').trim();
    if (!extra) return base;
    return base + '\n\n# 追加の指示\n' + extra;
  }

  function formatAddressList(list) {
    if (!Array.isArray(list) || list.length === 0) return '（なし）';
    return list
      .map(function (a) {
        if (!a) return '';
        const name = a.displayName ? a.displayName + ' ' : '';
        return name + '<' + (a.emailAddress || '') + '>';
      })
      .filter(Boolean)
      .join(', ');
  }

  // mail: { subject, from, to, cc, body }  purpose: 利用者が入力した趣旨
  function buildUserMessage(input) {
    const mail = input.mail || {};
    const from = mail.from || {};
    const fromText = (from.displayName ? from.displayName + ' ' : '') + '<' + (from.emailAddress || '') + '>';
    const styleSample = (input.styleSample || '').trim();
    const parts = [
      '#過去の経緯確認用のメール履歴',
      '件名: ' + (mail.subject || '（なし）'),
      '差出人: ' + fromText,
      '宛先: ' + formatAddressList(mail.to),
      'CC: ' + formatAddressList(mail.cc),
      '本文:',
      mail.body || '（本文を取得できませんでした）',
      ''
    ];
    if (styleSample) {
      parts.push('#文体参考用メール履歴', styleSample, '');
    }
    parts.push('#新規メールの内容（返信メールの本文に盛り込む趣旨）', input.purpose || '');
    return parts.join('\n');
  }

  function buildRequestBody(input) {
    return {
      model: input.model || DEFAULT_MODEL,
      max_tokens: MAX_TOKENS,
      system: buildSystemPrompt(input.extraInstructions, {
        sender: buildSenderText(input),
        hasStyle: !!(input.styleSample || '').trim()
      }),
      messages: [{ role: 'user', content: buildUserMessage(input) }]
    };
  }

  function buildHeaders(apiKey) {
    return {
      'x-api-key': apiKey,
      'anthropic-version': API_VERSION,
      'anthropic-dangerous-direct-browser-access': 'true',
      'content-type': 'application/json'
    };
  }

  // 応答 JSON から本文（text ブロックの連結）を取り出す。
  function extractReplyText(responseJson) {
    const blocks = (responseJson && Array.isArray(responseJson.content)) ? responseJson.content : [];
    return blocks
      .filter(function (b) { return b && b.type === 'text' && typeof b.text === 'string'; })
      .map(function (b) { return b.text; })
      .join('')
      .trim();
  }

  // HTTP ステータスと通信失敗を、利用者向けの日本語説明にする。
  function describeApiError(status) {
    if (status === 'network') {
      return '通信に失敗しました。ネットワーク接続を確認して、もう一度お試しください。';
    }
    if (status === 'timeout') {
      return '応答がありませんでした（時間切れ）。もう一度お試しください。';
    }
    if (status === 400) return 'リクエストが正しくありません（モデル名などを設定画面で確認してください）。';
    if (status === 401) return 'APIキーが正しくありません。設定画面でキーを確認してください。';
    if (status === 403) return 'このAPIキーではこの操作が許可されていません。キーの権限を確認してください。';
    if (status === 404) return 'モデルが見つかりません。設定画面でモデル名を確認してください。';
    if (status === 429) return '利用の上限（レート制限）に達しました。少し待ってから、もう一度お試しください。';
    if (typeof status === 'number' && status >= 500) {
      return 'Claude 側でエラーが起きました。時間をおいて、もう一度お試しください。';
    }
    return '想定外のエラーが起きました（ステータス ' + status + '）。';
  }

  function utf8ByteLength(text) {
    // TextEncoder に頼らず、どの実行環境でも同じ UTF-8 バイト数を返す。
    return unescape(encodeURIComponent(String(text))).length;
  }

  // roamingSettings に保存する全項目（キーと値）の合計バイト数。
  function roamingTotalBytes(settingsObject) {
    return Object.keys(settingsObject).reduce(function (sum, k) {
      const v = settingsObject[k];
      const text = typeof v === 'string' ? v : JSON.stringify(v);
      return sum + utf8ByteLength(k) + utf8ByteLength(text === undefined ? '' : text);
    }, 0);
  }

  function checkRoamingSize(settingsObject) {
    const bytes = roamingTotalBytes(settingsObject);
    return { bytes: bytes, limit: ROAMING_LIMIT_BYTES, ok: bytes <= ROAMING_LIMIT_BYTES };
  }

  // 文字列を HTML として安全にし、改行を <br> にする。
  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function textToHtml(text) {
    return escapeHtml(text).replace(/\r\n|\r|\n/g, '<br>');
  }

  // 返信画面の htmlBody に入れる文面。
  function buildReplyHtmlBody(replyText) {
    return textToHtml(replyText);
  }

  // Claude API を1回呼ぶ。fetchImpl は fetch 互換（テストで差し替える）。
  // 戻り値: { ok: true, text } または { ok: false, status, message }
  async function callClaude(options) {
    const fetchImpl = options.fetchImpl;
    const timeoutMs = options.timeoutMs || 60000;
    const controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
    let timer = null;
    if (controller) timer = setTimeout(function () { controller.abort(); }, timeoutMs);
    let response;
    try {
      response = await fetchImpl(API_URL, {
        method: 'POST',
        headers: buildHeaders(options.apiKey),
        body: JSON.stringify(options.requestBody),
        signal: controller ? controller.signal : undefined
      });
    } catch (err) {
      if (timer) clearTimeout(timer);
      const isTimeout = !!(err && err.name === 'AbortError');
      const kind = isTimeout ? 'timeout' : 'network';
      return { ok: false, status: kind, message: describeApiError(kind) };
    }
    if (timer) clearTimeout(timer);
    if (!response.ok) {
      return { ok: false, status: response.status, message: describeApiError(response.status) };
    }
    let json;
    try {
      json = await response.json();
    } catch (err) {
      return { ok: false, status: 'network', message: describeApiError('network') };
    }
    const text = extractReplyText(json);
    if (!text) {
      return { ok: false, status: 'empty', message: '返信文が空でした。趣旨を書き足してもう一度お試しください。' };
    }
    return { ok: true, text: text };
  }

  const api = {
    DEFAULT_MODEL: DEFAULT_MODEL,
    API_URL: API_URL,
    ROAMING_LIMIT_BYTES: ROAMING_LIMIT_BYTES,
    SYSTEM_PROMPT_BASE: SYSTEM_PROMPT_BASE,
    buildSystemPrompt: buildSystemPrompt,
    buildSenderText: buildSenderText,
    buildUserMessage: buildUserMessage,
    buildRequestBody: buildRequestBody,
    buildHeaders: buildHeaders,
    extractReplyText: extractReplyText,
    describeApiError: describeApiError,
    utf8ByteLength: utf8ByteLength,
    roamingTotalBytes: roamingTotalBytes,
    checkRoamingSize: checkRoamingSize,
    escapeHtml: escapeHtml,
    textToHtml: textToHtml,
    buildReplyHtmlBody: buildReplyHtmlBody,
    callClaude: callClaude
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    root.ReplyCore = api;
  }
})(typeof window !== 'undefined' ? window : globalThis);
