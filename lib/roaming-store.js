/*
 * Office.context.roamingSettings への読み書き。
 * API キーはここだけで扱い、画面やログには出さない。
 */
(function (root) {
  'use strict';

  const DEFAULT_MODEL = 'claude-sonnet-4-5';

  // roaming に保存する項目名。
  const NAMES = ['apiKey', 'model', 'styleSample', 'extraInstructions', 'replyAll'];

  function readAll(roaming) {
    const raw = function (k) { const v = roaming.get(k); return v === undefined || v === null ? '' : v; };
    return {
      apiKey: String(raw('apiKey')),
      model: String(raw('model')) || DEFAULT_MODEL,
      styleSample: String(raw('styleSample')),
      extraInstructions: String(raw('extraInstructions')),
      // 未設定のときは「全員に返信」を既定にする。
      replyAll: roaming.get('replyAll') !== false
    };
  }

  // Outlook on the web は、作業ウィンドウ内でページを移る（設定画面→作成画面）と、
  // 作業ウィンドウを開いた時点の roamingSettings を渡し直す。設定画面で保存した値は、
  // 作業ウィンドウを閉じて開き直すまで見えない。そのため保存した値を同じ作業ウィンドウの
  // sessionStorage（閉じると消える）にも控え、作成時はそちらを優先する。
  const RECENT_KEY = 'aiReplyRecentSettings';

  function sessionStore() {
    try {
      return root.sessionStorage || null;
    } catch (e) {
      return null;
    }
  }

  function readRecent() {
    const ss = sessionStore();
    if (!ss) return {};
    try {
      const v = JSON.parse(ss.getItem(RECENT_KEY) || '{}');
      return v && typeof v === 'object' ? v : {};
    } catch (e) {
      return {};
    }
  }

  function rememberRecent(values) {
    const ss = sessionStore();
    if (!ss) return;
    const recent = readRecent();
    NAMES.forEach(function (k) {
      if (Object.prototype.hasOwnProperty.call(values, k)) recent[k] = values[k];
    });
    try {
      ss.setItem(RECENT_KEY, JSON.stringify(recent));
    } catch (e) {
      // 控えられなくても roamingSettings への保存は済んでいる。
    }
  }

  // 作成のたびに呼ぶ。roamingSettings の値に、この作業ウィンドウで保存した直後の値を重ねる。
  function readLatest(roaming) {
    const recent = readRecent();
    const view = {
      get: function (k) {
        return Object.prototype.hasOwnProperty.call(recent, k) ? recent[k] : roaming.get(k);
      }
    };
    return readAll(view);
  }

  // values に含まれるキーだけを set し、保存完了を Promise で返す。
  function writeSome(roaming, values) {
    Object.keys(values).forEach(function (k) { roaming.set(k, values[k]); });
    return new Promise(function (resolve, reject) {
      roaming.saveAsync(function (result) {
        if (result && result.status === 'succeeded') {
          rememberRecent(values);
          resolve();
        } else {
          reject(new Error('設定の保存に失敗しました。'));
        }
      });
    });
  }

  root.ReplyStore = {
    NAMES: NAMES,
    DEFAULT_MODEL: DEFAULT_MODEL,
    readAll: readAll,
    readLatest: readLatest,
    writeSome: writeSome
  };
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = root.ReplyStore;
  }
})(typeof window !== 'undefined' ? window : globalThis);
