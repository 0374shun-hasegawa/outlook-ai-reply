/*
 * Office.context.roamingSettings への読み書き。
 * API キーはここだけで扱い、画面やログには出さない。
 * API キーはブラウザの保存領域には置かない（設計意図/outlook-addon.md）。
 */
(function (root) {
  'use strict';

  const DEFAULT_MODEL = 'claude-sonnet-5-5';

  // roaming に保存する項目名。
  // modelPicked は「利用者がモデルを既定以外に明示的に選んだ」ことの印。
  // 過去の保存では model に旧既定（claude-sonnet-4-5）が必ず入っていたため、印が無ければ既定を使う。
  const NAMES = ['apiKey', 'model', 'modelPicked', 'styleSample', 'extraInstructions', 'replyAll'];

  function readAll(roaming) {
    const raw = function (k) { const v = roaming.get(k); return v === undefined || v === null ? '' : v; };
    const picked = roaming.get('modelPicked') === true;
    return {
      apiKey: String(raw('apiKey')),
      model: picked && String(raw('model')) ? String(raw('model')) : DEFAULT_MODEL,
      styleSample: String(raw('styleSample')),
      extraInstructions: String(raw('extraInstructions')),
      // 未設定のときは「全員に返信」を既定にする。
      replyAll: roaming.get('replyAll') !== false
    };
  }

  // Outlook on the web は、作業ウィンドウ内でページを移る（設定画面→作成画面）と、
  // 作業ウィンドウを開いた時点の roamingSettings を渡し直す。設定画面と作成画面を同じページ内の
  // 切り替えにしたうえで、保存した値はこのページの中（メモリ上）にも控え、作成時はそちらを優先する。
  // メモリ上だけなので、作業ウィンドウを閉じると消える（ブラウザの保存領域には書かない）。
  const recent = {};

  // 作成のたびに呼ぶ。roamingSettings の値に、この画面で保存した直後の値を重ねる。
  function readLatest(roaming) {
    const view = {
      get: function (k) {
        return Object.prototype.hasOwnProperty.call(recent, k) ? recent[k] : roaming.get(k);
      }
    };
    return readAll(view);
  }

  function rememberRecent(values) {
    NAMES.forEach(function (k) {
      if (Object.prototype.hasOwnProperty.call(values, k)) recent[k] = values[k];
    });
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
