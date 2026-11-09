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

  // values に含まれるキーだけを set し、保存完了を Promise で返す。
  function writeSome(roaming, values) {
    Object.keys(values).forEach(function (k) { roaming.set(k, values[k]); });
    return new Promise(function (resolve, reject) {
      roaming.saveAsync(function (result) {
        if (result && result.status === 'succeeded') {
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
    writeSome: writeSome
  };
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = root.ReplyStore;
  }
})(typeof window !== 'undefined' ? window : globalThis);
