/*
 * 設定画面（作成画面と同じページ内の切り替え）。API キーは「入力された時だけ」保存し、保存済みの値は画面に戻さない。
 * 保存した値は roamingSettings（と、このページの中で控える分）だけに置く。ブラウザの保存領域には書かない。
 */
(function () {
  'use strict';

  var Core = window.ReplyCore;
  var Store = window.ReplyStore;

  function $(id) { return document.getElementById(id); }

  var el = {
    createView: $('createView'),
    settingsView: $('settingsView'),
    apiKey: $('apiKey'),
    keyState: $('keyState'),
    clearKey: $('clearKey'),
    model: $('model'),
    styleSample: $('styleSample'),
    extraInstructions: $('extraInstructions'),
    sizeNote: $('sizeNote'),
    save: $('save'),
    status: $('settingsStatus')
  };

  function setStatus(message, kind) {
    el.status.textContent = message;
    el.status.className = 'status' + (kind ? ' ' + kind : '');
  }

  function kb(bytes) {
    return (bytes / 1024).toFixed(1);
  }

  // 保存しようとしている内容（入力欄の値で上書きしたもの）。
  function buildCandidate(stored) {
    var typedKey = el.apiKey.value.trim();
    return {
      apiKey: typedKey || stored.apiKey,
      model: el.model.value || Store.DEFAULT_MODEL,
      styleSample: el.styleSample.value,
      extraInstructions: el.extraInstructions.value,
      replyAll: stored.replyAll
    };
  }

  function updateSizeNote() {
    var stored = Store.readLatest(Office.context.roamingSettings);
    var check = Core.checkRoamingSize(buildCandidate(stored));
    el.sizeNote.className = 'note' + (check.ok ? '' : ' error');
    el.sizeNote.textContent = check.ok
      ? '保存量の目安：約 ' + kb(check.bytes) + ' KB（上限 約32KB）'
      : '保存量が上限を超えています（約 ' + kb(check.bytes) + ' KB / 上限 約32KB）。文体の参考を短くしてください。';
  }

  function renderKeyState(hasKey) {
    el.keyState.textContent = hasKey
      ? '保存済みです（キーの値は表示しません）。新しいキーを入力すると置き換えます。'
      : '未保存です。';
    el.clearKey.hidden = !hasKey;
  }

  function onSave() {
    var roaming = Office.context.roamingSettings;
    var stored = Store.readLatest(roaming);
    var candidate = buildCandidate(stored);
    var check = Core.checkRoamingSize(candidate);

    if (!check.ok) {
      setStatus(
        '文体の参考が長すぎます（約 ' + kb(check.bytes) + ' KB）。Outlook に保存できる上限は約32KBです。' +
        '代表的な数通に絞るなど短くして、もう一度保存してください。今回は保存していません。',
        'error'
      );
      updateSizeNote();
      return;
    }

    var values = {
      model: candidate.model,
      // 既定以外を選んだときだけ印を付ける。印が無い保存値は既定として扱う（Store.readAll）。
      modelPicked: candidate.model !== Store.DEFAULT_MODEL,
      styleSample: candidate.styleSample,
      extraInstructions: candidate.extraInstructions
    };
    if (el.apiKey.value.trim()) values.apiKey = el.apiKey.value.trim();

    el.save.disabled = true;
    Store.writeSome(roaming, values)
      .then(function () {
        el.apiKey.value = '';
        renderKeyState(!!candidate.apiKey);
        setStatus('保存しました。', 'ok');
      })
      .catch(function () {
        setStatus('保存に失敗しました。もう一度お試しください。', 'error');
      })
      .then(function () {
        el.save.disabled = false;
        updateSizeNote();
      });
  }

  function onClearKey() {
    el.clearKey.disabled = true;
    Store.writeSome(Office.context.roamingSettings, { apiKey: '' })
      .then(function () {
        renderKeyState(false);
        setStatus('APIキーを削除しました。', 'ok');
      })
      .catch(function () {
        setStatus('削除に失敗しました。もう一度お試しください。', 'error');
      })
      .then(function () {
        el.clearKey.disabled = false;
      });
  }

  // 作成画面と設定画面の切り替え（data-view-link="create" | "settings"）。
  function showView(name) {
    var toSettings = name === 'settings';
    el.settingsView.hidden = !toSettings;
    el.createView.hidden = toSettings;
  }

  function onViewLinkClick(event) {
    var link = event.target && event.target.closest ? event.target.closest('[data-view-link]') : null;
    if (!link) return;
    event.preventDefault();
    showView(link.getAttribute('data-view-link'));
  }

  function init() {
    var stored = Store.readLatest(Office.context.roamingSettings);
    el.model.value = stored.model;
    el.styleSample.value = stored.styleSample;
    el.extraInstructions.value = stored.extraInstructions;
    el.apiKey.value = '';
    renderKeyState(!!stored.apiKey);
    updateSizeNote();

    el.styleSample.addEventListener('input', updateSizeNote);
    el.extraInstructions.addEventListener('input', updateSizeNote);
    el.model.addEventListener('change', updateSizeNote);
    el.apiKey.addEventListener('input', updateSizeNote);
    el.save.addEventListener('click', onSave);
    el.clearKey.addEventListener('click', onClearKey);
    document.addEventListener('click', onViewLinkClick);
    showView('create');
  }

  Office.onReady(function () {
    init();
  });
})();
