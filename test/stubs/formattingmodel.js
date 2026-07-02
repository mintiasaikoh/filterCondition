"use strict";
// powerbi-visuals-utils-formattingmodel のテスト用最小スタブ。
// 本物は ESM ビルドのみで Node の CJS require から読めないため、
// run.js の Module._load フックで本モジュールに差し替える。
class Base { constructor(o) { Object.assign(this, o ?? {}); } }
class SimpleCard extends Base {}
class Model extends Base {}
class Slice extends Base {}
class FontPicker extends Base {}
class NumUpDown extends Base {}
class ColorPicker extends Base {}
class TextInput extends Base {}
class ToggleSwitch extends Base {}

const formattingSettings = {
    SimpleCard, Model, Slice, FontPicker, NumUpDown, ColorPicker, TextInput, ToggleSwitch,
};

class FormattingSettingsService {
    populateFormattingSettingsModel(ctor, _dv) { return new ctor(); }
    buildFormattingModel(_m) { return {}; }
}

module.exports = { formattingSettings, FormattingSettingsService };
