"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsModel = formattingSettings.Model;
import FormattingSettingsSlice = formattingSettings.Slice;

class AppearanceCard extends FormattingSettingsCard {
    name = "appearance";
    displayName = "見た目";

    fontFamily = new formattingSettings.FontPicker({
        name: "fontFamily",
        displayName: "フォント",
        value: "Segoe UI, sans-serif",
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "フォントサイズ",
        value: 12,
    });

    accentColor = new formattingSettings.ColorPicker({
        name: "accentColor",
        displayName: "アクセントカラー",
        value: { value: "#0078d4" },
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "文字色",
        value: { value: "#252423" },
    });

    backgroundColor = new formattingSettings.ColorPicker({
        name: "backgroundColor",
        displayName: "背景色",
        value: { value: "#ffffff" },
    });

    slices: FormattingSettingsSlice[] = [
        this.fontFamily, this.fontSize, this.accentColor, this.fontColor, this.backgroundColor,
    ];
}

class BehaviorCard extends FormattingSettingsCard {
    name = "behavior";
    displayName = "動作";

    selfFilterEnabled = new formattingSettings.ToggleSwitch({
        name: "selfFilterEnabled",
        displayName: "候補のサーバー連動（実験的）",
        description: "「列」側の条件でも候補を絞り込む。環境によってはフィルター異常でビジュアルが壊れるため既定 OFF",
        value: false,
    });

    slices: FormattingSettingsSlice[] = [this.selfFilterEnabled];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    appearanceCard = new AppearanceCard();
    behaviorCard = new BehaviorCard();
    cards = [this.appearanceCard, this.behaviorCard];
}
