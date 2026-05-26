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

class SuggestionsCard extends FormattingSettingsCard {
    name = "suggestions";
    displayName = "候補表示";

    targetColumnName = new formattingSettings.TextInput({
        name: "targetColumnName",
        displayName: "候補を表示する列名（カンマ区切りで複数可）",
        value: "",
        placeholder: "例: 組織名, 部署",
    });

    slices: FormattingSettingsSlice[] = [this.targetColumnName];
}

export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    appearanceCard = new AppearanceCard();
    suggestionsCard = new SuggestionsCard();
    cards = [this.appearanceCard, this.suggestionsCard];
}
