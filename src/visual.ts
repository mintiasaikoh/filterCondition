"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import VisualUpdateType = powerbi.VisualUpdateType;

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import { ConditionForm } from "./conditionForm";
import { FilterCondition } from "./filterEngine";
import {
    emitAdvancedFilter,
    restoreFromAdvancedFilters,
    ColumnLogic,
} from "./advancedFilterEmitter";
import { VisualFormattingSettingsModel } from "./settings";

export class Visual implements IVisual {
    private host: IVisualHost;
    private root: HTMLElement;
    private form: ConditionForm;

    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private lastDataView: DataView | null = null;
    private lastFilterSig = "";
    private pendingApply = false;
    private applyFallbackTimer: number | null = null;
    private uniquesCache: { catRef: unknown; result: string[][] } | null = null;
    private lastColsRef: unknown = null;
    private lastUniquesRef: unknown = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();

        this.root = document.createElement("div");
        this.root.className = "fc-visual";
        options.element.appendChild(this.root);

        this.form = new ConditionForm(this.root, {
            onChange: () => this.onFormChange(),
        });
    }

    public update(options: VisualUpdateOptions): void {
        // Power BI は 1 ビジュアル 1 クエリ。mapping は categorical 一本
        // （categories=候補列, values=列メジャー）。dataViews[0].categorical を使う。
        const dvs = options.dataViews ?? [];
        const dv = dvs[0] ?? null;
        this.lastDataView = dv;

        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, dv);
        this.applyAppearance();

        const cols = this.resolveColumns(dv);
        const uniques = this.extractUniques(dv, cols);
        // cols / uniques の参照が変わっていなければ DOM 再構築をスキップ
        if (cols !== this.lastColsRef || uniques !== this.lastUniquesRef) {
            this.lastColsRef = cols;
            this.lastUniquesRef = uniques;
            this.form.setColumns(cols, uniques);
        }

        // UI 状態は jsonFilters を唯一の真実源とする（dateCalendar と同方針）。
        // metadata.objects.state は書くが読まない: 「全フィルターリセット」ブックマークは
        // filter 層のみクリアして metadata を残すため、metadata 読み込みは stale state の
        // 温床になる。cross-session 復元もブックマークも jsonFilters が担う。
        // Data 更新時のみ未適用入力もリセット対象に（resize/style では typing を壊さない）。
        const isDataUpdate = ((options.type ?? VisualUpdateType.All) & VisualUpdateType.Data) !== 0;
        this.restoreFromJsonFilters(options.jsonFilters, cols, isDataUpdate);
    }

    /**
     * 条件フォームに出す列リストを構築。
     * categorical の categories（候補列）∪ values（列=メジャー）の source を
     * queryName で dedupe マージ。両方 categorical で確実に届く。
     * 「列」はメジャー集計されるが値は無視し metadata（queryName）のみ使用。
     * buildFilterTarget が集計ラッパー（Count/First/Sum）を剥がすのでフィルタ可。
     * 表示名は conditionForm.cleanLabel が「最初の 〜」を剥がして素の列名にする。
     */
    private resolveColumns(dv: DataView | null): powerbi.DataViewMetadataColumn[] {
        const out: powerbi.DataViewMetadataColumn[] = [];
        const seen = new Set<string>();
        const add = (c?: powerbi.DataViewMetadataColumn): void => {
            if (!c) return;
            const key = c.queryName ?? c.displayName ?? "";
            if (key) {
                if (seen.has(key)) return;
                seen.add(key);
            }
            out.push(c);
        };
        for (const cat of dv?.categorical?.categories ?? []) add(cat?.source);
        for (const val of dv?.categorical?.values ?? []) add(val?.source);
        return out;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ==========================================================

    private onFormChange(): void {
        const cols = (this.lastColsRef as powerbi.DataViewMetadataColumn[] | null) ?? [];
        if (cols.length === 0) return;

        const conds = this.form.getConditions();
        const columnLogic = this.form.getColumnLogic();

        const result = emitAdvancedFilter(this.host, cols, conds, columnLogic, this.lastFilterSig);
        if (result.emitted || result.sig !== this.lastFilterSig) {
            this.lastFilterSig = result.sig;
        }
        this.persist(conds, columnLogic);

        // 発火したら jsonFilters エコー受信までスピナー継続。発火しなければ即解除。
        if (result.emitted) {
            this.pendingApply = true;
            if (this.applyFallbackTimer !== null) clearTimeout(this.applyFallbackTimer);
            // 念のためのフォールバック（エコーが来ないケースで固まらないよう）
            this.applyFallbackTimer = window.setTimeout(() => this.clearPending(), 10000);
        } else {
            this.clearPending();
        }
    }

    /** スピナー解除＋フォールバックタイマー停止 */
    private clearPending(): void {
        this.pendingApply = false;
        if (this.applyFallbackTimer !== null) {
            clearTimeout(this.applyFallbackTimer);
            this.applyFallbackTimer = null;
        }
        this.form.setApplying(false);
    }

    // ==========================================================
    // 永続化
    // ==========================================================

    private persist(conds: FilterCondition[], columnLogic: ColumnLogic): void {
        this.host.persistProperties({
            merge: [{
                objectName: "state",
                selector: null,
                properties: {
                    conditionsJson: JSON.stringify(conds),
                    columnLogicJson: JSON.stringify(columnLogic),
                },
            }],
        });
    }

    // ==========================================================
    // jsonFilters 受信
    // ==========================================================

    private restoreFromJsonFilters(
        jsonFilters: powerbi.IFilter[] | undefined,
        cols: powerbi.DataViewMetadataColumn[],
        isDataUpdate: boolean,
    ): void {
        const restored = restoreFromAdvancedFilters(jsonFilters, cols);

        // ブックマーク / 外部スライサーで自分の filter が解除された場合は UI をリセット。
        // 未適用の入力中テキストも Data 更新（ブックマーク含む）時のみ掃除する。
        // resize / style の update では typing 中のテキストを壊さない。
        if (!restored) {
            // 自分の filter 解除が反映された（remove のエコー）→ スピナー解除
            if (this.pendingApply && isDataUpdate) this.clearPending();
            const shouldResetDirty = isDataUpdate && this.form.hasAnyInput();
            if (this.lastFilterSig !== "" || shouldResetDirty) {
                this.lastFilterSig = "";
                this.form.resetToDefault();
            }
            return;
        }

        // 自己発火エコー = 自分の filter がモデルに反映された確証 → スピナー解除
        if (restored.sig === this.lastFilterSig) {
            if (this.pendingApply) this.clearPending();
            return;
        }

        // 有効な active 条件が入ってきたら UI を上書き（ブックマーク含む）
        this.form.setState(restored.conditions, restored.columnLogic);
        this.lastFilterSig = restored.sig;
    }

    // ==========================================================

    /**
     * categorical.categories（=候補列にバインドされた列のみ）から distinct を抽出。
     * 候補列はユーザーが明示指定する低 card 列なので全 distinct を表示（DISPLAY_CAP まで）。
     * TEST/ダミーを含む値は除外。queryName で cols に割当。列(metadata のみ)は候補なし。
     */
    private extractUniques(
        dv: DataView | null,
        cols: powerbi.DataViewMetadataColumn[],
    ): string[][] {
        const categories = dv?.categorical?.categories ?? [];

        if (this.uniquesCache && this.uniquesCache.catRef === categories) {
            return this.uniquesCache.result;
        }

        if (categories.length === 0) {
            const empty = cols.map(() => [] as string[]);
            this.uniquesCache = { catRef: categories, result: empty };
            return empty;
        }

        const DISPLAY_CAP = 200; // datalist が肥大しないための上限
        const byQueryName = new Map<string, string[]>();
        for (const cat of categories) {
            const src = cat?.source;
            if (!src) continue;
            const set = new Set<string>();
            for (const v of cat.values ?? []) {
                if (v == null) continue;
                const s = String(v);
                if (s === "") continue;
                if (s.toUpperCase().includes("TEST")) continue;
                if (s.includes("ダミー")) continue;
                set.add(s);
                if (set.size >= DISPLAY_CAP) break;
            }
            byQueryName.set(src.queryName ?? "", Array.from(set).sort((a, b) => a.localeCompare(b)));
        }

        const result = cols.map(c => byQueryName.get(c?.queryName ?? "") ?? []);
        this.uniquesCache = { catRef: categories, result };
        return result;
    }

    private applyAppearance(): void {
        const a = this.formattingSettings?.appearanceCard;
        if (!a) return;
        const s = this.root.style;
        s.setProperty("--fc-font", a.fontFamily.value);
        s.setProperty("--fc-fontsize", `${a.fontSize.value}px`);
        s.setProperty("--fc-accent", a.accentColor.value.value);
        s.setProperty("--fc-fg", a.fontColor.value.value);
        s.setProperty("--fc-bg", a.backgroundColor.value.value);
    }
}
