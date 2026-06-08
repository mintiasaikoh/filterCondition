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
    private persistedSeen = false;
    private lastStateSig = "";
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

        // ブックマーク対応: ブックマークはカスタムビジュアルの metadata.objects.state
        // を復元するが、適用済み AdvancedFilter を filter 層から必ずしも消さない。
        // そこで metadata.state の外部変化を検知し、UI を復元したうえで「現状態で
        // フィルタを再発火」して filter 層を同期する（空条件なら remove＝リセット成立）。
        const stateSig = this.computeStateSig(dv);
        if (!this.persistedSeen) {
            this.persistedSeen = true;
            this.lastStateSig = stateSig;
        } else if (stateSig !== this.lastStateSig) {
            // 自分の persist では無い外部変化（＝ブックマーク）
            this.lastStateSig = stateSig;
            this.restoreFromPersisted(dv);
            this.emitCurrent();   // filter 層を UI に合わせて同期（空なら除去）
            this.clearPending();
            return;               // jsonFilters 復元は次サイクルに委ねる
        }

        // 外部 jsonFilters からの復元（スライサー同期）。
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
        // 「列」(values) を先、「候補列」(categories) を後に並べる
        // （候補列がドロップダウン先頭に来てしまうのを防ぐ）
        for (const val of dv?.categorical?.values ?? []) add(val?.source);
        for (const cat of dv?.categorical?.categories ?? []) add(cat?.source);
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
        // 自分の persist は外部変化として誤検知しないよう sig を更新
        this.lastStateSig = this.makeStateSig(conds, columnLogic);

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

    /** 現在のフォーム状態で filter を再発火（ブックマーク復元時の同期用、スピナー無し） */
    private emitCurrent(): void {
        const cols = (this.lastColsRef as powerbi.DataViewMetadataColumn[] | null) ?? [];
        const conds = this.form.getConditions();
        const columnLogic = this.form.getColumnLogic();
        const result = emitAdvancedFilter(this.host, cols, conds, columnLogic, this.lastFilterSig);
        if (result.emitted || result.sig !== this.lastFilterSig) {
            this.lastFilterSig = result.sig;
        }
    }

    /** metadata.objects.state の状態シグネチャ（外部=ブックマーク書き換え検知用） */
    private computeStateSig(dv: DataView | null): string {
        const s = dv?.metadata?.objects?.["state"];
        if (!s) return "";
        const c = String(s["conditionsJson"] ?? "");
        const l = String(s["columnLogicJson"] ?? s["logic"] ?? "");
        return `${c}\0${l}`;
    }

    /** 自分が書き込む state のシグネチャ（computeStateSig と一致する形式） */
    private makeStateSig(conds: FilterCondition[], columnLogic: ColumnLogic): string {
        return `${JSON.stringify(conds)}\0${JSON.stringify(columnLogic)}`;
    }

    /** metadata.objects.state から UI を復元（空ならデフォルト1行へ）。ブックマーク復元用 */
    private restoreFromPersisted(dv: DataView | null): void {
        const s = dv?.metadata?.objects?.["state"];
        const json = s ? String(s["conditionsJson"] ?? "") : "";
        const colLogicJson = s ? String(s["columnLogicJson"] ?? "") : "";

        const conds: FilterCondition[] = [];
        if (json) {
            try {
                const parsed = JSON.parse(json) as unknown;
                if (Array.isArray(parsed)) {
                    for (const raw of parsed) {
                        if (!raw || typeof raw !== "object") continue;
                        const r = raw as Record<string, unknown>;
                        const ci = Number(r.columnIndex);
                        const op = r.operator;
                        const val = String(r.value ?? "");
                        if (!Number.isFinite(ci)) continue;
                        if (op !== "contains" && op !== "notContains" && op !== "gte" && op !== "lte") continue;
                        conds.push({ columnIndex: ci, operator: op, value: val });
                    }
                }
            } catch { /* ignore */ }
        }

        const columnLogic: ColumnLogic = {};
        if (colLogicJson) {
            try {
                const parsed = JSON.parse(colLogicJson) as unknown;
                if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
                    for (const [k, v] of Object.entries(parsed as Record<string, unknown>)) {
                        if (v === "OR" || v === "AND") columnLogic[k] = v;
                    }
                }
            } catch { /* ignore */ }
        }

        if (conds.length === 0) {
            this.form.resetToDefault();
        } else {
            this.form.setState(conds, columnLogic);
        }
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
