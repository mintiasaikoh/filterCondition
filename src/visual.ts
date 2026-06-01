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
    GlobalLogic,
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
        // capabilities.json で 1 つのエントリに table と categorical を両方定義しているため、
        // dataViews[0] に .table と .categorical の両方が populated されている前提。
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

        // 永続化された条件を初回復元
        if (!this.persistedSeen) {
            this.restoreFromPersisted(dv);
            this.persistedSeen = true;
            this.lastStateSig = this.computeStateSig(dv);
        } else {
            // 2 回目以降: metadata.objects.state が外部（ブックマーク等）で書き換えられたら同期
            const curSig = this.computeStateSig(dv);
            if (curSig !== this.lastStateSig) {
                this.lastStateSig = curSig;
                this.restoreFromPersisted(dv);
            }
        }

        // 外部 jsonFilters からの復元（スライサー同期）
        // Data 更新時のみ未適用入力もリセット対象に（resize/style では typing を壊さない）
        const isDataUpdate = ((options.type ?? VisualUpdateType.All) & VisualUpdateType.Data) !== 0;
        this.restoreFromJsonFilters(options.jsonFilters, cols, isDataUpdate);
    }

    /**
     * 条件フォームに出す列リストを構築。
     * table.columns ∪ categorical sources を queryName で dedupe してマージする。
     * - 候補列だけにバインドした列が table.columns に出なくても categorical source 経由で必ず出る
     * - 同一列を「列」「候補列」両方に入れても重複しない
     * - table が空でも取りこぼさない
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
        for (const c of dv?.table?.columns ?? []) add(c);
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
        this.lastStateSig = this.makeStateSig(conds, columnLogic);
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

    private restoreFromPersisted(dv: DataView | null): void {
        const s = dv?.metadata?.objects?.["state"];
        if (!s) {
            // 状態が完全に空（ブックマーク等で消えた）→ UI もリセット
            this.form.resetToDefault();
            return;
        }
        const json = String(s["conditionsJson"] ?? "");
        const colLogicJson = String(s["columnLogicJson"] ?? "");
        // 旧版（単一 logic）からの後方互換
        const legacyLogic: GlobalLogic = (s["logic"] === "OR" ? "OR" : "AND");

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

        let columnLogic: ColumnLogic = {};
        if (colLogicJson) {
            try {
                const parsed = JSON.parse(colLogicJson) as unknown;
                if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
                    for (const [k, v] of Object.entries(parsed as Record<string, unknown>)) {
                        if (v === "OR" || v === "AND") columnLogic[k] = v;
                    }
                }
            } catch { /* ignore */ }
        } else if (s["logic"] !== undefined) {
            // 旧スキーマ: 全列に同じ logic を適用
            const colCount = new Map<number, number>();
            for (const c of conds) colCount.set(c.columnIndex, (colCount.get(c.columnIndex) ?? 0) + 1);
            for (const [ci, n] of colCount) if (n >= 2) columnLogic[String(ci)] = legacyLogic;
        }

        // 状態が空相当なら UI を初期化
        if (conds.length === 0) {
            this.form.resetToDefault();
            return;
        }
        this.form.setState(conds, columnLogic);
    }

    /** metadata.objects.state の状態シグネチャ（外部書き換え検知用） */
    private computeStateSig(dv: DataView | null): string {
        const s = dv?.metadata?.objects?.["state"];
        if (!s) return "";
        const c = String(s["conditionsJson"] ?? "");
        const l = String(s["columnLogicJson"] ?? s["logic"] ?? "");
        return `${c}\0${l}`;
    }

    /** 自分が書き込もうとしている state のシグネチャ */
    private makeStateSig(conds: FilterCondition[], columnLogic: ColumnLogic): string {
        return `${JSON.stringify(conds)}\0${JSON.stringify(columnLogic)}`;
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
            const shouldResetDirty = isDataUpdate && this.form.hasAnyInput();
            if (this.lastFilterSig !== "" || shouldResetDirty) {
                this.lastFilterSig = "";
                this.form.resetToDefault();
            }
            return;
        }

        // 自己発火エコーは skip
        if (restored.sig === this.lastFilterSig) return;

        // 有効な active 条件が入ってきたら UI を上書き（ブックマーク含む）
        this.form.setState(restored.conditions, restored.columnLogic);
        this.lastFilterSig = restored.sig;
    }

    // ==========================================================

    /**
     * categorical DV から distinct 値を抽出。
     * categorical には「候補列」role にバインドされた列だけが入るので、
     * 来た列すべてが候補対象。queryName で table 側 cols と突き合わせる。
     * 候補列が空なら categorical も空 → 全 col [] を返す（datalist なし）。
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

        const LIMIT = 15;
        const byQueryName = new Map<string, string[]>();
        for (const cat of categories) {
            const src = cat?.source;
            if (!src) continue;
            const seen = new Set<string>();
            const out: string[] = [];
            for (const v of cat.values ?? []) {
                if (v == null) continue;
                const s = String(v);
                if (s === "") continue;
                if (s.toUpperCase().includes("TEST")) continue;
                if (s.includes("ダミー")) continue;
                if (seen.has(s)) continue;
                seen.add(s);
                out.push(s);
                if (out.length >= LIMIT) break;
            }
            out.sort((a, b) => a.localeCompare(b));
            byQueryName.set(src.queryName ?? "", out);
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
