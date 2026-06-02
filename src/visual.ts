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
