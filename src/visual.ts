"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;

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
        const dv = options.dataViews?.[0];
        this.lastDataView = dv ?? null;

        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, dv);
        this.applyAppearance();

        const cols = dv?.table?.columns ?? [];
        const targetName = String(this.formattingSettings?.suggestionsCard?.targetColumnName?.value ?? "").trim();
        const uniques = this.extractUniques(dv, targetName);
        this.form.setColumns(cols, uniques);

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
        this.restoreFromJsonFilters(options.jsonFilters, cols);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // ==========================================================

    private onFormChange(): void {
        const dv = this.lastDataView;
        const cols = dv?.table?.columns ?? [];
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
    ): void {
        const restored = restoreFromAdvancedFilters(jsonFilters, cols);

        // ブックマーク / 外部スライサーで自分の filter が解除された場合は UI をリセット
        if (!restored) {
            if (this.lastFilterSig !== "") {
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

    private extractUniques(dv: DataView | null, targetName: string): string[][] {
        const cols = dv?.table?.columns ?? [];
        const rows = dv?.table?.rows ?? [];
        const LIMIT = 15;
        const targets = new Set(
            targetName.split(",").map(s => s.trim()).filter(s => s.length > 0)
        );
        if (targets.size === 0) return cols.map(() => []);
        return cols.map((c, ci) => {
            if (!targets.has(c?.displayName ?? "")) return [];
            // 全行で出現回数を数え、頻度の少ない順に LIMIT 件
            const counts = new Map<string, number>();
            for (const r of rows) {
                const v = r[ci];
                if (v == null) continue;
                const s = String(v);
                if (s === "") continue;
                if (s.toUpperCase().includes("TEST")) continue;
                if (s.includes("ダミー")) continue;
                counts.set(s, (counts.get(s) ?? 0) + 1);
            }
            return Array.from(counts.entries())
                .sort((a, b) => a[1] - b[1] || a[0].localeCompare(b[0]))
                .slice(0, LIMIT)
                .map(e => e[0]);
        });
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
