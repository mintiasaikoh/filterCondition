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
import { FilterCondition, isConditionActive } from "./filterEngine";
import {
    emitAdvancedFilter,
    restoreFromAdvancedFilters,
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
        this.form.setColumns(cols);

        // 永続化された条件を初回復元
        if (!this.persistedSeen) {
            this.restoreFromPersisted(dv);
            this.persistedSeen = true;
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
        const logic = this.form.getLogic();

        const result = emitAdvancedFilter(this.host, cols, conds, logic, this.lastFilterSig);
        if (result.emitted || result.sig !== this.lastFilterSig) {
            this.lastFilterSig = result.sig;
        }
        this.persist(conds, logic);
    }

    // ==========================================================
    // 永続化
    // ==========================================================

    private persist(conds: FilterCondition[], logic: GlobalLogic): void {
        this.host.persistProperties({
            merge: [{
                objectName: "state",
                selector: null,
                properties: {
                    conditionsJson: JSON.stringify(conds),
                    logic,
                },
            }],
        });
    }

    private restoreFromPersisted(dv: DataView | null): void {
        const s = dv?.metadata?.objects?.["state"];
        if (!s) return;
        const json = String(s["conditionsJson"] ?? "");
        const logic: GlobalLogic = (s["logic"] === "OR" ? "OR" : "AND");
        if (!json) return;
        try {
            const parsed = JSON.parse(json) as unknown;
            if (!Array.isArray(parsed)) return;
            const conds: FilterCondition[] = [];
            for (const raw of parsed) {
                if (!raw || typeof raw !== "object") continue;
                const r = raw as Record<string, unknown>;
                const ci = Number(r.columnIndex);
                const op = r.operator;
                const val = String(r.value ?? "");
                if (!Number.isFinite(ci)) continue;
                if (op !== "contains" && op !== "notContains") continue;
                conds.push({ columnIndex: ci, operator: op, value: val });
            }
            this.form.setState(conds, logic);
        } catch {
            // ignore broken persisted state
        }
    }

    // ==========================================================
    // jsonFilters 受信
    // ==========================================================

    private restoreFromJsonFilters(
        jsonFilters: powerbi.IFilter[] | undefined,
        cols: powerbi.DataViewMetadataColumn[],
    ): void {
        const restored = restoreFromAdvancedFilters(jsonFilters, cols);
        if (!restored) return;

        // 自己発火エコーは skip
        if (restored.sig === this.lastFilterSig) return;

        // UI 未タッチの場合のみ上書き
        const current = this.form.getConditions();
        const uiIsEmpty = current.length === 0 || current.every(c => !isConditionActive(c));
        if (uiIsEmpty) {
            this.form.setState(restored.conditions, restored.logic);
        }
        this.lastFilterSig = restored.sig;
    }

    // ==========================================================

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
