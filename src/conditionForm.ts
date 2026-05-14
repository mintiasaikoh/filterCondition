"use strict";

import powerbi from "powerbi-visuals-api";
import { FilterCondition, FilterOp } from "./filterEngine";
import { ColumnLogic, GlobalLogic } from "./advancedFilterEmitter";

export interface ConditionFormCallbacks {
    onChange: () => void;
}

interface ColumnOption {
    index: number;
    label: string;
}

const MAX_PER_COLUMN = 2;

export class ConditionForm {
    private root: HTMLElement;
    private rowsHost: HTMLElement;
    private addBtn: HTMLButtonElement;
    private applyBtn: HTMLButtonElement;

    private conditions: FilterCondition[] = [];
    private columnLogic: ColumnLogic = {};
    private columns: ColumnOption[] = [];
    private uniquesPerCol: string[][] = [];
    private datalistHost: HTMLElement;
    private initialized = false;

    constructor(container: HTMLElement, private cb: ConditionFormCallbacks) {
        this.root = document.createElement("div");
        this.root.className = "fc-form";
        container.appendChild(this.root);

        this.rowsHost = document.createElement("div");
        this.rowsHost.className = "fc-rows";
        this.root.appendChild(this.rowsHost);

        this.datalistHost = document.createElement("div");
        this.datalistHost.className = "fc-datalists";
        this.datalistHost.style.display = "none";
        this.root.appendChild(this.datalistHost);

        const footer = document.createElement("div");
        footer.className = "fc-footer";
        this.root.appendChild(footer);

        this.addBtn = document.createElement("button");
        this.addBtn.type = "button";
        this.addBtn.className = "fc-add-btn";
        this.addBtn.textContent = "+ 条件追加";
        this.addBtn.onclick = () => this.onAddCondition();
        footer.appendChild(this.addBtn);

        const clearBtn = document.createElement("button");
        clearBtn.type = "button";
        clearBtn.className = "fc-clear-btn";
        clearBtn.textContent = "クリア";
        clearBtn.onclick = () => this.onClearAll();
        footer.appendChild(clearBtn);

        this.applyBtn = document.createElement("button");
        this.applyBtn.type = "button";
        this.applyBtn.className = "fc-apply-btn";
        this.applyBtn.textContent = "適用";
        this.applyBtn.onclick = () => this.cb.onChange();
        footer.appendChild(this.applyBtn);
    }

    private onClearAll(): void {
        this.resetToDefault();
        this.cb.onChange();
    }

    /** UI を「1 行空の条件」状態に戻す（発火しない）。ブックマーク等の外部 filter 解除からも呼ばれる */
    public resetToDefault(): void {
        if (this.columns.length > 0) {
            this.conditions = [{
                columnIndex: this.columns[0].index,
                operator: "contains",
                value: "",
            }];
        } else {
            this.conditions = [];
        }
        this.columnLogic = {};
        this.render();
    }

    /** 入力（適用済みかどうか問わず）が残っているか */
    public hasAnyInput(): boolean {
        return this.conditions.some(c => c.value.trim() !== "");
    }

    setColumns(cols: powerbi.DataViewMetadataColumn[], uniquesPerCol: string[][] = []): void {
        this.columns = cols.map((c, i) => ({
            index: i,
            label: c?.displayName ?? `列 ${i + 1}`,
        }));
        this.uniquesPerCol = uniquesPerCol;
        this.rebuildDatalists();
        // 既存条件で列 index が範囲外なら削除
        const valid = this.columns.map(c => c.index);
        this.conditions = this.conditions.filter(c => valid.includes(c.columnIndex));

        // 初回列バインド時、条件がゼロならデフォルト行を 1 つ出しておく
        if (!this.initialized && this.columns.length > 0 && this.conditions.length === 0) {
            this.conditions.push({
                columnIndex: this.columns[0].index,
                operator: "contains",
                value: "",
            });
        }
        if (this.columns.length > 0) this.initialized = true;

        this.render();
    }

    setState(conditions: FilterCondition[], columnLogic: ColumnLogic): void {
        this.conditions = conditions.map(c => ({ ...c }));
        this.columnLogic = { ...columnLogic };
        this.initialized = true;
        this.render();
    }

    getConditions(): FilterCondition[] {
        return this.conditions.map(c => ({ ...c }));
    }

    getColumnLogic(): ColumnLogic {
        return { ...this.columnLogic };
    }

    // ==========================================================

    private onAddCondition(): void {
        const firstFree = this.findFreeColumn();
        if (firstFree < 0) return;
        this.conditions.push({ columnIndex: firstFree, operator: "contains", value: "" });
        this.render();
    }

    private findFreeColumn(): number {
        const count = new Map<number, number>();
        for (const c of this.conditions) count.set(c.columnIndex, (count.get(c.columnIndex) ?? 0) + 1);
        for (const col of this.columns) {
            if ((count.get(col.index) ?? 0) < MAX_PER_COLUMN) return col.index;
        }
        return -1;
    }

    private render(): void {
        while (this.rowsHost.firstChild) this.rowsHost.removeChild(this.rowsHost.firstChild);

        if (this.columns.length === 0) {
            const empty = document.createElement("div");
            empty.className = "fc-empty";
            empty.textContent = "列を「列」フィールドにバインドしてください";
            this.rowsHost.appendChild(empty);
            this.addBtn.disabled = true;
            return;
        }

        this.addBtn.disabled = this.findFreeColumn() < 0;

        if (this.conditions.length === 0) {
            const hint = document.createElement("div");
            hint.className = "fc-hint";
            hint.textContent = "「+ 条件追加」で絞り込み条件を作成";
            this.rowsHost.appendChild(hint);
            return;
        }

        // 同じ列の 2 行目には AND/OR バッジを前置
        const seen = new Set<number>();
        this.conditions.forEach((c, idx) => {
            if (seen.has(c.columnIndex)) {
                this.rowsHost.appendChild(this.makeLogicBadge(c.columnIndex));
            }
            seen.add(c.columnIndex);
            this.rowsHost.appendChild(this.makeRow(c, idx));
        });

        // 条件配列に残っていない列のロジックは破棄
        const activeCols = new Set(this.conditions.map(c => c.columnIndex).map(String));
        for (const k of Object.keys(this.columnLogic)) {
            if (!activeCols.has(k)) delete this.columnLogic[k];
        }
    }

    private makeLogicBadge(colIdx: number): HTMLElement {
        const wrap = document.createElement("div");
        wrap.className = "fc-logic-badge";
        const sel = document.createElement("select");
        sel.className = "fc-logic-sel";
        for (const v of ["AND", "OR"] as GlobalLogic[]) {
            const opt = document.createElement("option");
            opt.value = v;
            opt.textContent = v;
            sel.appendChild(opt);
        }
        sel.value = this.columnLogic[String(colIdx)] === "OR" ? "OR" : "AND";
        sel.onchange = () => {
            this.columnLogic[String(colIdx)] = sel.value === "OR" ? "OR" : "AND";
        };
        wrap.appendChild(sel);
        return wrap;
    }

    private makeRow(cond: FilterCondition, idx: number): HTMLElement {
        const row = document.createElement("div");
        row.className = "fc-row";

        // 列セレクタ
        const colSel = document.createElement("select");
        colSel.className = "fc-col-sel";
        const usage = this.colUsageExcluding(idx);
        for (const co of this.columns) {
            const opt = document.createElement("option");
            opt.value = String(co.index);
            const used = usage.get(co.index) ?? 0;
            const full = used >= MAX_PER_COLUMN && co.index !== cond.columnIndex;
            opt.disabled = full;
            opt.textContent = full ? `${co.label}（上限）` : co.label;
            if (co.index === cond.columnIndex) opt.selected = true;
            colSel.appendChild(opt);
        }
        colSel.onchange = () => {
            cond.columnIndex = parseInt(colSel.value, 10);
            this.render();
        };
        row.appendChild(colSel);

        // 演算子
        const opSel = document.createElement("select");
        opSel.className = "fc-op-sel";
        for (const [v, label] of [
            ["contains", "含む"],
            ["notContains", "含まない"],
            ["gte", "以上"],
            ["lte", "以下"],
        ] as [FilterOp, string][]) {
            const opt = document.createElement("option");
            opt.value = v;
            opt.textContent = label;
            if (v === cond.operator) opt.selected = true;
            opSel.appendChild(opt);
        }
        opSel.onchange = () => {
            cond.operator = opSel.value as FilterOp;
        };
        row.appendChild(opSel);

        // 値入力
        const input = document.createElement("input");
        input.type = "text";
        input.className = "fc-val-input";
        input.placeholder = "値を入力";
        input.value = cond.value;
        const listId = this.datalistIdFor(cond.columnIndex);
        if (listId) input.setAttribute("list", listId);
        input.oninput = () => { cond.value = input.value; };
        input.onchange = () => { cond.value = input.value; };
        input.onkeydown = (e: KeyboardEvent) => {
            if (e.key === "Enter") {
                cond.value = input.value;
                this.cb.onChange();
            }
        };
        row.appendChild(input);

        // 削除
        const del = document.createElement("button");
        del.type = "button";
        del.className = "fc-del-btn";
        del.title = "条件を削除";
        del.textContent = "×";
        del.onclick = () => {
            this.conditions.splice(idx, 1);
            this.render();
        };
        row.appendChild(del);

        return row;
    }

    private datalistIdFor(colIdx: number): string | null {
        const uniques = this.uniquesPerCol[colIdx];
        if (!uniques || uniques.length === 0) return null;
        return `fc-vals-${colIdx}`;
    }

    private rebuildDatalists(): void {
        while (this.datalistHost.firstChild) this.datalistHost.removeChild(this.datalistHost.firstChild);
        this.uniquesPerCol.forEach((vals, ci) => {
            if (!vals || vals.length === 0) return;
            const dl = document.createElement("datalist");
            dl.id = `fc-vals-${ci}`;
            for (const v of vals) {
                const opt = document.createElement("option");
                opt.value = v;
                dl.appendChild(opt);
            }
            this.datalistHost.appendChild(dl);
        });
    }

    private colUsageExcluding(excludeIdx: number): Map<number, number> {
        const m = new Map<number, number>();
        this.conditions.forEach((c, i) => {
            if (i === excludeIdx) return;
            m.set(c.columnIndex, (m.get(c.columnIndex) ?? 0) + 1);
        });
        return m;
    }
}
