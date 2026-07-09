"use strict";

import powerbi from "powerbi-visuals-api";
import { FilterCondition, FilterOp } from "./filterEngine";
import { ColumnLogic, GlobalLogic } from "./advancedFilterEmitter";

export interface ConditionFormCallbacks {
    onChange: () => void;
    /** 適用せずとも構造が変わった時（行追加/削除/列変更）に呼ぶ。UI 状態の persist 用 */
    onEdit?: () => void;
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
    /** filter target を作れない列（DAX メジャー等）。該当行に警告表示 */
    private unfilterable = new Set<number>();

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
        this.applyBtn.onclick = () => this.triggerApply();
        footer.appendChild(this.applyBtn);
    }

    /** 適用発火。スピナーは visual 側が jsonFilters エコー受信まで点灯させる */
    private triggerApply(): void {
        this.setApplying(true);
        this.cb.onChange();
    }

    /** 「適用中…」スピナーの ON/OFF（解除は visual がフィルタ反映を検知して呼ぶ） */
    public setApplying(on: boolean): void {
        if (on) {
            this.applyBtn.classList.add("fc-applying");
            this.applyBtn.textContent = "適用中…";
        } else {
            this.applyBtn.classList.remove("fc-applying");
            this.applyBtn.textContent = "適用";
        }
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

    /** フィルタ不能列（メジャー等）を設定し、該当行に警告表示する */
    public setUnfilterable(colIdxSet: Set<number>): void {
        this.unfilterable = colIdxSet;
        this.render();
    }

    setColumns(cols: powerbi.DataViewMetadataColumn[], uniquesPerCol: string[][] = []): void {
        this.columns = cols.map((c, i) => ({
            index: i,
            label: this.cleanLabel(c, i),
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

    /**
     * 表示ラベルを queryName から綺麗な列名に導出。
     * 「列」role はメジャーウェルなので displayName が「最初の 部署」等の集計名になる。
     * queryName の集計ラッパー（First/Count/Sum...）を剥がし、Table.Col の Col 部分を使う。
     * （CLAUDE.md 方針: column 名は queryName 後半を採用、displayName はリネームでズレる）
     */
    private cleanLabel(c: powerbi.DataViewMetadataColumn, i: number): string {
        const qn = c?.queryName;
        if (qn) {
            // 集計ラッパー strip: Agg(inner) → inner
            const m = qn.match(/^\w+\((.+)\)$/);
            const inner = m ? m[1] : qn;
            // Table[Col]（DAX 形式）→ Col
            const br = inner.match(/\[([^\]]+)\]\s*$/);
            if (br) return br[1];
            // Table.Col 形式 → Col
            const dot = inner.lastIndexOf(".");
            if (dot >= 0 && dot < inner.length - 1) return inner.substring(dot + 1);
            if (inner) return inner;
        }
        // displayName フォールバック（集計接辞「最初の 〜」「〜 の合計」等を除去）
        const dn = c?.displayName ?? `列 ${i + 1}`;
        return dn
            .replace(/^(最初の|最後の|合計|平均|最小|最大|個別カウント|カウント|分散|標準偏差|中央値)\s*/, "")
            .replace(/\s*の(最初|最後|合計|平均|最小|最大|カウント|個数|個別カウント)$/, "")
            .trim() || dn;
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
        this.cb.onEdit?.();
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

        // 列ごとにグループ化して描画。同一列の 2 条件を必ず隣接させ、間に AND/OR バッジを置く。
        // 配列上で非隣接でもバッジが正しい位置に出る。元の配列 index は makeRow に渡す
        // （削除 splice・上限判定が元 index 依存のため）。
        const orderCols: number[] = [];
        const idxByCol = new Map<number, number[]>();
        this.conditions.forEach((c, idx) => {
            if (!idxByCol.has(c.columnIndex)) {
                idxByCol.set(c.columnIndex, []);
                orderCols.push(c.columnIndex);
            }
            idxByCol.get(c.columnIndex)!.push(idx);
        });
        for (const colIdx of orderCols) {
            const idxs = idxByCol.get(colIdx)!;
            idxs.forEach((idx, k) => {
                if (k === 1) this.rowsHost.appendChild(this.makeLogicBadge(colIdx));
                this.rowsHost.appendChild(this.makeRow(this.conditions[idx], idx));
            });
        }

        // 条件配列に残っていない列のロジックは破棄
        const activeCols = new Set(orderCols.map(String));
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
        if (this.unfilterable.has(cond.columnIndex)) {
            row.classList.add("fc-row-invalid");
            row.title = "この列には条件フィルターを適用できません（メジャー / 集計式のため）。モデル上の元の列を使ってください";
        }

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
            this.cb.onEdit?.();
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
        input.onchange = () => { cond.value = input.value; this.cb.onEdit?.(); };
        input.onkeydown = (e: KeyboardEvent) => {
            if (e.key === "Enter") {
                // IME 変換確定の Enter では適用しない（日本語入力対応）
                if (e.isComposing || e.keyCode === 229) return;
                cond.value = input.value;
                this.triggerApply();
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
            this.cb.onEdit?.();
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
