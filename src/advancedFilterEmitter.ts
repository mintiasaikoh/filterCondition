"use strict";

import powerbi from "powerbi-visuals-api";
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import FilterAction = powerbi.FilterAction;

import {
    AdvancedFilter,
    IAdvancedFilter,
    IAdvancedFilterCondition,
    IFilterColumnTarget,
    FilterType,
    AdvancedFilterLogicalOperators,
    AdvancedFilterConditionOperators,
} from "powerbi-models";

import {
    FilterCondition,
    FilterOp,
    isConditionActive,
    buildFilterTarget,
    filterConditionSignature,
} from "./filterEngine";

export type GlobalLogic = "AND" | "OR";

const SIG_PREFIX = "ADV|";

export interface EmitResult {
    sig: string;       // 発火した signature（""=発火せず / remove 時）
    emitted: boolean;  // 実際に applyJsonFilter を呼んだか
}

/**
 * 条件から AdvancedFilter を組み立てて発火。
 * 列内の logicalOperator は globalLogic に従う（列間は Power BI が暗黙 AND で結合）。
 */
export function emitAdvancedFilter(
    host: IVisualHost,
    cols: powerbi.DataViewMetadataColumn[],
    conds: FilterCondition[],
    globalLogic: GlobalLogic,
    lastSig: string,
): EmitResult {
    const active = conds.filter(isConditionActive);

    // 条件なし → 既発火があれば remove
    if (active.length === 0) {
        if (lastSig === "") return { sig: "", emitted: false };
        host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        return { sig: "", emitted: true };
    }

    // 列ごとにグループ化（1 列最大 2 条件）
    const byCol = new Map<number, FilterCondition[]>();
    for (const c of active) {
        const arr = byCol.get(c.columnIndex) ?? [];
        if (arr.length < 2) arr.push(c);
        byCol.set(c.columnIndex, arr);
    }

    const opMap = (op: FilterOp): AdvancedFilterConditionOperators | null => {
        if (op === "contains")    return "Contains";
        if (op === "notContains") return "DoesNotContain";
        return null;
    };

    const logical: AdvancedFilterLogicalOperators = globalLogic === "OR" ? "Or" : "And";

    const filters: AdvancedFilter[] = [];
    const sigParts: string[] = [];

    for (const [ci, condList] of byCol) {
        const col = cols[ci];
        if (!col) continue;
        const target = buildFilterTarget(col);
        if (!target) continue;

        const advConds: IAdvancedFilterCondition[] = [];
        const sigItems: string[] = [];

        for (const c of condList) {
            const op = opMap(c.operator);
            if (!op) continue;
            advConds.push({ operator: op, value: c.value } as unknown as IAdvancedFilterCondition);
            sigItems.push(`${op}:${c.value}`);
        }
        if (advConds.length === 0) continue;

        filters.push(new AdvancedFilter(target, logical, ...advConds));
        sigParts.push(filterConditionSignature(target, logical, sigItems));
    }

    if (filters.length === 0) {
        if (lastSig === "") return { sig: "", emitted: false };
        host.applyJsonFilter(null, "general", "filter", FilterAction.remove);
        return { sig: "", emitted: true };
    }

    const sig = SIG_PREFIX + sigParts.slice().sort().join("|");
    if (sig === lastSig) return { sig, emitted: false };

    host.applyJsonFilter(filters, "general", "filter", FilterAction.merge);
    return { sig, emitted: true };
}

// ==========================================================
// 受信: jsonFilters → FilterCondition[] へ復元
// ==========================================================

export interface RestoredFilterState {
    conditions: FilterCondition[];
    logic: GlobalLogic;
    sig: string;
}

export function restoreFromAdvancedFilters(
    jsonFilters: powerbi.IFilter[] | undefined,
    cols: powerbi.DataViewMetadataColumn[],
): RestoredFilterState | null {
    if (!jsonFilters || jsonFilters.length === 0) return null;

    const advanced: IAdvancedFilter[] = [];
    for (const f of jsonFilters) {
        const ft = (f as unknown as { filterType?: FilterType })?.filterType;
        if (ft === FilterType.Advanced) advanced.push(f as unknown as IAdvancedFilter);
    }
    if (advanced.length === 0) return null;

    const opMapIn = (op: string): FilterOp | null => {
        if (op === "Contains")       return "contains";
        if (op === "DoesNotContain") return "notContains";
        return null;
    };

    interface RestoredCond { op: FilterOp; value: string; sigItem: string; }
    interface Restored { colIdx: number; logic: AdvancedFilterLogicalOperators; conds: RestoredCond[]; sig: string; }

    const restored: Restored[] = [];
    let globalLogic: GlobalLogic = "AND";

    for (const af of advanced) {
        const tgt = af.target as IFilterColumnTarget;
        if (!tgt || !af.conditions || af.conditions.length === 0) continue;

        let colIdx = -1;
        for (let i = 0; i < cols.length; i++) {
            const t = buildFilterTarget(cols[i]);
            if (t && t.table === tgt.table && t.column === tgt.column) { colIdx = i; break; }
        }
        if (colIdx < 0) continue;

        const condsRaw: RestoredCond[] = [];
        for (const c of af.conditions) {
            const mapped = opMapIn(String(c.operator));
            if (!mapped) continue;
            const valStr = String(c.value ?? "");
            if (valStr === "") continue;
            condsRaw.push({ op: mapped, value: valStr, sigItem: `${c.operator}:${valStr}` });
        }
        if (condsRaw.length === 0) continue;

        const kept = condsRaw.slice(0, 2);
        const logic = (af.logicalOperator || "And") as AdvancedFilterLogicalOperators;
        if (kept.length >= 2) globalLogic = logic === "Or" ? "OR" : "AND";

        restored.push({
            colIdx, logic, conds: kept,
            sig: filterConditionSignature(tgt, logic, kept.map(k => k.sigItem)),
        });
    }
    if (restored.length === 0) return null;

    const conditions: FilterCondition[] = [];
    for (const r of restored) {
        for (const c of r.conds) {
            conditions.push({ columnIndex: r.colIdx, operator: c.op, value: c.value });
        }
    }

    const sig = SIG_PREFIX + restored.map(r => r.sig).sort().join("|");
    return { conditions, logic: globalLogic, sig };
}
