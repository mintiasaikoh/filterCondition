"use strict";

import powerbi from "powerbi-visuals-api";
import { IFilterColumnTarget } from "powerbi-models";

// ==========================================================
// 共有型
// ==========================================================

export type FilterOp = "contains" | "notContains" | "gte" | "lte";

export interface FilterCondition {
    columnIndex: number;
    operator: FilterOp;
    value: string;
}

// ==========================================================
// 条件の判定・signature
// ==========================================================

export function isConditionActive(c: FilterCondition): boolean {
    return c.value.trim() !== "";
}

/** target + 条件アイテム配列の比較キー */
export function filterConditionSignature(
    target: IFilterColumnTarget,
    logic: string,
    sigItems: string[],
): string {
    const condSig = sigItems.slice().sort().join(",");
    return `${target.table}\0${target.column}\0${logic}\0${condSig}`;
}

/**
 * DataViewMetadataColumn から AdvancedFilter の target を生成。
 * 集計ラッパー "Sum(Table.Column)" 等は中身を剥がす。DAX メジャーは対象外。
 */
export function buildFilterTarget(col: powerbi.DataViewMetadataColumn): IFilterColumnTarget | null {
    if (!col?.queryName) return null;
    let qn = col.queryName;
    const aggMatch = qn.match(/^\w+\((.+)\)$/);
    const hasAgg = !!aggMatch;
    if (hasAgg) qn = aggMatch[1];
    if (!hasAgg && col.isMeasure) return null;
    const dotIdx = qn.indexOf(".");
    if (dotIdx < 1) return null;
    return { table: qn.substring(0, dotIdx), column: qn.substring(dotIdx + 1) };
}
