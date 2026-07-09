"use strict";

import powerbi from "powerbi-visuals-api";
import { IFilterColumnTarget } from "powerbi-models";

// ==========================================================
// 共有型
// ==========================================================

// 以上/以下（gte/lte）は廃止。数値レンジは標準スライサーの領分:
// 集計フィールドへの selfFilter を Power BI が拒否しビジュアルが壊れる問題、
// および数値比較 UX は標準スライサーが優れているため
export type FilterOp = "contains" | "notContains";

export interface FilterCondition {
    columnIndex: number;
    operator: FilterOp;
    value: string;
}

// ==========================================================
// 条件の判定・signature
// ==========================================================

/**
 * 演算子が列のデータ型と互換か。
 * 型不一致の AdvancedFilter を発火すると Power BI が
 * 「1つまたは複数のフィルターに問題があります」でビジュアルごと壊すため、
 * 発火前に弾く（該当行は UI 警告）。型情報が無い列は従来通り許容。
 */
export function isOperatorCompatible(
    col: powerbi.DataViewMetadataColumn,
    _op: FilterOp,
): boolean {
    const t = col?.type as {
        numeric?: boolean; integer?: boolean; dateTime?: boolean; text?: boolean; bool?: boolean;
    } | undefined;
    if (!t) return true;
    // contains / notContains は数値・日付・bool に無効
    return !(t.numeric || t.integer || t.dateTime || t.bool);
}

export function isConditionActive(c: FilterCondition): boolean {
    const v = c.value.trim();
    if (v === "") return false;
    // 廃止済み演算子（旧永続化 state の gte/lte 等）は発火させない
    if (c.operator !== "contains" && c.operator !== "notContains") return false;
    return true;
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
    // Table[Column] / 'Table'[Column] 形式（DAX 角括弧。* 始まり等の特殊列名で出る）
    const brMatch = qn.match(/^'?(.+?)'?\[(.+)\]$/);
    if (brMatch) {
        return { table: brMatch[1], column: brMatch[2] };
    }
    // Table.Column 形式
    const dotIdx = qn.indexOf(".");
    if (dotIdx < 1) return null;
    return { table: qn.substring(0, dotIdx), column: qn.substring(dotIdx + 1) };
}
