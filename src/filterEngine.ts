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

/**
 * 数値入力の正規化パース。通貨列向けの日本語入力を許容する:
 * 全角数字・記号 → 半角、カンマ区切り・通貨記号（¥/￥/$/＄）・空白を除去。
 * 数値にならなければ null。
 */
export function parseNumericFilterValue(raw: string): number | null {
    let s = raw.trim();
    if (s === "") return null;
    s = s
        .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
        .replace(/[．]/g, ".")
        .replace(/[－ー−]/g, "-")
        .replace(/[，、]/g, ",")
        .replace(/[¥￥$＄\s,]/g, "");
    if (s === "") return null;
    const n = Number(s);
    return Number.isFinite(n) ? n : null;
}

/**
 * 演算子が列のデータ型と互換か。
 * 型不一致の AdvancedFilter を発火すると Power BI が
 * 「1つまたは複数のフィルターに問題があります」でビジュアルごと壊すため、
 * 発火前に弾く（該当行は UI 警告）。型情報が無い列は従来通り許容。
 */
export function isOperatorCompatible(
    col: powerbi.DataViewMetadataColumn,
    op: FilterOp,
): boolean {
    const t = col?.type as {
        numeric?: boolean; integer?: boolean; dateTime?: boolean; text?: boolean; bool?: boolean;
    } | undefined;
    if (!t) return true;
    if (op === "gte" || op === "lte") {
        return !!(t.numeric || t.integer);
    }
    // contains / notContains は数値・日付・bool に無効
    return !(t.numeric || t.integer || t.dateTime || t.bool);
}

export function isConditionActive(c: FilterCondition): boolean {
    const v = c.value.trim();
    if (v === "") return false;
    // gte/lte は数値のみ有効（非数値は黙ってマッチ無しになるので発火させない）。
    // カンマ・全角・通貨記号は parseNumericFilterValue が吸収する
    if (c.operator === "gte" || c.operator === "lte") {
        return parseNumericFilterValue(v) !== null;
    }
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
