"use strict";
/* filterCondition ロジックテスト（Power BI 実機なし）
 *
 * 実行: npm test（tsc で ../src を test/build にコンパイル後、本ファイルを実行）
 *
 * カバー範囲: filterEngine / advancedFilterEmitter / extractUniques カスケード /
 * update フロー（stale ガード・ブックマークリセット・スピナー）
 * カバー外: Power BI ホストの実挙動（selfFilter の再クエリ、bookmark の
 * metadata 書き換え形式、実データの queryName 形状）→ 実機 smoke test が必要
 */
const path = require("path");
const Module = require("module");

// --- モジュールフック: less は空、formattingmodel はスタブ ---
const origLoad = Module._load;
Module._load = function (request, parent, isMain) {
    if (request.endsWith(".less")) return {};
    if (request === "powerbi-visuals-utils-formattingmodel") {
        return require("./stubs/formattingmodel.js");
    }
    return origLoad.call(this, request, parent, isMain);
};

// --- jsdom で DOM グローバル ---
const { JSDOM } = require("jsdom");
const dom = new JSDOM("<!doctype html><html><body></body></html>");
global.window = dom.window;
global.document = dom.window.document;
global.HTMLElement = dom.window.HTMLElement;
global.KeyboardEvent = dom.window.KeyboardEvent;

let pass = 0, fail = 0;
function check(name, cond, detail) {
    if (cond) { pass++; console.log(`  ok   ${name}`); }
    else { fail++; console.log(`  FAIL ${name}${detail ? " :: " + detail : ""}`); }
}
function section(t) { console.log(`\n== ${t} ==`); }

const build = (p) => path.join(__dirname, "build", "src", p);
const { buildFilterTarget, isConditionActive } = require(build("filterEngine.js"));
const { emitAdvancedFilter, restoreFromAdvancedFilters } = require(build("advancedFilterEmitter.js"));
const { Visual } = require(build("visual.js"));

// ---------------- A: filterEngine ----------------
section("A: filterEngine");
{
    let t = buildFilterTarget({ queryName: "T.Col" });
    check("dot 形式", t && t.table === "T" && t.column === "Col", JSON.stringify(t));

    t = buildFilterTarget({ queryName: "Sum(T.Col)", isMeasure: true });
    check("集計ラッパー strip", t && t.table === "T" && t.column === "Col", JSON.stringify(t));

    t = buildFilterTarget({ queryName: "T[＊決裁していただきたい事項]" });
    check("角括弧 形式（＊始まり）", t && t.table === "T" && t.column === "＊決裁していただきたい事項", JSON.stringify(t));

    t = buildFilterTarget({ queryName: "First(T[＊決裁していただきたい事項])", isMeasure: true });
    check("集計+角括弧", t && t.table === "T" && t.column === "＊決裁していただきたい事項", JSON.stringify(t));

    t = buildFilterTarget({ queryName: "'My Table'[Col]" });
    check("クォート付きテーブル", t && t.table === "My Table" && t.column === "Col", JSON.stringify(t));

    t = buildFilterTarget({ queryName: "[MyMeasure]", isMeasure: true });
    check("DAX メジャーは null", t === null, JSON.stringify(t));

    check("gte 非数値は inactive", isConditionActive({ columnIndex: 0, operator: "gte", value: "abc" }) === false);
    check("gte 数値は active", isConditionActive({ columnIndex: 0, operator: "gte", value: "5" }) === true);
    check("contains 空白のみは inactive", isConditionActive({ columnIndex: 0, operator: "contains", value: "   " }) === false);
}

// ---------------- B: emit / restore 往復 ----------------
section("B: emit/restore（selfFilter・dedupe・echo）");
{
    const calls = [];
    const host = {
        applyJsonFilter: (f, obj, prop, action) => calls.push({ f, obj, prop, action }),
        persistProperties: () => {},
    };
    const cols = [
        { queryName: "T.A", displayName: "A" },
        { queryName: "Sum(T.B)", displayName: "B の合計", isMeasure: true },
    ];
    const conds = [
        { columnIndex: 0, operator: "contains", value: "foo" },
        { columnIndex: 0, operator: "notContains", value: "bar" },
        { columnIndex: 1, operator: "gte", value: "5" },
    ];
    const result = emitAdvancedFilter(host, cols, conds, { "0": "OR" }, "");

    check("emitted=true / sig 非空", result.emitted === true && result.sig.length > 0);
    const mainCalls = calls.filter(c => c.prop === "filter");
    const selfCalls = calls.filter(c => c.prop === "selfFilter");
    check("main filter が 1 回発火", mainCalls.length === 1);
    check("selfFilter も 1 回発火", selfCalls.length === 1);
    check("main と selfFilter が同一内容", JSON.stringify(mainCalls[0]?.f) === JSON.stringify(selfCalls[0]?.f));
    const filters = mainCalls[0].f;
    check("列ごとに 2 フィルタ", Array.isArray(filters) && filters.length === 2);
    const colA = filters.find(f => f.target.column === "A");
    check("col0 の logic=Or", colA && colA.logicalOperator === "Or", colA && colA.logicalOperator);
    const colB = filters.find(f => f.target.column === "B");
    check("gte 値が number 化", colB && typeof colB.conditions[0].value === "number");

    // echo: filter + selfFilter が両方 jsonFilters に載って戻るケース
    const echoed = [...filters, ...filters].map(f => JSON.parse(JSON.stringify(f)));
    echoed.forEach(f => { f.filterType = 0; }); // FilterType.Advanced = 0
    const restored = restoreFromAdvancedFilters(echoed, cols);
    check("restore 成功", restored !== null);
    check("dedupe: 条件 3 件（重複なし）", restored && restored.conditions.length === 3,
        restored && JSON.stringify(restored.conditions));
    check("echo sig 一致（無限ループしない）", restored && restored.sig === result.sig,
        restored && `${restored.sig} vs ${result.sig}`);
    check("columnLogic 復元 OR", restored && restored.columnLogic["0"] === "OR");

    // クリア: main + selfFilter 両方 remove
    calls.length = 0;
    const cleared = emitAdvancedFilter(host, cols, [], {}, result.sig);
    check("クリアで emitted", cleared.emitted === true && cleared.sig === "");
    check("main remove", calls.some(c => c.prop === "filter" && c.action === 1 && c.f === null));
    check("selfFilter remove", calls.some(c => c.prop === "selfFilter" && c.action === 1 && c.f === null));

    // --- selfFilter の候補列条件除外（自列候補を殺さないため） ---
    // col0 が候補列の場合: main は全条件、selfFilter は列(メジャー)側の条件のみ
    calls.length = 0;
    emitAdvancedFilter(host, cols, conds, { "0": "OR" }, "", new Set([0]));
    const mainSplit = calls.find(c => c.prop === "filter");
    const selfSplit = calls.find(c => c.prop === "selfFilter");
    check("split: main は全列（2 フィルタ）", mainSplit && Array.isArray(mainSplit.f) && mainSplit.f.length === 2,
        mainSplit && JSON.stringify(mainSplit.f?.map(f => f.target.column)));
    check("split: selfFilter は列側のみ（B のみ）", selfSplit && Array.isArray(selfSplit.f)
        && selfSplit.f.length === 1 && selfSplit.f[0].target.column === "B",
        selfSplit && JSON.stringify(Array.isArray(selfSplit.f) ? selfSplit.f.map(f => f.target.column) : selfSplit.f));

    // 候補列の条件しか無い場合 → selfFilter は remove（サーバー側では一切絞らない）
    calls.length = 0;
    emitAdvancedFilter(host, cols, [conds[0], conds[1]], { "0": "OR" }, "", new Set([0]));
    check("split: 候補列条件のみ → selfFilter remove",
        calls.some(c => c.prop === "selfFilter" && c.f === null && c.action === 1),
        JSON.stringify(calls.filter(c => c.prop === "selfFilter")));
}

// ---------------- ヘルパ ----------------
function makeVisual() {
    const calls = [];
    const persists = [];
    const host = {
        applyJsonFilter: (f, obj, prop, action) => calls.push({ f, obj, prop, action }),
        persistProperties: (p) => persists.push(p),
    };
    const element = document.createElement("div");
    const v = new Visual({ host, element });
    return { v, host, calls, persists, element };
}

function makeDv(stateObjects) {
    // 候補列: 組織名 × 部署（タプル整列）、列: Sum(T.金額) メジャー
    return {
        metadata: { columns: [], objects: stateObjects },
        categorical: {
            categories: [
                {
                    source: { queryName: "T.組織名", displayName: "組織名" },
                    values: ["営業", "営業", "開発", "TEST部", "ダミー課", "総務"],
                },
                {
                    source: { queryName: "T.部署", displayName: "部署" },
                    values: ["一課", "二課", "三課", "四課", "五課", "六課"],
                },
            ],
            values: [
                { source: { queryName: "Sum(T.金額)", displayName: "金額 の合計", isMeasure: true }, values: [1, 2, 3, 4, 5, 6] },
            ],
        },
    };
}

// ---------------- C: 候補カスケード ----------------
section("C: extractUniques カスケード");
{
    const { v } = makeVisual();
    const dv = makeDv(undefined);
    const cols = v.resolveColumns(dv);

    check("resolveColumns: 列(values)が先頭", cols[0].queryName === "Sum(T.金額)", cols.map(c => c.queryName).join(","));
    check("resolveColumns: 3 列", cols.length === 3);

    let uniq = v.extractUniques(dv, cols, []);
    const orgIdx = cols.findIndex(c => c.queryName === "T.組織名");
    const depIdx = cols.findIndex(c => c.queryName === "T.部署");
    check("組織名候補: TEST/ダミー除外で 3 件", JSON.stringify(uniq[orgIdx].slice().sort()) === JSON.stringify(["営業", "総務", "開発"].sort()),
        JSON.stringify(uniq[orgIdx]));
    check("部署候補: 6 件", uniq[depIdx].length === 6, JSON.stringify(uniq[depIdx]));
    check("メジャー列は候補なし", uniq[0].length === 0);

    uniq = v.extractUniques(dv, cols, [{ columnIndex: orgIdx, operator: "contains", value: "営業" }]);
    check("カスケード: 部署が営業の 2 件に減る", JSON.stringify(uniq[depIdx].slice().sort()) === JSON.stringify(["一課", "二課"].sort()),
        JSON.stringify(uniq[depIdx]));
    check("カスケード: 自列の選択肢は残る", uniq[orgIdx].length === 3, JSON.stringify(uniq[orgIdx]));

    const again = v.extractUniques(dv, cols, [{ columnIndex: orgIdx, operator: "contains", value: "営業" }]);
    check("キャッシュヒット（同一参照）", again === uniq);
}

// ---------------- D: update フロー ----------------
section("D: update フロー");
{
    const { v, calls, persists } = makeVisual();
    const DATA = 2; // VisualUpdateType.Data

    v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    const form = v.form;
    check("初回: デフォルト 1 行", form.getConditions().length === 1 && form.getConditions()[0].value === "");

    const cols = v.lastColsRef;
    const orgIdx = cols.findIndex(c => c.queryName === "T.組織名");
    form.setState([{ columnIndex: orgIdx, operator: "contains", value: "営業" }], {});
    v.onFormChange();
    check("適用: main+selfFilter 発火", calls.filter(c => c.prop === "filter").length === 1 && calls.filter(c => c.prop === "selfFilter").length === 1);
    // 組織名は候補列 → selfFilter に自列条件を載せない（remove になる）
    const selfCall = calls.find(c => c.prop === "selfFilter");
    check("適用: selfFilter は候補列条件を含まない", selfCall && selfCall.f === null && selfCall.action === 1,
        selfCall && JSON.stringify(Array.isArray(selfCall.f) ? selfCall.f.map(f => f.target.column) : selfCall.f));
    check("適用: persist 実行", persists.length === 1);
    check("適用: スピナー点灯 (pendingApply)", v.pendingApply === true);
    const emittedFilters = calls.find(c => c.prop === "filter").f;
    const sigAfterApply = v.lastFilterSig;

    // stale metadata（persist 未反映）+ 自己エコー
    const echo = emittedFilters.map(f => { const o = JSON.parse(JSON.stringify(f)); o.filterType = 0; return o; });
    v.update({ dataViews: [makeDv(undefined)], jsonFilters: echo, type: DATA });
    check("stale ガード: 入力が消えない", v.form.getConditions()[0]?.value === "営業",
        JSON.stringify(v.form.getConditions()));
    check("エコーでスピナー解除", v.pendingApply === false);
    check("エコーで sig 不変", v.lastFilterSig === sigAfterApply);

    // 自分の persist 反映 → awaitingPersist 解除
    const myState = {
        state: {
            conditionsJson: JSON.stringify(v.form.getConditions()),
            columnLogicJson: JSON.stringify(v.form.getColumnLogic()),
        },
    };
    v.update({ dataViews: [makeDv(myState)], jsonFilters: echo, type: DATA });
    check("persist 反映で awaitingPersist 解除", v.awaitingPersist === false);
    check("入力保持", v.form.getConditions()[0]?.value === "営業");

    // ブックマーク（外部が state を空条件へ、filter 層は残ったまま）
    calls.length = 0;
    const bookmarkState = { state: { conditionsJson: "[]", columnLogicJson: "{}" } };
    v.update({ dataViews: [makeDv(bookmarkState)], jsonFilters: echo, type: DATA });
    check("ブックマーク: UI がデフォルトへリセット", v.form.getConditions().length === 1 && v.form.getConditions()[0].value === "",
        JSON.stringify(v.form.getConditions()));
    check("ブックマーク: filter remove 発火（emitCurrent）", calls.some(c => c.prop === "filter" && c.f === null && c.action === 1));
    check("ブックマーク: selfFilter も remove", calls.some(c => c.prop === "selfFilter" && c.f === null && c.action === 1));
    check("ブックマーク: lastFilterSig クリア", v.lastFilterSig === "");

    // 条件追加で入力が消えない
    const { v: v2, persists: p2 } = makeVisual();
    v2.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    const cols2 = v2.lastColsRef;
    const oi2 = cols2.findIndex(c => c.queryName === "T.組織名");
    v2.form.setState([{ columnIndex: oi2, operator: "contains", value: "未適用ワード" }], {});
    v2.form.onAddCondition();
    check("条件追加: 入力ワード保持", v2.form.getConditions()[0].value === "未適用ワード",
        JSON.stringify(v2.form.getConditions()));
    check("条件追加: 2 行になる", v2.form.getConditions().length === 2);
    check("条件追加: persistUiState 実行", p2.length >= 1);
    v2.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    check("未適用+stale update: 入力保持", v2.form.getConditions()[0]?.value === "未適用ワード",
        JSON.stringify(v2.form.getConditions()));
}

console.log(`\n===== 結果: pass=${pass} fail=${fail} =====`);
process.exit(fail === 0 ? 0 : 1);
