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
const { emitAdvancedFilter, restoreFromAdvancedFilters, syncSelfFilter } = require(build("advancedFilterEmitter.js"));
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

    check("contains 空白のみは inactive", isConditionActive({ columnIndex: 0, operator: "contains", value: "   " }) === false);
    check("contains 値ありは active", isConditionActive({ columnIndex: 0, operator: "contains", value: "営業" }) === true);
    // 以上/以下は廃止（数値レンジは標準スライサーの領分）: 旧 op が残っていても発火しない
    check("旧 gte 条件は active にならない", isConditionActive({ columnIndex: 0, operator: "gte", value: "5" }) === false);
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
        { columnIndex: 1, operator: "contains", value: "5" },
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
    // 入力値のトリム正規化（前後の空白・全角スペース・改行はフィルタ値に含めない）
    {
        const calls5 = [];
        const host5 = { applyJsonFilter: (f, o, p, a) => calls5.push({ f, o, p, a }), persistProperties: () => {} };
        emitAdvancedFilter(host5, cols, [
            { columnIndex: 0, operator: "contains", value: "  営業  " },
        ], {}, "");
        let f5 = calls5.find(c => c.p === "filter");
        check("emit: 前後スペースはトリムされる", f5 && f5.f[0].conditions[0].value === "営業",
            f5 && JSON.stringify(f5.f[0].conditions));

        calls5.length = 0;
        emitAdvancedFilter(host5, cols, [
            { columnIndex: 0, operator: "contains", value: "　営業　\n" },
        ], {}, "");
        f5 = calls5.find(c => c.p === "filter");
        check("emit: 全角スペース・改行もトリム", f5 && f5.f[0].conditions[0].value === "営業",
            f5 && JSON.stringify(f5.f[0].conditions));

        // 内側のスペースは保持（検索語を壊さない）
        calls5.length = 0;
        emitAdvancedFilter(host5, cols, [
            { columnIndex: 0, operator: "contains", value: " 東京 支店 " },
        ], {}, "");
        f5 = calls5.find(c => c.p === "filter");
        check("emit: 内側のスペースは保持", f5 && f5.f[0].conditions[0].value === "東京 支店",
            f5 && JSON.stringify(f5.f[0].conditions));
    }

    // 演算子と列型の不一致は発火しない（Power BI が filter 異常でビジュアルを壊すため）
    {
        const calls4 = [];
        const host4 = { applyJsonFilter: (f, o, p, a) => calls4.push({ f, o, p, a }), persistProperties: () => {} };
        const colsT = [
            { queryName: "T.名称", displayName: "名称", type: { text: true } },
            { queryName: "Sum(T.金額)", displayName: "金額", isMeasure: true, type: { numeric: true } },
        ];
        // 数値列に contains → dropped（Contains は数値に無効）
        let r4 = emitAdvancedFilter(host4, colsT, [
            { columnIndex: 1, operator: "contains", value: "10" },
        ], {}, "");
        check("型: 数値列に contains は dropped", r4.dropped.includes(1), JSON.stringify(r4.dropped));

        // テキスト列に contains → 正常発火
        r4 = emitAdvancedFilter(host4, colsT, [
            { columnIndex: 0, operator: "contains", value: "x" },
        ], {}, "");
        check("型: テキスト列に contains は発火", r4.dropped.length === 0, JSON.stringify(r4.dropped));
    }

    // DAX メジャー列への条件は filter target を作れない → dropped として報告される
    {
        const calls3 = [];
        const host3 = { applyJsonFilter: (f, o, p, a) => calls3.push({ f, o, p, a }), persistProperties: () => {} };
        const colsM = [
            { queryName: "T.A", displayName: "A" },
            { queryName: "T.実績合計", displayName: "実績合計", isMeasure: true }, // agg ラッパー無しのメジャー
        ];
        const r = emitAdvancedFilter(host3, colsM, [
            { columnIndex: 0, operator: "contains", value: "x" },
            { columnIndex: 1, operator: "contains", value: "1000" },
        ], {}, "");
        check("メジャー条件は dropped に列挙", Array.isArray(r.dropped) && r.dropped.includes(1),
            JSON.stringify(r.dropped));
        const fm = calls3.find(c => c.p === "filter");
        check("メジャー条件は filter に含まれない", fm && fm.f.length === 1 && fm.f[0].target.column === "A",
            fm && JSON.stringify(fm.f.map(x => x.target.column)));
    }

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

    // 復旧経路: 履歴なし（lastSig=""）でも空適用で必ず remove を発火する
    // （壊れたフィルタがレポートに永続化されているときの唯一の復旧手段）
    calls.length = 0;
    emitAdvancedFilter(host, cols, [], {}, "");
    check("復旧: 履歴なし空適用でも main remove",
        calls.some(c => c.prop === "filter" && c.f === null && c.action === 1), JSON.stringify(calls));
    check("復旧: 履歴なし空適用でも selfFilter remove",
        calls.some(c => c.prop === "selfFilter" && c.f === null && c.action === 1), JSON.stringify(calls));

    // selfFilterEnabled=false（既定 OFF の安全動作）: main は通常発火、selfFilter は常に remove
    calls.length = 0;
    emitAdvancedFilter(host, cols, conds, { "0": "OR" }, "", new Set([0]), false);
    check("selfFilter 無効: main は発火", calls.some(c => c.prop === "filter" && Array.isArray(c.f)));
    check("selfFilter 無効: selfFilter は remove",
        calls.some(c => c.prop === "selfFilter" && c.f === null && c.action === 1),
        JSON.stringify(calls.filter(c => c.prop === "selfFilter")));

    // syncSelfFilter のトグル両面（機能はトグル ON で温存されている）
    calls.length = 0;
    syncSelfFilter(host, cols, conds, { "0": "OR" }, new Set([0]), true);
    check("sync 有効: 列側条件を merge",
        calls.some(c => c.prop === "selfFilter" && Array.isArray(c.f) && c.f.length === 1 && c.f[0].target.column === "B"),
        JSON.stringify(calls));
    calls.length = 0;
    syncSelfFilter(host, cols, conds, { "0": "OR" }, new Set([0]), false);
    check("sync 無効: remove",
        calls.some(c => c.prop === "selfFilter" && c.f === null && c.action === 1),
        JSON.stringify(calls));

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

    // 前後スペース付きの条件でもカスケードが効く（トリム正規化）
    const uniqSp = v.extractUniques(dv, cols, [{ columnIndex: orgIdx, operator: "contains", value: "  営業　" }]);
    check("カスケード: 前後スペース付きでも一致",
        JSON.stringify(uniqSp[depIdx].slice().sort()) === JSON.stringify(["一課", "二課"].sort()),
        JSON.stringify(uniqSp[depIdx]));
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

// ---------------- E: 復元時の selfFilter 同期 ----------------
section("E: 復元時の selfFilter 同期");
{
    const { v, calls } = makeVisual();
    const DATA = 2;

    // レポート再オープン相当: 初回 update の jsonFilters に自分の main filter が既に載っている
    const jf = [
        {
            filterType: 0, target: { table: "T", column: "組織名" }, logicalOperator: "And",
            conditions: [{ operator: "Contains", value: "営業" }],
        },
        {
            filterType: 0, target: { table: "T", column: "部署" }, logicalOperator: "And",
            conditions: [{ operator: "Contains", value: "一" }],
        },
    ];
    v.update({ dataViews: [makeDv(undefined)], jsonFilters: jf, type: DATA });
    check("復元: UI に 2 条件", v.form.getConditions().length === 2,
        JSON.stringify(v.form.getConditions()));
    // selfFilter は既定 OFF（実験トグル）→ 復元時は残留除去の remove が発火する
    check("復元: selfFilter は既定で remove（残留 selfFilter の自動除去）",
        calls.some(c => c.prop === "selfFilter" && c.f === null && c.action === 1)
        && !calls.some(c => c.prop === "selfFilter" && Array.isArray(c.f)),
        JSON.stringify(calls.filter(c => c.prop === "selfFilter")));
    check("復元: main filter は再発火しない", !calls.some(c => c.prop === "filter"),
        JSON.stringify(calls.filter(c => c.prop === "filter")));

    // 外部クリア: フィルタ層から自分の filter が消えた → selfFilter も remove
    calls.length = 0;
    v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    check("外部クリア: UI リセット", v.form.getConditions().length === 1 && v.form.getConditions()[0].value === "",
        JSON.stringify(v.form.getConditions()));
    check("外部クリア: selfFilter remove（残留しない）",
        calls.some(c => c.prop === "selfFilter" && c.f === null && c.action === 1),
        JSON.stringify(calls));
    check("外部クリア: main は発火しない", !calls.some(c => c.prop === "filter"),
        JSON.stringify(calls.filter(c => c.prop === "filter")));
}

// ---------------- F: 未適用条件はカスケードに使わない ----------------
section("F: 未適用条件はカスケードに使わない");
{
    const { v } = makeVisual();
    const DATA = 2;
    v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    const cols = v.lastColsRef;
    const orgIdx = cols.findIndex(c => c.queryName === "T.組織名");
    const depIdx = cols.findIndex(c => c.queryName === "T.部署");

    // 入力しただけ（未適用）で update が来ても候補は絞られない
    v.form.setState([{ columnIndex: orgIdx, operator: "contains", value: "営業" }], {});
    v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    const uniq = v.lastUniquesRef;
    check("未適用 typing はカスケードに効かない（部署 6 件のまま）", uniq[depIdx].length === 6,
        JSON.stringify(uniq[depIdx]));

    // 適用したら効く
    v.onFormChange();
    v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    const uniq2 = v.lastUniquesRef;
    check("適用後はカスケードが効く（部署 2 件）", uniq2[depIdx].length === 2,
        JSON.stringify(uniq2[depIdx]));
}

// ---------------- H: フィルタ不能列の可視化 ----------------
section("H: フィルタ不能列（メジャー）の警告表示");
{
    const DATA = 2;
    const { v, element } = makeVisual();
    // 「列」に DAX メジャー（agg ラッパー無し・isMeasure）が入っているケース
    const dvM = {
        metadata: { columns: [] },
        categorical: {
            categories: [
                { source: { queryName: "T.組織名", displayName: "組織名" }, values: ["営業", "開発"] },
            ],
            values: [
                { source: { queryName: "T.実績合計", displayName: "実績合計", isMeasure: true }, values: [1, 2] },
            ],
        },
    };
    v.update({ dataViews: [dvM], jsonFilters: [], type: DATA });
    const cols = v.lastColsRef;
    const mIdx = cols.findIndex(c => c.queryName === "T.実績合計");
    v.form.setState([{ columnIndex: mIdx, operator: "contains", value: "1000" }], {});
    v.onFormChange();
    check("メジャー条件の行に警告クラスが付く", !!element.querySelector(".fc-row-invalid"),
        element.querySelector(".fc-row") ? element.querySelector(".fc-row").className : "行なし");
}

// ---------------- I: 演算子は 含む/含まない のみ ----------------
section("I: 演算子セット（以上/以下は廃止）");
{
    const DATA = 2;
    const { v, element } = makeVisual();
    v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
    const opts = element.querySelectorAll(".fc-op-sel option");
    const labels = Array.from(opts).map(o => o.textContent);
    check("演算子は 2 つだけ（含む/含まない）",
        labels.length === 2 && labels.includes("含む") && labels.includes("含まない"),
        JSON.stringify(labels));

    // 適用時に UI の入力値もトリムされる（見た目とフィルタ値の一致）
    const cols = v.lastColsRef;
    const oi = cols.findIndex(c => c.queryName === "T.組織名");
    v.form.setState([{ columnIndex: oi, operator: "contains", value: "  営業　" }], {});
    element.querySelector(".fc-apply-btn").click();
    check("適用時に UI 値がトリムされる", v.form.getConditions()[0].value === "営業",
        JSON.stringify(v.form.getConditions()));
}

// ---------------- G: UI ランタイム堅牢性 ----------------
section("G: UI ランタイム堅牢性");
{
    const DATA = 2;

    // --- IME 変換確定の Enter で誤適用しない ---
    {
        const { v, calls, element } = makeVisual();
        v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
        const input = element.querySelector(".fc-val-input");
        check("値入力欄が存在", !!input);
        input.value = "営業";
        input.dispatchEvent(new dom.window.Event("input", { bubbles: true }));
        calls.length = 0;
        // IME 変換確定の Enter（isComposing=true）
        const imeEnter = new dom.window.KeyboardEvent("keydown", { key: "Enter" });
        Object.defineProperty(imeEnter, "isComposing", { value: true });
        input.dispatchEvent(imeEnter);
        check("IME 確定 Enter では適用しない", !calls.some(c => c.prop === "filter"),
            JSON.stringify(calls.map(c => c.prop)));
        // 通常の Enter では適用する
        const plainEnter = new dom.window.KeyboardEvent("keydown", { key: "Enter" });
        input.dispatchEvent(plainEnter);
        check("通常 Enter では適用する", calls.some(c => c.prop === "filter"),
            JSON.stringify(calls.map(c => c.prop)));
    }

    // --- 内容が同じなら DOM を再構築しない（フォーカス/IME 保護） ---
    {
        const { v, element } = makeVisual();
        v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
        const inputBefore = element.querySelector(".fc-val-input");
        // 参照は毎回新しいが内容は同一の dataview（実際の PBI update と同じ）
        v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
        const inputAfter = element.querySelector(".fc-val-input");
        check("内容同一 update で入力 DOM が保持される（再構築なし）", inputBefore === inputAfter);
    }

    // --- destroy でタイマー残留しない ---
    {
        const { v } = makeVisual();
        v.update({ dataViews: [makeDv(undefined)], jsonFilters: [], type: DATA });
        const cols = v.lastColsRef;
        const oi = cols.findIndex(c => c.queryName === "T.組織名");
        v.form.setState([{ columnIndex: oi, operator: "contains", value: "営業" }], {});
        v.onFormChange();
        check("適用後フォールバックタイマー存在", v.applyFallbackTimer !== null);
        check("destroy が実装されている", typeof v.destroy === "function");
        if (typeof v.destroy === "function") {
            v.destroy();
            check("destroy でタイマー解除", v.applyFallbackTimer === null);
        } else {
            check("destroy でタイマー解除", false, "destroy 未実装");
        }
    }
}

console.log(`\n===== 結果: pass=${pass} fail=${fail} =====`);
process.exit(fail === 0 ? 0 : 1);
