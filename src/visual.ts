"use strict";

import "./../style/visual.less";
import powerbi from "powerbi-visuals-api";
import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisualHost = powerbi.extensibility.visual.IVisualHost;
import DataView = powerbi.DataView;
import VisualUpdateType = powerbi.VisualUpdateType;

import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";

import { ConditionForm } from "./conditionForm";
import { FilterCondition, FilterOp, isConditionActive } from "./filterEngine";
import {
    emitAdvancedFilter,
    restoreFromAdvancedFilters,
    syncSelfFilter,
    ColumnLogic,
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
    private pendingApply = false;
    private applyFallbackTimer: number | null = null;
    private persistedSeen = false;
    private lastStateSig = "";
    private awaitingPersist = false;
    /** 適用済み条件（カスケードの根拠）。typing 中の未適用値はここに入れない */
    private appliedConds: FilterCondition[] = [];
    /**
     * remove 発火後、消したはずの旧フィルタが stale echo で戻ってきて
     * クリアを取り消してしまうのを防ぐガード。remove した sig を記憶し、
     * 同一 sig の受信は無視する（空 echo か別フィルタの受信で解除）
     */
    private removedSig = "";
    private uniquesCache: { catRef: unknown; condsSig: string; result: string[][] } | null = null;
    private lastColsRef: unknown = null;
    private lastUniquesRef: unknown = null;

    constructor(options: VisualConstructorOptions) {
        this.host = options.host;
        this.formattingSettingsService = new FormattingSettingsService();

        this.root = document.createElement("div");
        this.root.className = "fc-visual";
        options.element.appendChild(this.root);

        this.form = new ConditionForm(this.root, {
            onChange: () => this.onFormChange(),
            onEdit: () => this.persistUiState(),
        });
    }

    public update(options: VisualUpdateOptions): void {
        // Power BI は 1 ビジュアル 1 クエリ。mapping は categorical 一本
        // （categories=候補列, values=列メジャー）。dataViews[0].categorical を使う。
        const dvs = options.dataViews ?? [];
        const dv = dvs[0] ?? null;
        this.lastDataView = dv;

        this.formattingSettings = this.formattingSettingsService
            .populateFormattingSettingsModel(VisualFormattingSettingsModel, dv);
        this.applyAppearance();

        // cols / uniques は「内容が同じなら前回と同じ参照」に安定化する。
        // update は persist やフィルタ操作のたびに来るため、毎回 DOM を再構築すると
        // 入力中のフォーカス / IME 状態が破壊される。内容比較で参照を保てば
        // 下の参照チェックで setColumns（全再描画）がスキップされる。
        const cols = this.stableCols(this.resolveColumns(dv));
        // カスケードは「適用済み」条件のみ根拠にする（typing 中の値で候補を絞らない）
        const uniques = this.extractUniques(dv, cols, this.appliedConds);
        if (cols !== this.lastColsRef || uniques !== this.lastUniquesRef) {
            this.lastColsRef = cols;
            this.lastUniquesRef = uniques;
            this.form.setColumns(cols, uniques);
        }

        // ブックマーク対応: ブックマークはカスタムビジュアルの metadata.objects.state
        // を復元するが、適用済み AdvancedFilter を filter 層から必ずしも消さない。
        // そこで metadata.state の外部変化を検知し、UI を復元したうえで「現状態で
        // フィルタを再発火」して filter 層を同期する（空条件なら remove＝リセット成立）。
        const stateSig = this.computeStateSig(dv);
        if (!this.persistedSeen) {
            this.persistedSeen = true;
            this.lastStateSig = stateSig;
        } else if (stateSig === this.lastStateSig) {
            // 自分の persist が反映された（または変化なし）→ 待機解除
            this.awaitingPersist = false;
        } else if (this.awaitingPersist) {
            // 自分の persist がまだ metadata に届いていない stale 読み。
            // ここで「外部変化」と誤検知すると古い条件を復元してクリアが取り消される。
            // 反映を待つだけ（何もしない）。
        } else {
            // 自分の persist では無い真の外部変化（＝ブックマーク）
            this.lastStateSig = stateSig;
            this.restoreFromPersisted(dv);
            this.emitCurrent();   // filter 層を UI に合わせて同期（空なら除去）
            this.appliedConds = this.form.getConditions();
            this.clearPending();
            return;               // jsonFilters 復元は次サイクルに委ねる
        }

        // 外部 jsonFilters からの復元（スライサー同期）。
        // Data 更新時のみ未適用入力もリセット対象に（resize/style では typing を壊さない）。
        const isDataUpdate = ((options.type ?? VisualUpdateType.All) & VisualUpdateType.Data) !== 0;
        this.restoreFromJsonFilters(options.jsonFilters, cols, isDataUpdate);
    }

    /**
     * 条件フォームに出す列リストを構築。
     * categorical の categories（候補列）∪ values（列=メジャー）の source を
     * queryName で dedupe マージ。両方 categorical で確実に届く。
     * 「列」はメジャー集計されるが値は無視し metadata（queryName）のみ使用。
     * buildFilterTarget が集計ラッパー（Count/First/Sum）を剥がすのでフィルタ可。
     * 表示名は conditionForm.cleanLabel が「最初の 〜」を剥がして素の列名にする。
     */
    private resolveColumns(dv: DataView | null): powerbi.DataViewMetadataColumn[] {
        const out: powerbi.DataViewMetadataColumn[] = [];
        const seen = new Set<string>();
        const add = (c?: powerbi.DataViewMetadataColumn): void => {
            if (!c) return;
            const key = c.queryName ?? c.displayName ?? "";
            if (key) {
                if (seen.has(key)) return;
                seen.add(key);
            }
            out.push(c);
        };
        // 「列」(values) を先、「候補列」(categories) を後に並べる
        // （候補列がドロップダウン先頭に来てしまうのを防ぐ）
        for (const val of dv?.categorical?.values ?? []) add(val?.source);
        for (const cat of dv?.categorical?.categories ?? []) add(cat?.source);
        return out;
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    /** ビジュアル破棄時にタイマーを残さない */
    public destroy(): void {
        if (this.applyFallbackTimer !== null) {
            clearTimeout(this.applyFallbackTimer);
            this.applyFallbackTimer = null;
        }
    }

    // ==========================================================

    private onFormChange(): void {
        const cols = (this.lastColsRef as powerbi.DataViewMetadataColumn[] | null) ?? [];
        if (cols.length === 0) return;

        const conds = this.form.getConditions();
        const columnLogic = this.form.getColumnLogic();

        const result = emitAdvancedFilter(
            this.host, cols, conds, columnLogic, this.lastFilterSig,
            this.candidateColIdx(cols), this.selfFilterEnabled());
        // remove を発火した場合、消した sig を記憶して stale echo を無視できるようにする
        if (result.emitted && result.sig === "" && this.lastFilterSig !== "") {
            this.removedSig = this.lastFilterSig;
        }
        if (result.emitted || result.sig !== this.lastFilterSig) {
            this.lastFilterSig = result.sig;
        }
        // filter target を作れなかった列（メジャー等）を行警告で可視化
        this.form.setUnfilterable(new Set(result.dropped));
        this.appliedConds = conds;
        this.persist(conds, columnLogic);
        // 自分の persist は外部変化として誤検知しないよう sig を更新＋反映待ちフラグ。
        // （persist は非同期で、直後の update には古い metadata が届きうるため）
        this.lastStateSig = this.makeStateSig(conds, columnLogic);
        this.awaitingPersist = true;

        // 発火したら jsonFilters エコー受信までスピナー継続。発火しなければ即解除。
        if (result.emitted) {
            this.pendingApply = true;
            if (this.applyFallbackTimer !== null) clearTimeout(this.applyFallbackTimer);
            // 念のためのフォールバック（エコーが来ないケースで固まらないよう）
            this.applyFallbackTimer = window.setTimeout(() => this.clearPending(), 10000);
        } else {
            this.clearPending();
        }
    }

    /**
     * 適用せずとも UI 状態（条件構造）を persist。フィルタは発火しない。
     * これで「未適用で行追加しただけ」でも metadata が UI を反映し、
     * ブックマークリセット時に変化が検知されて UI がリセットされる。
     */
    private persistUiState(): void {
        const conds = this.form.getConditions();
        const columnLogic = this.form.getColumnLogic();
        this.persist(conds, columnLogic);
        this.lastStateSig = this.makeStateSig(conds, columnLogic);
        this.awaitingPersist = true;
    }

    /** スピナー解除＋フォールバックタイマー停止 */
    private clearPending(): void {
        this.pendingApply = false;
        if (this.applyFallbackTimer !== null) {
            clearTimeout(this.applyFallbackTimer);
            this.applyFallbackTimer = null;
        }
        this.form.setApplying(false);
    }

    /** 現在のフォーム状態で filter を再発火（ブックマーク復元時の同期用、スピナー無し） */
    private emitCurrent(): void {
        const cols = (this.lastColsRef as powerbi.DataViewMetadataColumn[] | null) ?? [];
        const conds = this.form.getConditions();
        const columnLogic = this.form.getColumnLogic();
        const result = emitAdvancedFilter(
            this.host, cols, conds, columnLogic, this.lastFilterSig,
            this.candidateColIdx(cols), this.selfFilterEnabled());
        // ブックマーク経由の remove も stale echo で取り消されないよう sig を記憶
        if (result.emitted && result.sig === "" && this.lastFilterSig !== "") {
            this.removedSig = this.lastFilterSig;
        }
        if (result.emitted || result.sig !== this.lastFilterSig) {
            this.lastFilterSig = result.sig;
        }
    }

    /** selfFilter（候補のサーバー連動）トグル。既定 OFF（安全側） */
    private selfFilterEnabled(): boolean {
        return this.formattingSettings?.behaviorCard?.selfFilterEnabled?.value === true;
    }

    /** 列リストの内容が前回と同一なら前回の配列参照を返す（DOM 再構築抑止） */
    private stableCols(cols: powerbi.DataViewMetadataColumn[]): powerbi.DataViewMetadataColumn[] {
        const prev = this.lastColsRef as powerbi.DataViewMetadataColumn[] | null;
        if (!prev || prev.length !== cols.length) return cols;
        for (let i = 0; i < cols.length; i++) {
            const a = prev[i], b = cols[i];
            if ((a?.queryName ?? "") !== (b?.queryName ?? "")
                || (a?.displayName ?? "") !== (b?.displayName ?? "")) {
                return cols;
            }
        }
        return prev;
    }

    /** 候補列（categorical.categories 由来）に対応する col index の集合 */
    private candidateColIdx(cols: powerbi.DataViewMetadataColumn[]): Set<number> {
        const catQn = new Set<string>();
        for (const cat of this.lastDataView?.categorical?.categories ?? []) {
            const qn = cat?.source?.queryName;
            if (qn) catQn.add(qn);
        }
        const s = new Set<number>();
        cols.forEach((c, i) => {
            if (catQn.has(c?.queryName ?? "")) s.add(i);
        });
        return s;
    }

    /** metadata.objects.state の状態シグネチャ（外部=ブックマーク書き換え検知用） */
    private computeStateSig(dv: DataView | null): string {
        const s = dv?.metadata?.objects?.["state"];
        if (!s) return "";
        const c = String(s["conditionsJson"] ?? "");
        const l = String(s["columnLogicJson"] ?? s["logic"] ?? "");
        return `${c}\0${l}`;
    }

    /** 自分が書き込む state のシグネチャ（computeStateSig と一致する形式） */
    private makeStateSig(conds: FilterCondition[], columnLogic: ColumnLogic): string {
        return `${JSON.stringify(conds)}\0${JSON.stringify(columnLogic)}`;
    }

    /** metadata.objects.state から UI を復元（空ならデフォルト1行へ）。ブックマーク復元用 */
    private restoreFromPersisted(dv: DataView | null): void {
        const s = dv?.metadata?.objects?.["state"];
        const json = s ? String(s["conditionsJson"] ?? "") : "";
        const colLogicJson = s ? String(s["columnLogicJson"] ?? "") : "";

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
                        if (op !== "contains" && op !== "notContains") continue;
                        conds.push({ columnIndex: ci, operator: op, value: val });
                    }
                }
            } catch { /* ignore */ }
        }

        const columnLogic: ColumnLogic = {};
        if (colLogicJson) {
            try {
                const parsed = JSON.parse(colLogicJson) as unknown;
                if (parsed && typeof parsed === "object" && !Array.isArray(parsed)) {
                    for (const [k, v] of Object.entries(parsed as Record<string, unknown>)) {
                        if (v === "OR" || v === "AND") columnLogic[k] = v;
                    }
                }
            } catch { /* ignore */ }
        }

        if (conds.length === 0) {
            this.form.resetToDefault();
        } else {
            this.form.setState(conds, columnLogic);
        }
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

    // ==========================================================
    // jsonFilters 受信
    // ==========================================================

    private restoreFromJsonFilters(
        jsonFilters: powerbi.IFilter[] | undefined,
        cols: powerbi.DataViewMetadataColumn[],
        isDataUpdate: boolean,
    ): void {
        const restored = restoreFromAdvancedFilters(jsonFilters, cols);

        // ブックマーク / 外部スライサーで自分の filter が解除された場合は UI をリセット。
        // 未適用の入力中テキストも Data 更新（ブックマーク含む）時のみ掃除する。
        // resize / style の update では typing 中のテキストを壊さない。
        if (!restored) {
            // remove が反映された空 echo → stale ガード解除
            this.removedSig = "";
            // 自分の filter 解除が反映された（remove のエコー）→ スピナー解除
            if (this.pendingApply && isDataUpdate) this.clearPending();
            // 適用済み filter が消えた場合のみ UI リセット。
            // 未適用入力は metadata 経路（typed 値も persist 済み）でブックマーク検知する。
            // ここで「入力あり=リセット」にすると条件追加等の自前更新で入力が消える。
            if (this.lastFilterSig !== "") {
                this.lastFilterSig = "";
                this.appliedConds = [];
                this.form.resetToDefault();
                // 外部クリアで main が消えても selfFilter は残留しうるので揃えて除去
                syncSelfFilter(this.host, cols, [], {}, undefined, this.selfFilterEnabled());
            }
            return;
        }

        // 自己発火エコー = 自分の filter がモデルに反映された確証 → スピナー解除
        if (restored.sig === this.lastFilterSig) {
            if (this.pendingApply) this.clearPending();
            return;
        }

        // remove 直後の stale echo（消したはずの旧フィルタが in-flight update で戻ってきた）
        // → 復元するとクリアが取り消されるので無視。空 echo か別フィルタ受信で解除される
        if (this.removedSig !== "" && restored.sig === this.removedSig) {
            return;
        }
        this.removedSig = "";

        // 有効な active 条件が入ってきたら UI を上書き（ブックマーク含む）
        this.form.setState(restored.conditions, restored.columnLogic);
        this.lastFilterSig = restored.sig;
        this.appliedConds = restored.conditions;
        // レポート再オープン / ページ再訪では selfFilter は誰も発火していないため、
        // サーバー側候補カスケードを復元状態に同期する（main は既に filter 層にあり再発火しない）。
        // トグル OFF（既定）では remove が送られ、過去バージョンの残留 selfFilter を掃除する
        syncSelfFilter(this.host, cols, restored.conditions, restored.columnLogic,
            this.candidateColIdx(cols), this.selfFilterEnabled());
    }

    // ==========================================================

    /**
     * categorical.categories（=候補列）から distinct を抽出。カスケード絞り込み付き。
     * 適用中の active 条件（候補列を target にするもの）でタプルを絞り、各候補列は
     * 「自分以外の条件を満たすタプル」の distinct を返す。
     * → 1 行目で組織名を絞ると 2 行目以降の候補がそれに連動して減る。
     * Power BI は適用元ビジュアル自身を再クエリしないので、ここでクライアント側に実施。
     * TEST/ダミー除外、DISPLAY_CAP まで。列(metadata のみ)は候補なし。
     */
    private extractUniques(
        dv: DataView | null,
        cols: powerbi.DataViewMetadataColumn[],
        conds: FilterCondition[],
    ): string[][] {
        const categories = dv?.categorical?.categories ?? [];
        const condsSig = JSON.stringify(
            conds.filter(isConditionActive).map(c => [c.columnIndex, c.operator, c.value])
        );

        if (this.uniquesCache
            && this.uniquesCache.catRef === categories
            && this.uniquesCache.condsSig === condsSig) {
            return this.uniquesCache.result;
        }

        if (categories.length === 0) {
            let empty = cols.map(() => [] as string[]);
            if (this.uniquesCache
                && JSON.stringify(this.uniquesCache.result) === JSON.stringify(empty)) {
                empty = this.uniquesCache.result;
            }
            this.uniquesCache = { catRef: categories, condsSig, result: empty };
            return empty;
        }

        // queryName → category(候補列) index
        const catIdxByQn = new Map<string, number>();
        categories.forEach((cat, i) => {
            const qn = cat?.source?.queryName;
            if (qn) catIdxByQn.set(qn, i);
        });

        // 候補列を target にする active 条件だけをタプル制約にする
        const constraints: { catIdx: number; op: FilterOp; value: string }[] = [];
        for (const c of conds) {
            if (!isConditionActive(c)) continue;
            const qn = cols[c.columnIndex]?.queryName;
            if (!qn) continue;
            const catIdx = catIdxByQn.get(qn);
            if (catIdx === undefined) continue; // 「列」側 or 未マッチ → 制約にしない
            constraints.push({ catIdx, op: c.operator, value: c.value });
        }

        const DISPLAY_CAP = 200;
        const tupleCount = categories[0]?.values?.length ?? 0;
        const byQueryName = new Map<string, string[]>();
        for (let ci = 0; ci < categories.length; ci++) {
            const cat = categories[ci];
            const src = cat?.source;
            if (!src) continue;
            const vals = cat.values ?? [];
            const set = new Set<string>();
            for (let t = 0; t < tupleCount; t++) {
                // 自分以外の制約を満たすタプルだけ（カスケード）
                let ok = true;
                for (const con of constraints) {
                    if (con.catIdx === ci) continue;
                    if (!this.evalCond(categories[con.catIdx].values[t], con.op, con.value)) {
                        ok = false;
                        break;
                    }
                }
                if (!ok) continue;
                const v = vals[t];
                if (v == null) continue;
                const s = String(v);
                if (s === "") continue;
                if (s.toUpperCase().includes("TEST")) continue;
                if (s.includes("ダミー")) continue;
                set.add(s);
                if (set.size >= DISPLAY_CAP) break;
            }
            byQueryName.set(src.queryName ?? "", Array.from(set).sort((a, b) => a.localeCompare(b)));
        }

        let result = cols.map(c => byQueryName.get(c?.queryName ?? "") ?? []);
        // 内容が前回と同一なら前回の配列参照を返す（setColumns の全再描画を抑止）
        if (this.uniquesCache
            && JSON.stringify(this.uniquesCache.result) === JSON.stringify(result)) {
            result = this.uniquesCache.result;
        }
        this.uniquesCache = { catRef: categories, condsSig, result };
        return result;
    }

    /** タプル値が条件 (op, value) を満たすか（カスケード用、emit と同じ意味論） */
    private evalCond(tv: powerbi.PrimitiveValue, op: FilterOp, value: string): boolean {
        if (tv == null) return false;
        // トリム＋内側空白の全角/半角ゆらぎ畳み込み（emit のバリアント展開と同じ意味論）
        const fold = (s: string): string => s.replace(/[\s　]+/g, " ");
        const hit = fold(String(tv)).toLowerCase().includes(fold(value.trim()).toLowerCase());
        return op === "contains" ? hit : !hit;
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
