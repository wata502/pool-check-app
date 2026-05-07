/**
 * weather_judge.js  v7 — ハイブリッド取得・2要素配列版
 *
 * 設計方針:
 *  - 広島県福山市アメダス観測所（67401）のデータのみを使用する。
 *  - map API（1件）と point API（1ファイル）のハイブリッドで取得した
 *    要素数2の配列を受け取り、厳格なルールで天気を判定する。
 *  - モジュール内部にステート変数を持たない（呼び出しごとに完全決定）。
 *  - 外部APIへの通信処理は含まない純粋判定ロジック。
 *
 * judge() の引数:
 *   amedasHistory : Array  要素数2の配列
 *     [0] : 最新10分マップデータ { prec10: number|null, temp: number|null, sun10m: number|null }
 *     [1] : 地点専用データ      { weather: number|null }  ← point APIから抽出した天気コード
 *
 * 天気コードマッピング（気象庁 point API）:
 *   0 → 晴れ  ※ 0 は「データなし」ではなく「晴れ」の公式コード。絶対に無効値として弾かないこと
 *   1 → 曇り
 *   7 → 雨
 *   9 → 雪
 *
 * judge() の戻り値:
 *   { weather, notice, is_raining, judged_at }
 *     weather    : "雨" | "雪" | "曇り" | "晴れ" | "データなし"
 *     notice     : string|null
 *     is_raining : boolean
 *     judged_at  : string (ISO8601)
 */

const WeatherJudge = (() => {

  // ══════════════════════════════════════════════════
  // 天気コード → ラベル変換（気象庁 point API）
  //   0 → 晴れ  ※ 0 は「データなし」ではなく「晴れ」の公式コード
  //   1 → 曇り
  //   7 → 雨
  //   9 → 雪
  // ══════════════════════════════════════════════════
  const CODE_MAP = { 0: '晴れ', 1: '曇り', 7: '雨', 9: '雪' };

  /**
   * 天気コードをラベル文字列に変換する。
   * 有効なコード（0,1,7,9）以外は null を返す。
   * ※ code=0 は「晴れ」の公式コードであり null 扱いしてはならない。
   * @param {*} code
   * @returns {string|null}
   */
  function _codeToLabel(code) {
    if (code == null) return null;
    return Object.prototype.hasOwnProperty.call(CODE_MAP, code)
      ? CODE_MAP[code]
      : null;
  }

  /**
   * 戻り値オブジェクトを生成する。
   * @param {string} weather
   * @param {string} judged_at
   * @returns {{ weather: string, notice: string|null, is_raining: boolean, judged_at: string }}
   */
  function _makeResult(weather, judged_at) {
    return {
      weather,
      notice:     weather === 'データなし'
                    ? '※現在、気象庁の自動観測データが欠測しています'
                    : null,
      is_raining: (weather === '雨' || weather === '雪'),
      judged_at
    };
  }

  // ══════════════════════════════════════════════════
  // 公開 API
  // ══════════════════════════════════════════════════
  return {

    /**
     * 天気判定（同期関数）
     *
     * @param {Array} amedasHistory 要素数2の配列
     *   [0] 最新10分マップデータ { prec10, temp, sun10m }
     *   [1] 地点専用データ      { weather: 天気コード }
     * @returns {{ weather: string, notice: string|null, is_raining: boolean, judged_at: string }}
     */
    judge(amedasHistory) {
      const judged_at = new Date().toISOString();

      // ─────────────────────────────────────────────────
      // 入力チェック
      // ─────────────────────────────────────────────────
      if (!Array.isArray(amedasHistory) || amedasHistory.length === 0) {
        console.debug('[WJ] judge() → データなし（配列が空または不正）');
        return _makeResult('データなし', judged_at);
      }

      // ─────────────────────────────────────────────────
      // Step 0: BaseWeather の特定
      //   amedasHistory[1]（または配列全体）から有効な天気コード（0,1,7,9）を探し、
      //   初期状態（BaseWeather）とする。見つからなければ「データなし」。
      //
      //   【重要】code=0 は「晴れ」の公式コードであり「データなし」ではない。
      //   有効コード（0,1,7,9）以外が来た場合のみ無効とみなす。
      // ─────────────────────────────────────────────────
      let baseWeather = 'データなし';

      for (let i = 1; i < amedasHistory.length; i++) {
        const entry = amedasHistory[i];
        if (!entry || entry.weather == null) continue;
        const label = _codeToLabel(entry.weather);
        if (label !== null) {
          baseWeather = label;
          console.debug(
            `[WJ] Step0 → BaseWeather="${baseWeather}"（index=${i}, code=${entry.weather}）`
          );
          break;
        }
      }

      if (baseWeather === 'データなし') {
        console.debug('[WJ] Step0 → 有効コードなし → データなし');
        return _makeResult('データなし', judged_at);
      }

      // ─────────────────────────────────────────────────
      // Step 1: 絶対条件による「確定」と「状態更新」（最優先）
      //   最新の10分データ（amedasHistory[0]）を参照し、
      //   以下に該当すれば BaseWeather を更新して即座に結論を返す。
      // ─────────────────────────────────────────────────
      const latest = amedasHistory[0] || {};
      const prec10 = latest.prec10 != null ? latest.prec10 : null;
      const temp   = latest.temp   != null ? latest.temp   : null;
      const sun10m = latest.sun10m != null ? latest.sun10m : null;

      console.debug(
        `[WJ] Step1 参照値 → prec10=${prec10} temp=${temp} sun10m=${sun10m}`
      );

      // 【雪・確定】10分降水量 >= 0.5mm AND 気温 <= 0℃ AND 10分日照 == 0分
      if (prec10 !== null && prec10 >= 0.5 &&
          temp   !== null && temp   <= 0   &&
          sun10m !== null && sun10m === 0) {
        console.debug(
          `[WJ] Step1 → 雪・確定（prec10=${prec10} temp=${temp} sun10m=${sun10m}）`
        );
        return _makeResult('雪', judged_at);
      }

      // 【雨・確定】10分降水量 >= 0.5mm AND 気温 > 0℃ AND 10分日照 == 0分
      if (prec10 !== null && prec10 >= 0.5 &&
          temp   !== null && temp   >  0   &&
          sun10m !== null && sun10m === 0) {
        console.debug(
          `[WJ] Step1 → 雨・確定（prec10=${prec10} temp=${temp} sun10m=${sun10m}）`
        );
        return _makeResult('雨', judged_at);
      }

      // ─────────────────────────────────────────────────
      // Step 2: 10分データによる「状態遷移」
      //   Step 1 で結論が出なかった場合、Step 0 の BaseWeather を
      //   最新の10分値で更新して最終結果を返す。
      // ─────────────────────────────────────────────────

      if (baseWeather === '雨' || baseWeather === '雪') {
        // Case A: BaseWeather が「雨」または「雪」の場合
        if (prec10 !== null && prec10 === 0.0) {
          if (sun10m !== null && sun10m >= 5) {
            // 条件1: 10分降水量 == 0.0mm AND 10分日照 >= 5分 → 晴れ
            console.debug(
              `[WJ] Step2 CaseA → 晴れ（prec10=0 sun10m=${sun10m}）`
            );
            baseWeather = '晴れ';
          } else if (sun10m !== null && sun10m >= 1) {
            // 条件2: 10分降水量 == 0.0mm AND 10分日照 >= 1分 → 曇り
            console.debug(
              `[WJ] Step2 CaseA → 曇り（prec10=0 sun10m=${sun10m}）`
            );
            baseWeather = '曇り';
          } else {
            // 上記以外: 状態維持
            console.debug(
              `[WJ] Step2 CaseA → 状態維持（${baseWeather} prec10=0 sun10m=${sun10m}）`
            );
          }
        } else {
          // 降水継続（または降水データなし） → 状態維持
          console.debug(
            `[WJ] Step2 CaseA → 状態維持（${baseWeather} prec10=${prec10}）`
          );
        }

      } else if (baseWeather === '曇り') {
        // Case B: BaseWeather が「曇り」の場合
        if (sun10m !== null && sun10m >= 5) {
          // 条件1: 10分日照 >= 5分 → 晴れ
          console.debug(
            `[WJ] Step2 CaseB → 晴れ（sun10m=${sun10m}）`
          );
          baseWeather = '晴れ';
        } else {
          // 上記以外: 状態維持
          console.debug(
            `[WJ] Step2 CaseB → 状態維持（曇り sun10m=${sun10m}）`
          );
        }

      } else if (baseWeather === '晴れ') {
        // Case C: BaseWeather が「晴れ」の場合
        if (sun10m !== null && sun10m === 0) {
          // 条件1: 10分日照 == 0分 → 曇り
          console.debug('[WJ] Step2 CaseC → 曇り（sun10m=0）');
          baseWeather = '曇り';
        } else {
          // 上記以外: 状態維持
          console.debug(
            `[WJ] Step2 CaseC → 状態維持（晴れ sun10m=${sun10m}）`
          );
        }
      }

      return _makeResult(baseWeather, judged_at);
    }

  };

})();

// CommonJS / ES Module 両対応
if (typeof module !== 'undefined' && module.exports) {
  module.exports = WeatherJudge;
}
