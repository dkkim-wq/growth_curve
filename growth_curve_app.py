import streamlit as st
import openpyxl
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import os
import json

# ══════════════════════════════════════════════════════════════════
# 성장곡선 검증 Streamlit 웹 앱
# ══════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="성장곡선 검증 프로그램",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── 스타일 ────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        font-size: 2rem;
        font-weight: bold;
        color: #1A3A5C;
        padding: 0.5rem 0;
    }
    .sub-header {
        font-size: 1rem;
        color: #5D6D7E;
        margin-bottom: 1rem;
    }
    .metric-card {
        background: #EBF5FB;
        border-radius: 8px;
        padding: 1rem;
        border-left: 4px solid #2471A3;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
</style>
""", unsafe_allow_html=True)

# ── 파일 경로 ─────────────────────────────────────────────────────
FILE_PATH = os.path.join(os.path.dirname(__file__),
                         "260618 경쟁사유무에 따른 성장곡선(공유)V3.xlsx")
ADDED_STORES_PATH = os.path.join(os.path.dirname(__file__),
                                  "added_stores.json")


# ── 추가 매장 데이터 저장/로드 ────────────────────────────────────
def load_added_stores():
    if os.path.exists(ADDED_STORES_PATH):
        with open(ADDED_STORES_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []


def save_added_stores(stores_list):
    with open(ADDED_STORES_PATH, 'w', encoding='utf-8') as f:
        json.dump(stores_list, f, ensure_ascii=False, indent=2)


# ── 엑셀 데이터 로드 (캐시) ───────────────────────────────────────
@st.cache_data
def load_excel_data():
    """V3 엑셀에서 곡선지수와 매장 매출 데이터를 로드"""
    wb = openpyxl.load_workbook(FILE_PATH, data_only=True, read_only=True)

    # 1. 가중평균 곡선지수 (6열: 가중평균(0~2))
    ws_curve = wb['요약_성장곡선']
    curve_index = {}
    for row in ws_curve.iter_rows(min_row=3, max_row=150, max_col=10,
                                   values_only=True):
        m = row[0]   # 오픈후개월(m)
        w_avg = row[5]  # 가중평균(0~2)
        if m is not None and w_avg is not None and w_avg != 0:
            curve_index[int(m)] = w_avg

    # 2. RAW_매출A: 세로 형태 → 매장별 월매출 리스트로 변환
    ws_raw = wb['RAW_매출A']
    from collections import defaultdict
    store_data = defaultdict(lambda: {'open_date': None, 'monthly': {}})

    for row in ws_raw.iter_rows(min_row=2, max_col=10, values_only=True):
        date_val = row[0]      # 해당월
        store_name = row[1]    # 지점키
        open_date = row[3]     # 개점일
        sales = row[6]         # 세탁건조매출

        if store_name is None or date_val is None:
            continue

        if store_data[store_name]['open_date'] is None and open_date:
            store_data[store_name]['open_date'] = open_date

        if sales is not None and sales > 0:
            store_data[store_name]['monthly'][date_val] = sales

    # 매장별로 개점일 기준 월차(m0, m1, ...) 순서로 매출 정렬
    stores = []
    for name, data in store_data.items():
        if not data['monthly'] or not data['open_date']:
            continue

        # 날짜순 정렬
        sorted_months = sorted(data['monthly'].keys())
        sales_list = [data['monthly'][m] for m in sorted_months]

        if len(sales_list) >= 4:  # 최소 데이터 필요
            stores.append({'name': name, 'sales': sales_list})

    wb.close()
    return curve_index, stores


# ── 검증 함수 ─────────────────────────────────────────────────────
def get_base_revenue(store, curve_index, method):
    """단일 방식(A/B/C)의 기준매출 계산. 실패 시 None 반환."""
    sales = store['sales']
    if len(sales) < 10:
        return None

    start_m = {'A': 1, 'B': 2, 'C': 3}[method]
    end_m = 3

    base_estimates = []
    for m in range(start_m, end_m + 1):
        if m >= len(sales) or sales[m] is None or sales[m] <= 0:
            continue
        if m not in curve_index:
            continue
        base_est = sales[m] / (curve_index[m] / 100)
        base_estimates.append(base_est)

    if not base_estimates:
        return None
    return np.mean(base_estimates)


# 전체 방식 목록
ALL_METHODS = ['A', 'B', 'C', 'AB', 'BC', 'ABC']
METHOD_LABELS = {
    'A': 'm1부터', 'B': 'm2부터', 'C': 'm3부터',
    'AB': 'A+B 평균', 'BC': 'B+C 평균', 'ABC': 'A+B+C 평균'
}


def validate_store(store, curve_index, method='A'):
    """
    method: 'A','B','C' = 단일, 'AB','BC','ABC' = 복합(기준매출 평균)
    """
    sales = store['sales']
    if len(sales) < 10:
        return None

    actual_m4_m9 = sales[4:10]
    if any(v is None or v <= 0 for v in actual_m4_m9):
        return None

    # 기준매출 계산
    if method in ('A', 'B', 'C'):
        base_revenue = get_base_revenue(store, curve_index, method)
    elif method == 'AB':
        a = get_base_revenue(store, curve_index, 'A')
        b = get_base_revenue(store, curve_index, 'B')
        if a is None or b is None:
            return None
        base_revenue = (a + b) / 2
    elif method == 'BC':
        b = get_base_revenue(store, curve_index, 'B')
        c = get_base_revenue(store, curve_index, 'C')
        if b is None or c is None:
            return None
        base_revenue = (b + c) / 2
    elif method == 'ABC':
        a = get_base_revenue(store, curve_index, 'A')
        b = get_base_revenue(store, curve_index, 'B')
        c = get_base_revenue(store, curve_index, 'C')
        if a is None or b is None or c is None:
            return None
        base_revenue = (a + b + c) / 3
    else:
        return None

    if base_revenue is None:
        return None

    # m4~m9 예측
    predicted = []
    for m in range(4, 10):
        if m in curve_index:
            pred = base_revenue * (curve_index[m] / 100)
            predicted.append(pred)
        else:
            return None

    # 월별 오차율 (부호 포함)
    errors = []
    for pred, actual in zip(predicted, actual_m4_m9):
        error_pct = (pred - actual) / actual * 100
        errors.append(error_pct)

    # 오차율: 기준매출 vs 실제평균(m4~9)
    actual_avg = np.mean(actual_m4_m9)
    avg_error = (base_revenue - actual_avg) / actual_avg * 100

    return {
        'base_revenue': base_revenue,
        'predicted': predicted,
        'actual': actual_m4_m9,
        'errors': errors,
        'avg_error': avg_error
    }


# ── 전체 성장곡선 예측 (임의 개월까지) ────────────────────────────
def predict_growth_curve(base_revenue, curve_index, months=48):
    """기준매출과 곡선지수로 월별 예측매출 생성"""
    predictions = {}
    for m in range(1, months + 1):
        if m in curve_index:
            predictions[m] = base_revenue * (curve_index[m] / 100)
    return predictions


# ══════════════════════════════════════════════════════════════════
# 메인 앱
# ══════════════════════════════════════════════════════════════════
def main():
    st.markdown('<div class="main-header">📈 성장곡선 검증 프로그램</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="sub-header">'
                'm1~m3 매출로 역산한 기준매출 → m4~m9 예측 vs 실제 오차율 비교 '
                '| 곡선지수: 그룹0·1·2 매장수 가중평균</div>',
                unsafe_allow_html=True)

    # 데이터 로드
    if not os.path.exists(FILE_PATH):
        st.error(f"엑셀 파일을 찾을 수 없습니다: {FILE_PATH}")
        st.info("파일 경로를 확인하거나 엑셀 파일을 업로드해주세요.")
        return

    curve_index, stores = load_excel_data()

    # 추가 매장 데이터 로드
    added_stores = load_added_stores()
    all_stores = stores + added_stores

    # 사이드바
    with st.sidebar:
        st.header("⚙️ 설정")
        st.markdown("---")
        st.subheader("📊 데이터 현황")
        st.metric("엑셀 매장 수", len(stores))
        st.metric("추가 매장 수", len(added_stores))
        st.metric("곡선지수 월차 수", len(curve_index))
        st.markdown("---")
        st.subheader("🔧 검증 방식")
        selected_method = st.radio(
            "역산 시작월 선택",
            ALL_METHODS,
            format_func=lambda x: f"{x}: {METHOD_LABELS[x]}",
            index=0
        )

    # 탭 구성
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 전체 검증 결과", "🔍 개별 매장 조회",
        "📈 성장곡선 차트", "➕ 매장 데이터 추가"
    ])

    # ═══════════════════════════════════════════════════════════════
    # 탭 1: 전체 검증 결과
    # ═══════════════════════════════════════════════════════════════
    with tab1:
        st.subheader("🏪 전체 매장 검증 결과")

        # 6가지 방식 비교
        results_all = {}
        for method in ALL_METHODS:
            results_all[method] = []
            for store in all_stores:
                result = validate_store(store, curve_index, method)
                if result is not None:
                    results_all[method].append({
                        'name': store['name'], **result
                    })

        # 신뢰도 있는 평균 (트리밍 평균: 상하 10% 제거)
        def trimmed_mean(values, trim_pct=0.1):
            """상하 trim_pct% 제거 후 평균 (이상치 영향 최소화)"""
            if not values:
                return 0
            sorted_vals = sorted(values)
            n = len(sorted_vals)
            trim_n = int(n * trim_pct)
            if trim_n > 0:
                trimmed = sorted_vals[trim_n:-trim_n]
            else:
                trimmed = sorted_vals
            return np.mean(trimmed) if trimmed else np.mean(sorted_vals)

        # 요약 카드 (6개 방식, 2행 × 3열)
        st.markdown("##### 단일 방식")
        col1, col2, col3 = st.columns(3)
        for col, method in zip([col1, col2, col3], ['A', 'B', 'C']):
            errs = [abs(r['avg_error']) for r in results_all[method]]
            with col:
                if errs:
                    avg_abs_err = trimmed_mean(errs)
                    std_err = np.std(errs)
                    st.metric(
                        f"방식 {method} ({METHOD_LABELS[method]})",
                        f"{avg_abs_err:.2f}%",
                        f"±{std_err:.1f}% (σ) | {len(errs)}개 매장"
                    )
                else:
                    st.metric(f"방식 {method}", "데이터 없음", "0개 매장")

        st.markdown("##### 복합 방식")
        col4, col5, col6 = st.columns(3)
        for col, method in zip([col4, col5, col6], ['AB', 'BC', 'ABC']):
            errs = [abs(r['avg_error']) for r in results_all[method]]
            with col:
                if errs:
                    avg_abs_err = trimmed_mean(errs)
                    std_err = np.std(errs)
                    st.metric(
                        f"방식 {method} ({METHOD_LABELS[method]})",
                        f"{avg_abs_err:.2f}%",
                        f"±{std_err:.1f}% (σ) | {len(errs)}개 매장"
                    )
                else:
                    st.metric(f"방식 {method}", "데이터 없음", "0개 매장")

        # 최적 방식 표시
        valid_methods = [m for m in ALL_METHODS if results_all[m]]
        if valid_methods:
            best = min(valid_methods,
                       key=lambda m: trimmed_mean([abs(r['avg_error']) for r in results_all[m]]))
            best_val = trimmed_mean([abs(r['avg_error']) for r in results_all[best]])
            st.success(
                f"★ 최적 방식: **{best} ({METHOD_LABELS[best]})** — "
                f"트리밍 평균 오차율 {best_val:.2f}%"
            )

        # 매장별 최적 방식 카운트
        st.markdown("---")
        st.subheader("🏆 방식별 최적 매장 수 랭킹")
        best_count = {m: 0 for m in ALL_METHODS}
        store_best_method = {}  # 매장별 최적 방식 저장

        # 모든 매장에 대해 6가지 방식 중 절대 오차율이 가장 작은 것 선택
        for store in all_stores:
            min_err = None
            min_method = None
            for method in ALL_METHODS:
                result = validate_store(store, curve_index, method)
                if result is not None:
                    abs_err = abs(result['avg_error'])
                    if min_err is None or abs_err < min_err:
                        min_err = abs_err
                        min_method = method
            if min_method is not None:
                best_count[min_method] += 1
                store_best_method[store['name']] = min_method

        # 랭킹 표시 (많은 순)
        ranked = sorted(best_count.items(), key=lambda x: x[1], reverse=True)
        rank_df = pd.DataFrame([{
            '순위': i + 1,
            '방식': f"{m} ({METHOD_LABELS[m]})",
            '최적 매장 수': cnt,
            '비율': f"{cnt / sum(best_count.values()) * 100:.1f}%" if sum(best_count.values()) > 0 else "0%"
        } for i, (m, cnt) in enumerate(ranked)])
        st.dataframe(rank_df, use_container_width=True, hide_index=True)

        st.markdown("---")

        # 선택된 방식 상세 테이블
        st.subheader(f"📋 방식 {selected_method} 상세 결과")
        results = results_all[selected_method]

        if results:
            df = pd.DataFrame([{
                '매장명': r['name'],
                '기준매출': f"{r['base_revenue']:,.0f}",
                '실제평균(m4-9)': f"{np.mean(r['actual']):,.0f}",
                '오차(액수)': f"{r['base_revenue'] - np.mean(r['actual']):+,.0f}",
                '오차율': f"{r['avg_error']:+.2f}%",
                '최적방식': store_best_method.get(r['name'], '-')
            } for r in sorted(results, key=lambda x: abs(x['avg_error']))])

            st.dataframe(
                df,
                use_container_width=True,
                hide_index=True
            )

            # 분포 차트
            st.subheader("📊 오차율 분포")
            errors_list = [r['avg_error'] for r in results]
            fig_hist = go.Figure()
            fig_hist.add_trace(go.Histogram(
                x=errors_list, nbinsx=20,
                marker_color='#2471A3', opacity=0.8
            ))
            fig_hist.update_layout(
                xaxis_title="오차율 (%)",
                yaxis_title="매장 수",
                template="plotly_white",
                height=350
            )
            st.plotly_chart(fig_hist, use_container_width=True)
        else:
            st.warning("검증 가능한 매장이 없습니다.")

    # ═══════════════════════════════════════════════════════════════
    # 탭 2: 개별 매장 조회
    # ═══════════════════════════════════════════════════════════════
    with tab2:
        st.subheader("🔍 개별 매장 상세 조회")

        store_names = [s['name'] for s in all_stores]
        selected_store_name = st.selectbox(
            "매장 선택", store_names,
            index=0 if store_names else None
        )

        if selected_store_name:
            store = next(s for s in all_stores if s['name'] == selected_store_name)

            # 6가지 방식 결과
            st.markdown(f"### 📌 {selected_store_name}")

            st.markdown("**단일 방식**")
            cols = st.columns(3)
            for i, method in enumerate(['A', 'B', 'C']):
                result = validate_store(store, curve_index, method)
                with cols[i]:
                    st.markdown(f"**방식 {method} ({METHOD_LABELS[method]})**")
                    if result:
                        st.metric("기준매출", f"{result['base_revenue']:,.0f}원")
                        st.metric("실제평균(m4-9)", f"{np.mean(result['actual']):,.0f}원")
                        st.metric("오차율", f"{result['avg_error']:+.2f}%")
                    else:
                        st.info("데이터 부족")

            st.markdown("**복합 방식**")
            cols2 = st.columns(3)
            for i, method in enumerate(['AB', 'BC', 'ABC']):
                result = validate_store(store, curve_index, method)
                with cols2[i]:
                    st.markdown(f"**방식 {method} ({METHOD_LABELS[method]})**")
                    if result:
                        st.metric("기준매출", f"{result['base_revenue']:,.0f}원")
                        st.metric("실제평균(m4-9)", f"{np.mean(result['actual']):,.0f}원")
                        st.metric("오차율", f"{result['avg_error']:+.2f}%")
                    else:
                        st.info("데이터 부족")

            # 선택 방식 상세
            result = validate_store(store, curve_index, selected_method)
            if result:
                st.markdown("---")
                st.markdown(f"#### 방식 {selected_method} 상세")

                detail_df = pd.DataFrame({
                    '월차': [f'm{i}' for i in range(4, 10)],
                    '예측매출': [f"{p:,.0f}" for p in result['predicted']],
                    '실제매출': [f"{a:,.0f}" for a in result['actual']],
                    '오차율(%)': [f"{e:+.1f}%" for e in result['errors']]
                })
                st.dataframe(detail_df, use_container_width=True, hide_index=True)

                # 예측 vs 실제 비교 차트
                fig = go.Figure()
                months = [f'm{i}' for i in range(4, 10)]
                fig.add_trace(go.Bar(
                    x=months, y=result['actual'],
                    name='실제매출', marker_color='#2471A3'
                ))
                fig.add_trace(go.Bar(
                    x=months, y=result['predicted'],
                    name='예측매출', marker_color='#E74C3C', opacity=0.7
                ))
                fig.update_layout(
                    title=f"{selected_store_name} — m4~m9 예측 vs 실제",
                    xaxis_title="월차", yaxis_title="매출 (원)",
                    barmode='group', template="plotly_white",
                    height=400
                )
                st.plotly_chart(fig, use_container_width=True)

            # 전체 매출 추이
            st.markdown("---")
            st.markdown("#### 📈 전체 월별 매출 추이")
            sales = store['sales']
            valid_sales = [(i, s) for i, s in enumerate(sales)
                          if s is not None and s > 0]
            if valid_sales:
                fig2 = go.Figure()
                x_vals = [f'm{i}' for i, _ in valid_sales]
                y_vals = [s for _, s in valid_sales]
                fig2.add_trace(go.Scatter(
                    x=x_vals, y=y_vals,
                    mode='lines+markers',
                    name='실제매출',
                    line=dict(color='#1A3A5C', width=2),
                    marker=dict(size=8)
                ))

                # 성장곡선 예측 오버레이
                result_for_curve = validate_store(store, curve_index, selected_method)
                if result_for_curve:
                    preds = predict_growth_curve(
                        result_for_curve['base_revenue'], curve_index
                    )
                    pred_months = [f'm{m}' for m in sorted(preds.keys())]
                    pred_vals = [preds[m] for m in sorted(preds.keys())]
                    fig2.add_trace(go.Scatter(
                        x=pred_months, y=pred_vals,
                        mode='lines',
                        name='성장곡선 예측',
                        line=dict(color='#E74C3C', width=2, dash='dash')
                    ))

                fig2.update_layout(
                    title=f"{selected_store_name} — 월별 매출 추이",
                    xaxis_title="월차", yaxis_title="매출 (원)",
                    template="plotly_white", height=450
                )
                st.plotly_chart(fig2, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # 탭 3: 성장곡선 차트
    # ═══════════════════════════════════════════════════════════════
    with tab3:
        st.subheader("📈 성장곡선 지수 시각화")

        # 곡선지수 차트
        months_sorted = sorted(curve_index.keys())
        values_sorted = [curve_index[m] for m in months_sorted]

        fig_curve = go.Figure()
        fig_curve.add_trace(go.Scatter(
            x=[f'm{m}' for m in months_sorted],
            y=values_sorted,
            mode='lines+markers',
            name='가중평균 곡선지수',
            line=dict(color='#1A3A5C', width=3),
            marker=dict(size=6)
        ))
        fig_curve.add_hline(y=100, line_dash="dash",
                            line_color="red", opacity=0.5,
                            annotation_text="100% (기준)")
        fig_curve.update_layout(
            title="월차별 성장곡선 지수 (가중평균)",
            xaxis_title="월차",
            yaxis_title="곡선지수 (%)",
            template="plotly_white",
            height=450
        )
        st.plotly_chart(fig_curve, use_container_width=True)

        # 곡선지수 데이터 테이블
        with st.expander("📋 곡선지수 데이터 보기"):
            ci_df = pd.DataFrame({
                '월차': [f'm{m}' for m in months_sorted],
                '곡선지수(%)': [f"{v:.2f}" for v in values_sorted]
            })
            st.dataframe(ci_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("🏪 다중 매장 성장곡선 비교")

        # 매장 다중 선택
        compare_stores = st.multiselect(
            "비교할 매장 선택 (최대 5개)",
            [s['name'] for s in all_stores],
            max_selections=5
        )

        if compare_stores:
            fig_compare = go.Figure()
            colors = ['#1A3A5C', '#E74C3C', '#2ECC71', '#9B59B6', '#F39C12']

            for idx, store_name in enumerate(compare_stores):
                store = next(s for s in all_stores if s['name'] == store_name)
                sales = store['sales']
                valid = [(i, s) for i, s in enumerate(sales)
                         if s is not None and s > 0]
                if valid:
                    fig_compare.add_trace(go.Scatter(
                        x=[f'm{i}' for i, _ in valid],
                        y=[s for _, s in valid],
                        mode='lines+markers',
                        name=store_name,
                        line=dict(color=colors[idx % len(colors)], width=2),
                        marker=dict(size=6)
                    ))

            fig_compare.update_layout(
                title="매장별 매출 추이 비교",
                xaxis_title="월차", yaxis_title="매출 (원)",
                template="plotly_white", height=500
            )
            st.plotly_chart(fig_compare, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # 탭 4: 매장 데이터 추가
    # ═══════════════════════════════════════════════════════════════
    with tab4:
        st.subheader("➕ 신규 매장 데이터 추가")
        st.markdown("새 매장의 월별 매출 데이터를 입력하여 성장곡선 검증에 포함시킵니다.")

        with st.form("add_store_form"):
            store_name = st.text_input(
                "매장명", placeholder="예: (0099) 강남역점"
            )
            sales_input = st.text_area(
                "월별 매출 (m0부터, 쉼표로 구분)",
                placeholder="예: 500000, 1200000, 1500000, 2000000, ...",
                help="m0부터 순서대로 월매출을 입력하세요. 최소 10개월 필요."
            )
            submitted = st.form_submit_button("매장 추가", type="primary")

            if submitted:
                if not store_name:
                    st.error("매장명을 입력해주세요.")
                elif not sales_input:
                    st.error("매출 데이터를 입력해주세요.")
                else:
                    try:
                        sales_list = [
                            int(s.strip().replace(',', ''))
                            for s in sales_input.split(',')
                            if s.strip()
                        ]
                        if len(sales_list) < 4:
                            st.error("최소 4개월 이상의 매출 데이터가 필요합니다.")
                        else:
                            new_store = {
                                'name': store_name,
                                'sales': sales_list
                            }
                            added = load_added_stores()
                            added.append(new_store)
                            save_added_stores(added)
                            st.success(
                                f"✅ '{store_name}' 매장이 추가되었습니다! "
                                f"({len(sales_list)}개월 데이터)"
                            )
                            st.cache_data.clear()
                            st.rerun()
                    except ValueError:
                        st.error("매출 데이터 형식이 올바르지 않습니다. "
                                 "숫자와 쉼표만 사용해주세요.")

        # CSV 업로드 방식
        st.markdown("---")
        st.subheader("📁 CSV/엑셀 파일로 일괄 추가")
        uploaded_file = st.file_uploader(
            "파일 업로드 (CSV 또는 Excel)",
            type=['csv', 'xlsx'],
            help="컬럼: 매장명, m0, m1, m2, ... (월별 매출)"
        )

        if uploaded_file:
            try:
                if uploaded_file.name.endswith('.csv'):
                    upload_df = pd.read_csv(uploaded_file)
                else:
                    upload_df = pd.read_excel(uploaded_file)

                st.dataframe(upload_df.head(), use_container_width=True)

                if st.button("업로드 데이터 추가", type="primary"):
                    added = load_added_stores()
                    count = 0
                    for _, row in upload_df.iterrows():
                        name = str(row.iloc[0])
                        sales = [
                            int(v) for v in row.iloc[1:].dropna().values
                            if v and float(v) > 0
                        ]
                        if sales:
                            added.append({'name': name, 'sales': sales})
                            count += 1
                    save_added_stores(added)
                    st.success(f"✅ {count}개 매장이 추가되었습니다!")
                    st.cache_data.clear()
                    st.rerun()
            except Exception as e:
                st.error(f"파일 처리 오류: {e}")

        # 추가된 매장 관리
        st.markdown("---")
        st.subheader("📝 추가된 매장 관리")
        added = load_added_stores()
        if added:
            for i, store in enumerate(added):
                col_name, col_data, col_del = st.columns([3, 5, 1])
                with col_name:
                    st.markdown(f"**{store['name']}**")
                with col_data:
                    st.caption(
                        f"{len(store['sales'])}개월 | "
                        f"평균: {np.mean(store['sales']):,.0f}원"
                    )
                with col_del:
                    if st.button("🗑️", key=f"del_{i}"):
                        added.pop(i)
                        save_added_stores(added)
                        st.cache_data.clear()
                        st.rerun()
        else:
            st.info("추가된 매장이 없습니다.")


if __name__ == "__main__":
    main()
