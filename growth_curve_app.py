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
def validate_store(store, curve_index, method='A'):
    """
    method: 'A' = m1부터, 'B' = m2부터, 'C' = m3부터
    """
    sales = store['sales']
    if len(sales) < 10:
        return None

    actual_m4_m9 = sales[4:10]
    if any(v is None or v <= 0 for v in actual_m4_m9):
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

    base_revenue = np.mean(base_estimates)

    # m4~m9 예측
    predicted = []
    for m in range(4, 10):
        if m in curve_index:
            pred = base_revenue * (curve_index[m] / 100)
            predicted.append(pred)
        else:
            return None

    # 월별 오차율 (부호 포함: 예측 > 실제면 +, 예측 < 실제면 -)
    errors = []
    for pred, actual in zip(predicted, actual_m4_m9):
        error_pct = (pred - actual) / actual * 100
        errors.append(error_pct)

    # 평균 오차율: 예측평균(m4~9) vs 실제평균(m4~9)
    pred_avg = np.mean(predicted)
    actual_avg = np.mean(actual_m4_m9)
    avg_error = (pred_avg - actual_avg) / actual_avg * 100

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
            ['A', 'B', 'C'],
            format_func=lambda x: {
                'A': 'A: m1부터 역산',
                'B': 'B: m2부터 역산',
                'C': 'C: m3부터 역산'
            }[x],
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

        # 3가지 방식 비교
        results_all = {}
        for method in ['A', 'B', 'C']:
            results_all[method] = []
            for store in all_stores:
                result = validate_store(store, curve_index, method)
                if result is not None:
                    results_all[method].append({
                        'name': store['name'], **result
                    })

        # 요약 카드
        col1, col2, col3 = st.columns(3)
        method_labels = {'A': 'm1부터', 'B': 'm2부터', 'C': 'm3부터'}
        for col, method in zip([col1, col2, col3], ['A', 'B', 'C']):
            errs = [r['avg_error'] for r in results_all[method]]
            with col:
                if errs:
                    avg_abs_err = np.mean([abs(e) for e in errs])
                    avg_bias = np.mean(errs)
                    st.metric(
                        f"방식 {method} ({method_labels[method]})",
                        f"{avg_abs_err:.2f}%",
                        f"편향 {avg_bias:+.2f}% | {len(errs)}개 매장"
                    )
                else:
                    st.metric(f"방식 {method}", "데이터 없음", "0개 매장")

        # 최적 방식 표시
        valid_methods = [m for m in ['A', 'B', 'C'] if results_all[m]]
        if valid_methods:
            best = min(valid_methods,
                       key=lambda m: np.mean([abs(r['avg_error']) for r in results_all[m]]))
            best_abs = np.mean([abs(r['avg_error']) for r in results_all[best]])
            st.success(
                f"★ 최적 방식: **{best} ({method_labels[best]})** — "
                f"평균 오차율 {best_abs:.2f}%"
            )

        st.markdown("---")

        # 선택된 방식 상세 테이블
        st.subheader(f"📋 방식 {selected_method} 상세 결과")
        results = results_all[selected_method]

        if results:
            df = pd.DataFrame([{
                '매장명': r['name'],
                '기준매출': f"{r['base_revenue']:,.0f}",
                '예측평균(m4-9)': f"{np.mean(r['predicted']):,.0f}",
                '실제평균(m4-9)': f"{np.mean(r['actual']):,.0f}",
                '오차율': f"{r['avg_error']:+.2f}%",
                '|오차율|': f"{abs(r['avg_error']):.2f}%"
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
                xaxis_title="평균 오차율 (%)",
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

            # 3가지 방식 결과
            st.markdown(f"### 📌 {selected_store_name}")
            cols = st.columns(3)
            for i, method in enumerate(['A', 'B', 'C']):
                result = validate_store(store, curve_index, method)
                with cols[i]:
                    st.markdown(f"**방식 {method} ({method_labels[method]})**")
                    if result:
                        st.metric("기준매출", f"{result['base_revenue']:,.0f}원")
                        st.metric("평균오차율", f"{result['avg_error']:+.2f}%")
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
