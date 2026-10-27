import io
import hashlib
import zipfile
from pathlib import Path
import streamlit as st
import pandas as pd
from core import process, build_total, calculate_profile
import altair as alt

st.set_page_config(page_title='电表示数 · 预处理工作台',layout='wide')
st.markdown('''<style>.stApp{background:#f4f7fb}h1,h2,h3{color:#163552}.stMetric{background:white;border-left:4px solid #168b93;padding:16px} .block-container{padding-top:2.5rem}[data-testid="stMetricValue"]{font-size:1.7rem}[data-testid="stMetricLabel"]{white-space:normal}</style>''',unsafe_allow_html=True)
st.caption('江苏电表示数 / 数据准备')
st.title('企业用电分析工作台')
st.write('从累计示数到逐时电量，每块电表独立处理。异常保留痕迹，便于核查。')
def clear_result():
    st.session_state.pop('computed_results',None)
    st.session_state.pop('computed_errors',None)
    st.session_state.pop('computed_signature',None)


def files_changed():
    clear_result()
    st.session_state['use_sample']=False


with st.sidebar:
    st.header('输入设置')
    files=st.file_uploader('上传 Excel，可选择多个文件',type=['xls','xlsx'],accept_multiple_files=True,key='uploaded_files',on_change=files_changed)
    sample=False
    year=st.number_input('备用年份（无查询日期范围时使用）',min_value=2000,max_value=2100,value=2026,on_change=clear_result)
    sheet=st.text_input('工作表',value='正向有功示数',on_change=clear_result)
inputs=[(f.name,f.getvalue()) for f in files]
signature=(tuple((name,hashlib.sha256(data).hexdigest()) for name,data in inputs),int(year),sheet)
if st.session_state.get('computed_signature') != signature:
    clear_result()
if st.button('开始计算',type='primary',disabled=not bool(inputs)):
    results=[[],[],[],[]]
    errors=[]
    with st.spinner('正在预处理并计算…'):
        for name,data in inputs:
            try:
                out=process(data,name,sheet,int(year))
                for dst,frame in zip(results,out):dst.append(frame)
            except Exception as e:
                errors.append(f'{name}：{e}')
    st.session_state['computed_results']=results
    st.session_state['computed_errors']=errors
    st.session_state['computed_signature']=signature
results=st.session_state.get('computed_results',[[],[],[],[]])
for message in st.session_state.get('computed_errors',[]):
    st.error(message)
if results[0]:
    daily,issues,coverage,raw=[pd.concat(r,ignore_index=True) if r else pd.DataFrame() for r in results]
    daily,total=build_total(daily)
    summary,profile,calculation_days=calculate_profile(total)
    if not summary.empty:
        st.subheader('08:00–16:00 用电占比')
        selected=st.selectbox('选择企业（按文件区分）',summary['源文件'].tolist())
        result=summary[summary['源文件']==selected].iloc[0]
        kpis=st.columns(2)+st.columns(2)
        kpis[0].metric('八小时用电占比',f"{result['八小时占比']:.2%}" if pd.notna(result['八小时占比']) else '无法计算')
        kpis[1].metric('八小时日均电量',f"{result['八小时日均电量']:,.2f}" if pd.notna(result['八小时日均电量']) else '—')
        kpis[2].metric('全天日均电量',f"{result['全天日均电量']:,.2f}" if pd.notna(result['全天日均电量']) else '—')
        kpis[3].metric('预计年用电量',f"{result['预计年用电量']:,.2f}" if pd.notna(result['预计年用电量']) else '—')
        st.caption('预计年用电量 = 全天日均电量 × 365；统计天数：'+str(int(result['统计天数']))+' 天。')
        st.caption(f"{result['起始日期']} 至 {result['截止日期']}；排除 {result['排除天数']} 天；包含全天补零 {result['全天补零天数']} 天。电量单位沿用原表。")
        st.info('各小时跨日期日均电量相加（08:00–16:00）÷ 全天日均电量。所有小时及分母使用相同日期；补零日计入天数，存在排除电表的企业日整体不参与计算。')
        view=st.radio('每小时占比图', ['全天 24 小时','仅 08:00–16:00'],horizontal=True)
        chart_data=profile[profile['源文件']==selected]
        if view=='仅 08:00–16:00':
            chart_data=chart_data[chart_data['时段']=='08:00–16:00']
        if pd.notna(result['八小时占比']):
            chart=alt.Chart(chart_data).mark_bar(cornerRadiusTopLeft=3,cornerRadiusTopRight=3).encode(
                x=alt.X('小时区间:N',sort=chart_data['小时区间'].tolist(),title='用电时段',axis=alt.Axis(labelAngle=-45)),
                y=alt.Y('占全天比例:Q',title='占全天电量比例',axis=alt.Axis(format='.1%')),
                color=alt.Color('时段:N',scale=alt.Scale(domain=['08:00–16:00','其他时段'],range=['#168b93','#a9bbcc']),title='时段'),
                tooltip=[alt.Tooltip('小时区间:N'),alt.Tooltip('日均电量:Q',format=',.2f'),alt.Tooltip('占全天比例:Q',format='.2%'),alt.Tooltip('统计天数:Q')]).properties(height=330)
            st.altair_chart(chart,width='stretch')
        else:
            st.warning(result['计算状态'])
        with st.expander('查看逐小时计算明细和企业汇总'):
            st.caption('每小时占比的分母均为全天日均电量；八个目标小时的占比相加等于上方八小时占比。')
            st.dataframe(chart_data,hide_index=True,width='stretch',column_config={'占全天比例':st.column_config.NumberColumn(format='percent')})
            st.dataframe(summary,hide_index=True,width='stretch',column_config={'八小时占比':st.column_config.NumberColumn(format='percent')})
    st.divider()
    st.subheader('预处理与数据追溯')
    cols=st.columns(2)+st.columns(2)
    for c,label,value in zip(cols,['电表数（按文件区分）','电表日记录（含补零）','异常与提示','排除的电表日'],[len(coverage),len(daily),len(issues),int((~daily['纳入汇总']).sum()) if len(daily) else 0]):c.metric(label,value)
    st.warning('缺失日期按零纳入；负值置零并保留修正记录，默认排除此类异常电表日。总表只累计纳入的记录，部分电表被排除的日期会明确标记。')
    t,a,b,c,d=st.tabs(['多电表合并总表','电表与日期覆盖','逐时电量与纳入状态','异常清单','原始示数'])
    with t:
        st.caption('各电表分别差分并乘倍率，再逐日累加为总表。一个文件视为一户；比例计算的纳入状态见下表。')
        st.dataframe(calculation_days,width='stretch',hide_index=True)
    with a:
        st.dataframe(coverage,width='stretch',hide_index=True)
        st.caption('缺失日期依据表头查询范围核查；无查询范围时只核查每块电表首末日期之间的缺口。')
    with b:
        st.caption('1 时电量 = 00:00–01:00；24 时电量 = 23:00–24:00。单位沿用原表，需业务确认后用于正式测算。')
        st.dataframe(daily,width='stretch',hide_index=True)
    with c:
        st.dataframe(issues,width='stretch',hide_index=True)
    with d:st.dataframe(raw,width='stretch',hide_index=True)
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,'w',zipfile.ZIP_DEFLATED) as z:
        for label,frame in [('企业比例汇总',summary),('逐小时日均与占比',profile),('多电表合并总表',calculation_days),('逐时电量',daily),('异常清单',issues),('电表覆盖',coverage),('原始示数',raw)]:
            z.writestr(label+'.csv',frame.to_csv(index=False).encode('utf-8-sig'))
    st.download_button('下载预处理结果与核查清单（CSV 压缩包）',buf.getvalue(),'预处理结果.zip','application/zip')
else:st.info('文件已准备好，请点击“开始计算”。' if inputs else '请先上传文件，再点击“开始计算”。')
with st.expander('处理规则与当前边界',expanded=False):
    st.markdown('''- 仅在明细表头上方读取 L 列电表号与 P 列倍率，支持多个电表；倍率冲突停止该文件。
- 通过完整列名识别 0–24 时，不依赖小时列排列顺序。
- 日期输出为 YYYY-MM-DD；文本月日优先匹配表头查询范围中的唯一日期，无范围时使用备用年份；保留原始日期文本。
- 每小时电量 = 相邻累计示数之差 × 该电表倍率。
- 负差值按零处理；负差值、缺失示数、全天核对不一致、超范围日期和重复电表日默认排除。
- 日期无法定位时，默认阻止该电表进入汇总，避免将损坏记录误当缺日补零。
- 缺失日期按表头查询范围补齐 24 个零；无查询范围时按文件有效日期的首末日期补齐。
- 重复电表日期保留并排除，不自动挑选其中一条。
- 原表总电量存在时进行核对，当前容差为 1 个原表电量单位。
- 当前仅支持江苏此类模板；不自动识别换表、回零、倍率变更或跨文件同户关系。
- 同名文件中的重复电表日自动排除；不同名文件是否属于同户仍需确认。
- 08:00–16:00 对应 9 时至 16 时电量；后续在多电表合并总表上计算。''')
