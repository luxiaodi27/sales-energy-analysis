"""江苏累计电表示数预处理；不推断缺失值，不合并电表。"""
import io
import re
import datetime as dt
from collections import Counter
import pandas as pd


def read_rows(content, filename, sheet):
    if filename.lower().endswith('.xls'):
        import xlrd
        book = xlrd.open_workbook(file_contents=content)
        ws = book.sheet_by_name(sheet)
        rows = []
        for i in range(ws.nrows):
            row = []
            for j in range(ws.ncols):
                c = ws.cell(i, j)
                row.append(xlrd.xldate_as_datetime(c.value, book.datemode)
                           if c.ctype == xlrd.XL_CELL_DATE else c.value)
            rows.append(row)
        return rows
    from openpyxl import load_workbook
    book = load_workbook(io.BytesIO(content), data_only=True, read_only=True)
    try:
        return [list(r) for r in book[sheet].iter_rows(values_only=True)]
    finally:
        book.close()


def text(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def number(v):
    try:
        n = float(str(v).replace(',', '').strip())
        return n if pd.notna(n) and abs(n) != float('inf') else None
    except (ValueError, TypeError):
        return None


def process(content, filename, sheet='正向有功示数', year=2026):
    rows = read_rows(content, filename, sheet)
    issues, records, raw_records = [], [], []
    def issue(kind, row='', meter='', date='', detail=''):
        issues.append(dict(源文件=filename, 类型=kind, 原始行=row,
                           电能表号=meter, 日期=date, 说明=detail))
    header = next((i for i, r in enumerate(rows)
                   if {'电能表号', '日期', *[f'{h}时' for h in range(25)]}
                   <= {text(c) for c in r}), None)
    if header is None:
        raise ValueError('未找到完整表头：需要电能表号、日期、0时至24时。')
    labels = [text(c) for c in rows[header]]
    if any(labels.count(k) != 1 for k in ['电能表号','日期', *[f'{h}时' for h in range(25)]]):
        raise ValueError('必需列存在重复，无法确定对应关系。')
    cols = {k: labels.index(k) for k in ['电能表号','日期', *[f'{h}时' for h in range(25)]]}
    rates = {}
    for i, r in enumerate(rows[:header]):
        if len(r) <= 15:
            continue
        m, rate = text(r[11]), number(r[15])
        if not m or m == '电能表号':
            continue
        if rate is None or rate <= 0:
            issue('倍率无效', i+1, m, detail='综合倍率必须为正数。')
            continue
        if m in rates and rates[m] != rate:
            raise ValueError(f'电表 {m} 有冲突倍率，请先核对原文件。')
        rates[m] = rate
    if not rates:
        raise ValueError('表头区域未找到 L 列电表号与 P 列正数倍率。')
    query = re.search(r'(20\d{2})/(\d{1,2})/(\d{1,2})\s*[-—至]\s*(20\d{2})/(\d{1,2})/(\d{1,2})',
                      ' '.join(text(c) for r in rows[:header] for c in r))
    start = end = None
    if query:
        y,m,d,y2,m2,d2 = map(int,query.groups())
        start,end = dt.date(y,m,d),dt.date(y2,m2,d2)
    # 年度查询首尾月日相同时，使用同一电表的单调日期顺序约束歧义年份。
    resolved_dates={}
    if start and end:
        sequences={}
        for row_number,r in enumerate(rows[header+1:],header+2):
            if max(cols['电能表号'],cols['日期'])>=len(r):
                continue
            meter=text(r[cols['电能表号']]); value=r[cols['日期']]
            if meter not in rates or not isinstance(value,str):
                continue
            parts=[p for p in re.split(r'[./\-年月日]',value.strip()) if p]
            candidates=[]
            try:
                if len(parts)==2:
                    for y in range(start.year,end.year+1):
                        try:
                            candidate=dt.date(y,*map(int,parts))
                        except ValueError:
                            continue
                        if start<=candidate<=end:
                            candidates.append(candidate)
                elif len(parts)==3:
                    candidates=[dt.date(*map(int,parts))]
            except ValueError:
                pass
            sequences.setdefault(meter,[]).append((row_number,candidates))
        for sequence in sequences.values():
            anchors=[c[0] for _,c in sequence if len(c)==1]
            ascending=len(anchors)>=2 and all(a<=b for a,b in zip(anchors,anchors[1:])) and anchors[0]<anchors[-1]
            descending=len(anchors)>=2 and all(a>=b for a,b in zip(anchors,anchors[1:])) and anchors[0]>anchors[-1]
            if not (ascending or descending):
                continue
            for pos,(row_number,candidates) in enumerate(sequence):
                if len(candidates)<=1:
                    continue
                before=next((c[0] for _,c in reversed(sequence[:pos]) if len(c)==1),None)
                after=next((c[0] for _,c in sequence[pos+1:] if len(c)==1),None)
                valid=[c for c in candidates if
                       (before is None or (before<=c if ascending else before>=c)) and
                       (after is None or (c<=after if ascending else c>=after))]
                if len(valid)==1:
                    resolved_dates[row_number]=valid[0]
    for i,r in enumerate(rows[header+1:],header+2):
        get = lambda k: r[cols[k]] if cols[k] < len(r) else None
        m, original_date = text(get('电能表号')), get('日期')
        if not text(original_date):
            if m in rates:
                issue('缺失日期',i,m,detail='保留原文件，未进入标准明细。')
            continue
        if m not in rates:
            issue('倍率未匹配',i,m,text(original_date),'未进入标准明细；不推断电表号或倍率。')
            continue
        raw = dict(源文件=filename,原始行=i,电能表号=m,原始日期=text(original_date))
        raw.update({f'{h}时示数':get(f'{h}时') for h in range(25)})
        raw_records.append(raw)
        try:
            if isinstance(original_date,(dt.datetime,dt.date)):
                date = original_date.date() if isinstance(original_date,dt.datetime) else original_date
            elif isinstance(original_date,str):
                parts = re.split(r'[./\-年月日]',original_date.strip())
                parts = [int(p) for p in parts if p]
                if len(parts)==3:
                    date=dt.date(*parts)
                elif len(parts)==2 and start and end:
                    candidates=[]
                    for candidate_year in range(start.year,end.year+1):
                        try:
                            candidate=dt.date(candidate_year,*parts)
                        except ValueError:
                            continue
                        if start<=candidate<=end:
                            candidates.append(candidate)
                    if len(candidates)==1:
                        date=candidates[0]
                    elif i in resolved_dates:
                        date=resolved_dates[i]
                        issue('跨年日期顺序识别',i,m,date.isoformat(),'月日存在多个候选年份，已通过同电表相邻日期及整体排列方向唯一确定。')
                    else:
                        raise ValueError('查询范围内无法唯一确定月日所属年份')
                else:
                    date=dt.date(year,*parts)
            else:
                raise ValueError('数字日期无法区分 4.1 与 4.10')
        except (ValueError,TypeError):
            issue('日期无效或歧义',i,m,text(original_date),'日期无法唯一确定：请提供完整日期，或核对查询范围及备用年份；数字月日不作猜测。')
            continue
        rec = dict(源文件=filename,原始行=i,电能表号=m,日期=date.isoformat(),
                   原始日期=text(original_date),综合倍率=rates[m])
        vals=[number(get(f'{h}时')) for h in range(25)]
        for h in range(1,25):
            value=None
            if vals[h] is None or vals[h-1] is None:
                issue('缺失或非数字示数',i,m,date.isoformat(),f'{h-1}:00–{h}:00 无法计算；留空。')
            elif vals[h] < vals[h-1]:
                value=0.0
                issue('负差值',i,m,date.isoformat(),f'{h-1}:00–{h}:00 示数倒退；按规则置零，默认排除此电表日。')
            else:
                value=round((vals[h]-vals[h-1])*rates[m],8)
            rec[f'{h}时电量']=value
        rec['有效小时数']=sum(rec[f'{h}时电量'] is not None for h in range(1,25))
        rec['全天电量']=round(sum(rec[f'{h}时电量'] for h in range(1,25)),8) if rec['有效小时数']==24 else None
        total_col=next((j for j,k in enumerate(labels) if '总电量' in k),None)
        source_total=number(r[total_col]) if total_col is not None and total_col<len(r) else None
        rec['原表全天电量']=source_total
        rec['全天核对差额']=round(rec['全天电量']-source_total,8) if rec['全天电量'] is not None and source_total is not None else None
        if rec['全天核对差额'] is not None and abs(rec['全天核对差额'])>1:
            issue('全天电量不一致',i,m,date.isoformat(),f"差分电量与原表相差 {rec['全天核对差额']}，当前核对容差为 1。")
        if start and not start<=date<=end:
            issue('超出查询范围',i,m,date.isoformat(),'记录日期超出表头查询范围。')
        records.append(rec)
    df=pd.DataFrame(records)
    coverage=[]
    if not df.empty:
        dup=df.duplicated(['电能表号','日期'],keep=False)
        df['重复记录']=dup
        for _,r in df[dup].iterrows():
            issue('重复电表日期',r['原始行'],r['电能表号'],r['日期'],'保留记录并标记；后续汇总前必须处理。')
        for m,rate in rates.items():
            sub=df[df['电能表号']==m]
            dates=sorted(set(dt.date.fromisoformat(d) for d in sub['日期']))
            missing=[]
            if dates:
                lo,hi=(start or dates[0]),(end or dates[-1])
                missing=[(lo+dt.timedelta(days=n)).isoformat() for n in range((hi-lo).days+1) if lo+dt.timedelta(days=n) not in set(dates)]
            coverage.append(dict(源文件=filename,电能表号=m,综合倍率=rate,记录数=len(sub),有效日期数=len(dates),起始日期=str(dates[0]) if dates else '',截止日期=str(dates[-1]) if dates else '',缺失日期数=len(missing)))
            if missing:
                issue('日期覆盖不完整',meter=m,detail=f'按表头查询范围（无表头时按电表首末日期）缺 {len(missing)} 天：'+', '.join(missing))
            if not dates:
                issue('电表无有效明细',meter=m,detail='表头存在此电表，但无有效日期记录。')
    else:
        for m,rate in rates.items():
            coverage.append(dict(源文件=filename,电能表号=m,综合倍率=rate,记录数=0,有效日期数=0,起始日期='',截止日期='',缺失日期数=None))
            issue('电表无有效明细',meter=m,detail='无有效记录；先核对日期与明细内容。')
    # 有明确异常的日期不得被“缺日补零”掩盖；无法定位日期的异常阻止整块表汇总。
    fatal_types={'负差值','缺失或非数字示数','全天电量不一致','超出查询范围','重复电表日期'}
    blocked={x['电能表号'] for x in issues if x['类型'] in {'日期无效或歧义','缺失日期'}}
    reasons={}
    for x in issues:
        if x['类型'] in fatal_types:
            reasons.setdefault((x['电能表号'],x['日期']),set()).add(x['类型'])
    if not df.empty:
        df['记录来源']='原始记录'
        df['排除原因']=[('该电表存在无法定位日期的记录' if r['电能表号'] in blocked else
                       '、'.join(sorted(reasons.get((r['电能表号'],r['日期']),set())))) for _,r in df.iterrows()]
        df['纳入汇总']=df['排除原因'].eq('')
    observed=[dt.date.fromisoformat(r['日期']) for r in records]
    lo=start or (min(observed) if observed else None)
    hi=end or (max(observed) if observed else None)
    filled=[]
    if lo and hi:
        for m,rate in rates.items():
            existing={r['日期'] for r in records if r['电能表号']==m}
            for n in range((hi-lo).days+1):
                date=(lo+dt.timedelta(days=n)).isoformat()
                if date in existing:
                    continue
                rec=dict(源文件=filename,原始行=None,电能表号=m,日期=date,原始日期='',综合倍率=rate,
                         有效小时数=24,全天电量=0.0,原表全天电量=None,全天核对差额=None,
                         重复记录=False,记录来源='缺失日期补零',纳入汇总=m not in blocked,
                         排除原因='该电表存在无法定位日期的记录' if m in blocked else '')
                rec.update({f'{h}时电量':0.0 for h in range(1,25)})
                filled.append(rec)
        if filled:
            df=pd.concat([df,pd.DataFrame(filled)],ignore_index=True)
    if not df.empty:
        df=df.sort_values(['日期','电能表号','原始行'],na_position='last').reset_index(drop=True)
    return df,pd.DataFrame(issues,columns=['源文件','类型','原始行','电能表号','日期','说明']),pd.DataFrame(coverage),pd.DataFrame(raw_records)


def build_total(daily):
    """一个文件视为一户；先乘倍率的电表日再汇总为户日总表。"""
    if daily.empty:
        return daily,pd.DataFrame()
    daily=daily.copy()
    # 同文件名、电表、日期的多份上传也须排除，不能重复累加。
    duplicate=daily.duplicated(['源文件','电能表号','日期'],keep=False)
    daily.loc[duplicate,'纳入汇总']=False
    daily.loc[duplicate,'排除原因']='重复电表日期（含跨文件同名重叠）'
    totals=[]
    hours=[f'{h}时电量' for h in range(1,25)]
    for (file,date),group in daily.groupby(['源文件','日期'],sort=True):
        ok=group[group['纳入汇总']]
        rec=dict(源文件=file,日期=date,应有电表数=group['电能表号'].nunique(),
                 纳入电表数=ok['电能表号'].nunique(),排除电表数=group.loc[~group['纳入汇总'],'电能表号'].nunique(),
                 补零电表数=int(ok['记录来源'].eq('缺失日期补零').sum()))
        rec.update({h:round(ok[h].sum(),8) if len(ok) else None for h in hours})
        rec['全天电量']=round(sum(rec[h] for h in hours),8) if len(ok) else None
        rec['总表状态']='全部排除' if not len(ok) else ('部分电表排除' if rec['排除电表数'] else '完整纳入')
        totals.append(rec)
    return daily,pd.DataFrame(totals)


def calculate_profile(total):
    """方法一：同一组完整企业日 -> 逐小时日均 -> 各小时/全天日均。"""
    if total.empty:
        return pd.DataFrame(),pd.DataFrame(),total.copy()
    days=total.copy()
    hours=[f'{h}时电量' for h in range(1,25)]
    days['纳入比例计算']=days['总表状态'].eq('完整纳入') & days[hours].notna().all(axis=1)
    days['比例排除原因']=days['总表状态'].where(~days['纳入比例计算'],'')
    summaries,profiles=[],[]
    for file,group in days.groupby('源文件',sort=True):
        valid=group[group['纳入比例计算']]
        n=len(valid)
        means=valid[hours].mean() if n else pd.Series({h:float('nan') for h in hours})
        denominator=float(means.sum()) if n else None
        numerator=float(means.iloc[8:16].sum()) if n else None
        ratio=numerator/denominator if denominator is not None and denominator>0 else None
        summaries.append(dict(源文件=file,起始日期=group['日期'].min(),截止日期=group['日期'].max(),
                              统计天数=n,排除天数=len(group)-n,
                              全天补零天数=int(valid['补零电表数'].eq(valid['应有电表数']).sum()),
                              八小时日均电量=numerator,全天日均电量=denominator,
                              预计年用电量=denominator*365 if denominator is not None else None,
                              八小时占比=ratio,八小时总电量=float(valid[hours[8:16]].sum().sum()) if n else None,
                              全天总电量=float(valid[hours].sum().sum()) if n else None,
                              计算状态='完成' if ratio is not None else ('无有效日期' if not n else '全天电量为零，比例无法定义')))
        for h in range(1,25):
            avg=float(means.iloc[h-1]) if n else None
            profiles.append(dict(源文件=file,小时序号=h,小时区间=f'{h-1:02d}:00–{h:02d}:00',
                                 日均电量=avg,占全天比例=avg/denominator if denominator is not None and denominator>0 else None,
                                 时段='08:00–16:00' if 9<=h<=16 else '其他时段',统计天数=n))
    return pd.DataFrame(summaries),pd.DataFrame(profiles),days
