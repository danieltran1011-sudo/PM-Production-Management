# -*- coding: utf-8 -*-
"""Bước 2: dựng báo cáo điều hành (tiếng Việt + tiếng Nhật) từ dữ liệu trong WORK_DIR. Chạy qua tao_bao_cao.py."""
import os
import pandas as pd, numpy as np, datetime as dt
VN1=lambda x: f'{x:.1f}'.replace('.',',')
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule, ColorScaleRule
from openpyxl.comments import Comment

OUT=os.environ.get('OUT_FILE','/home/user/PM-Production-Management/bao-cao/Bao_cao_dieu_hanh_PLANTABLE_2026-11-14.xlsx')
import os
REF=pd.Timestamp(os.environ.get('REF_DATE','2026-11-14'))
PL_DATE=pd.Timestamp(os.environ.get('PL_DATE','2026-11-13'))
WORKDIR=os.environ.get('WORK_DIR','.')
o=pd.read_pickle(os.path.join(WORKDIR,'clean.pkl')).sort_values('row').reset_index(drop=True)
o['factory_raw']=o.factory
o.loc[o.factory.isin(['AP1A','AP1B','AP1C']),'factory']='APIC'
def hw(p,h):
    p=str(p) if isinstance(p,str) else ''
    if p.startswith('AB') or p in ('','SPL'): return 0.5
    return 1.0 if p.startswith(h) else 0.0
ORDER=['23SS','23FW','24SS','24FW','25SS','25FW','26SS','26FW','27SS','27FW']

# ---------- flags ----------
def stale(r):
    if pd.isna(r.start) or r.season not in ORDER: return False
    y=2000+int(r.season[:2]); lim=pd.Timestamp(f'{y}-06-30') if r.season.endswith('SS') else pd.Timestamp(f'{y}-12-31')
    return r.start>lim
o['late_season']=o.apply(stale,axis=1)
o['stale']=o.late_season&o.season.isin(['23SS','23FW','24SS','24FW','25SS','25FW'])
o['chk_season']=o.late_season&~o.stale
o['gap']=o.etd_gap.where(o.etd_gap.abs()<365)
o['gap_bad']=o.etd_gap.notna()&o.gap.isna()
def dq(r):
    f=[]
    if r.stale: f.append('Mùa cũ nhưng lịch SX muộn')
    if r.chk_season: f.append('Kiểm tra mã mùa (lịch SX muộn)')
    if r.brand=='(trống)': f.append('Thiếu Brand')
    if pd.isna(r.sam): f.append('Thiếu SAM')
    if pd.isna(r.start): f.append('Thiếu START')
    if r.gap_bad: f.append('Lệch ETD vô nghĩa (thiếu 工場ETD)')
    if pd.notna(r.c2_upc) and r.c2_upc>200: f.append('Đơn giá GC bất thường')
    return '; '.join(f)
o['dq']=o.apply(dq,axis=1)

# ---------- dòng chấm (QTY nền chấm gray0625) = dòng thứ 2 của chuyền lớn, KHÔNG cộng SL ----------
import json, os
_HERE=os.path.dirname(os.path.abspath(__file__))
DOT=set(int(k) for k in json.load(open(os.path.join(WORKDIR,'dotted.json'))))
_raw=pd.read_pickle(os.path.join(WORKDIR,'raw.pkl')).set_index('row')
o['P_code']=o.row.map(_raw['P']); o['I_code']=o.row.map(_raw['I'])
o['dotted']=o.row.isin(DOT); o['qty_raw']=o.qty
_nd=o[~o.dotted]
def _ucalc(AP,SAM,W):
    AP=np.asarray(AP,float); SAM=np.asarray(SAM,float); W=np.asarray(W,float)
    with np.errstate(all='ignore'):
        eff=(3.51*np.log(AP)+76.97)/100; learn=2.23*SAM**0.11; mx=np.floor(28800/SAM*W*eff)
        return np.where(mx>0,(AP/mx+learn/2)/AP,np.nan)
N_ADJ1=int((~o.dotted&((o.I_code.map(_nd.groupby('I_code').qty.sum())-o.c1_qty).abs()>=1)).sum())
for _q,_u,_key in (('c1_qty','c1_upc','I_code'),('c2_qty','c2_upc','P_code')):
    _new=o[_key].map(_nd.groupby(_key).qty.sum())
    _r=_ucalc(_new,o.sam,o.workers)/_ucalc(o[_q],o.sam,o.workers)
    _ok=_new.notna()&o[_q].notna()&((_new-o[_q]).abs()>=1)&np.isfinite(_r)
    o.loc[_ok,_u]=o.loc[_ok,_u]*_r[_ok]; o.loc[_ok,_q]=_new[_ok]
O_DOT=o[o.dotted].copy(); o=o[~o.dotted].reset_index(drop=True)
# ---------- Hàng SMS (mẫu chào hàng / triển lãm, hỗ trợ chuyền mẫu) – tách khỏi hàng sản xuất ----------
_key=(o['style'].fillna('')+' '+o['contract'].fillna('')+' '+o['item'].fillna('')).str.upper(); _com=o['comment'].fillna('')
o['sms']=(_key.str.contains('SMS')|(_com.str.contains('SMS',case=False)&~_com.str.contains('SMSサンプル対応'))|_com.str.contains('展示会サンプル|展示会投入'))&(o.qty<=500)
O_SMS=o[o.sms].copy(); o=o[~o.sms].reset_index(drop=True)
def _wdays(s,e):
    if pd.isna(s) or pd.isna(e) or e<s or (e-s).days>200: return 0
    return int(sum(1 for d in pd.date_range(s,e) if d.dayofweek<6))
O_SMS['cd']=[(w if pd.notna(w) else 0)*_wdays(s,e) for w,s,e in zip(O_SMS.workers,O_SMS.start,O_SMS.end)]
# ---------- K: số liệu dùng trong văn bản ----------
_sq=o.groupby('season').qty.sum()
K={}
K['y24']=_sq['24SS']+_sq['24FW']; K['y25']=_sq['25SS']+_sq['25FW']; K['y26']=_sq['26SS']+_sq['26FW']
K['g25']=K['y25']/K['y24']-1; K['g26']=K['y26']/K['y25']-1; K['s27']=_sq['27SS']; K['g27']=_sq['27SS']/_sq['26SS']-1
K['s27fw']=_sq['27FW']; K['s26fw']=_sq['26FW']
_sh=lambda d: (d.qty*d.sam/3600).sum()
K['h_lo']=min(_sh(o[o.season.isin([f'{y}SS',f'{y}FW'])]) for y in ('23','24','25','26'))/1e6
K['h_hi']=max(_sh(o[o.season.isin([f'{y}SS',f'{y}FW'])]) for y in ('23','24','25','26'))/1e6
_samS=[_sh(o[o.season==s])*60/_sq[s] for s in ORDER[:9]]
K['samFW_lo']=min(_samS[1::2]); K['samFW_hi']=max(_samS[1::2]); K['samSS_lo']=min(_samS[0::2]); K['samSS_hi']=max(_samS[0::2])
_bq=o.groupby('brand').qty.sum().sort_values(ascending=False)
K['tnf']=_bq['TNF']/_bq.sum(); K['top3']=_bq.head(3).sum()/_bq.sum()
_s27=o[o.season=='27SS']; K['tnf27']=_s27[_s27.brand=='TNF'].qty.sum()/_s27.qty.sum()
_mb=o[o.brand=='MONTBELL'].groupby('season').qty.sum(); K['mb']=[_mb.get(s,0)/1000 for s in ('24FW','25FW','26FW','27SS')]
_hh=o[o.brand=='HELLY HANSEN'].groupby('season').qty.sum(); K['hh']=_hh.get('27SS',0)/_hh.get('26SS',1)-1
_f26=o[o.season=='26FW']; K['bes26']=_f26[_f26.bespoke].qty.sum()/_f26.qty.sum()
_lb=o.c2_qty
_sm=o[_lb<=300]; _md=o[(_lb>2000)&(_lb<=5000)]
_u=lambda d: (d.qty*d.c2_upc).sum()/d[d.c2_upc.notna()].qty.sum()
K['lot_rows']=(_lb<=300).sum()/_lb.notna().sum(); K['lot_qty']=_sm.qty.sum()/o[_lb.notna()].qty.sum()
K['upc_s']=_u(_sm); K['upc_m']=_u(_md); K['lot_x']=K['upc_s']/K['upc_m']; K['eff_s']=(_sm.c2_avg/_sm.c2_max).mean()
_c1=(o.qty*o.c1_upc); _c2=(o.qty*o.c2_upc); K['split']=(_c2-_c1).sum()/1e6
_ps=lambda s: _c2[o.season==s].sum()/_c1[o.season==s].sum()-1
K['sp25']=_ps('25SS'); K['sp27']=_ps('27SS'); K['sp26fw']=_ps('26FW')
K['small26fw']=(o[o.season=='26FW'].c2_qty<=300).mean()
K['blank']=o[o.brand=='(trống)'].qty.sum()/1000
K['ndot']=len(O_DOT); K['qdot']=O_DOT.qty_raw.sum()/1000
K['nsms']=len(O_SMS); K['qsms']=O_SMS.qty.sum(); K['qsms_vn']=f'{O_SMS.qty.sum():,.0f}'.replace(',','.'); K['cdsms']=O_SMS.cd.sum()
V=lambda x,n=1: f'{x:.{n}f}'.replace('.',',')
import importlib.util as _ilu
_sp=_ilu.spec_from_file_location('juchu_plan',os.path.join(_HERE,'juchu_plan.py')); JP=_ilu.module_from_spec(_sp); _sp.loader.exec_module(JP)
_vnote=lambda d: d.row.map(_raw['V'])
_rows=pd.concat([o.assign(kind='SX'),O_DOT.assign(kind='DOT'),O_SMS.assign(kind='SMS')],ignore_index=True)
_rows=_rows[~_rows.stale].copy()
_rows['conf']=(_rows.start<=REF)|_vnote(_rows).map(JP.mat_ok_from_note)
JE=JP.expand(_rows,REF); JC=JP.coverage(JE)
_m0=REF.to_period('M'); JMONTHS=pd.period_range(_m0-2,_m0+9,freq='M')
JMON=JP.monthly(JC,JMONTHS)
JPER=[]
for _m in pd.period_range(_m0-1,_m0+9,freq='M'):
    JPER.append((_m.start_time,_m.start_time+pd.Timedelta(days=14),f'{_m.month:02d}/{str(_m.year)[2:]} 前'))
    JPER.append((_m.start_time+pd.Timedelta(days=15),_m.end_time.normalize(),f'{_m.month:02d}/{str(_m.year)[2:]} 後'))
JGRID=JP.grid(JE,JPER)
_JA=JP.allocate(JE[(JE.d>=(_m0+1).start_time)&(JE.d<=(_m0+9).end_time)]); _JA['m']=_JA.d.dt.to_period('M')
_JA['sb']=_JA.season.astype(str)+' '+_JA.brand.astype(str); _JA.loc[_JA.kind=='SMS','sb']=_JA.season.astype(str)+' SMS'
JSB=_JA.pivot_table(index='sb',columns='m',values='alloc',aggfunc='sum').fillna(0)
JSB['tot']=JSB.sum(axis=1); JSB['tent']=_JA[~_JA.conf].groupby('sb').alloc.sum().reindex(JSB.index).fillna(0); JSB=JSB.sort_values('tot',ascending=False)
JLN=JP.per_line(JE,REF)
_JY=JP.allocate(JE); _JY['y']=_JY.d.dt.year
JHB=_JY[(_JY.y>=2023)&(_JY.y<=REF.year)].pivot_table(index='brand',columns='y',values='alloc',aggfunc='sum').fillna(0)
# lịch sử bản chụp
HIST_CSV=os.environ.get('HIST_CSV',os.path.join(_HERE,'..','lich-su','do_phu_don_theo_ban_chup.csv'))
_h=pd.read_csv(HIST_CSV,dtype=str) if os.path.exists(HIST_CSV) else pd.DataFrame(columns=['ban_chup','nguon','thang','cong_suat','da_chot','tong'])
_snap=os.environ.get('SNAP_DATE',PL_DATE.date().isoformat())
_h=_h[~((_h.ban_chup==_snap)&(_h.nguon=='PLANTABLE'))]
_new=pd.DataFrame([dict(ban_chup=_snap,nguon='PLANTABLE',thang=str(x.m),cong_suat=x.cap,da_chot=round(x.conf,1),tong=round(x.tot,1)) for x in JMON.itertuples() if x.m>=_m0])
_h=pd.concat([_h,_new.astype(str)],ignore_index=True).sort_values(['ban_chup','nguon','thang'])
if os.environ.get('WRITE_HIST','1')=='1': _h.to_csv(HIST_CSV,index=False,encoding='utf-8')
for _c in ('cong_suat','da_chot','tong'): _h[_c]=_h[_c].astype(float)
_w1=[_m0+1,_m0+3]; _w2=[_m0+4,_m0+7]
def _win(g,w):
    ms=[str(p) for p in pd.period_range(w[0],w[1],freq='M')]; x=g[g.thang.isin(ms)]
    if len(x)<len(ms): return (np.nan,np.nan)
    return (x.tong.sum()/x.cong_suat.sum(), x.da_chot.sum()/x.cong_suat.sum())
_tr=[]
for (bd,src),g in _h.groupby(['ban_chup','nguon']):
    t1,c1=_win(g,_w1); t2,c2=_win(g,_w2); _tr.append(dict(date=pd.Timestamp(bd),src=src,t1=t1,c1=c1,t2=t2,c2=c2))
JTR=pd.DataFrame(_tr).sort_values(['date','src'])
_cur=JTR[(JTR.src=='PLANTABLE')&(JTR.date==pd.Timestamp(_snap))].iloc[0]
K['j_t1']=float(_cur.t1); K['j_c1']=float(_cur.c1); K['j_t2']=float(_cur.t2); K['j_date']=pd.Timestamp(_snap)
_pk=JTR.dropna(subset=['t2']); K['j_pk']=float(_pk.t2.max()); K['j_pkd']=_pk.loc[_pk.t2.idxmax(),'date']
K['j_w1']=f'T{_w1[0].month}/{_w1[0].year}–T{_w1[1].month}/{_w1[1].year}'; K['j_w2']=f'T{_w2[0].month}/{_w2[0].year}–T{_w2[1].month}/{_w2[1].year}'
_g=lambda b: (lambda s: s.get('27SS',0)/s.get('26SS',np.nan)-1)(o[o.brand==b].groupby('season').qty.sum())
K['gw']=_g('GOLDWIN'); K['at']=_g('ALLTERRAIN 81')
_gp=o.etd_gap.where(o.etd_gap.abs()<365); _lr=(_gp<0).groupby(o.season).mean()
K['late27']=_lr['27SS']; K['late_lo']=_lr[['24FW','25SS','25FW','26SS','26FW']].min(); K['late_hi']=_lr[['24FW','25SS','25FW','26SS','26FW']].max()

# ---------- styles ----------
AR='Arial'
F=lambda **k: Font(name=AR, **k)
HDR_FILL=PatternFill('solid',fgColor='1C5CAB'); HDR_FONT=F(bold=True,color='FFFFFF',size=10)
SUB_FILL=PatternFill('solid',fgColor='EAF1FB'); TOT_FILL=PatternFill('solid',fgColor='F2F2F2')
thin=Side(style='thin',color='D0D0D0'); BOX=Border(left=thin,right=thin,top=thin,bottom=thin)
BLUE=F(color='0000FF',size=10); GREEN=F(color='008000',size=10); BODY=F(size=10)
NUM='#,##0;(#,##0);-'; USD='$#,##0;($#,##0);-'; USD2='$#,##0.00;($#,##0.00);-'; PCT='0.0%;(0.0%);-'; DATE='yyyy-mm-dd'
def hdr(ws,row,cols,start=1,h=30):
    for i,c in enumerate(cols):
        x=ws.cell(row,start+i,c); x.font=HDR_FONT; x.fill=HDR_FILL; x.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center'); x.border=BOX
    ws.row_dimensions[row].height=h
def title(ws,t,sub=None):
    ws['A1']=t; ws['A1'].font=F(bold=True,size=14,color='1C5CAB')
    if sub: ws['A2']=sub; ws['A2'].font=F(italic=True,size=9,color='52514E')
def widths(ws,w):
    for i,x in enumerate(w,1): ws.column_dimensions[CL(i)].width=x
def body(ws,r1,r2,c1,c2):
    for r in range(r1,r2+1):
        for c in range(c1,c2+1):
            x=ws.cell(r,c); x.border=BOX
            if x.font.name!=AR: x.font=BODY
def chart_fix(ch):
    ch.roundedCorners=False
    if ch.legend is not None: ch.legend.overlay=False
    if ch.title is not None: ch.title.overlay=False
    for ax in (ch.x_axis,ch.y_axis):
        ax.delete=False
        if ax.title is not None: ax.title.overlay=False

wb=Workbook()
S1=wb.active; S1.title='1. Tóm tắt điều hành'
S2=wb.create_sheet('2. Mùa vụ'); S3=wb.create_sheet('3. Khách hàng'); S4=wb.create_sheet('4. Công suất')
S5=wb.create_sheet('5. Chi phí & Lô'); S6=wb.create_sheet('6. Rủi ro giao hàng'); S7=wb.create_sheet('7. Chất lượng DL')
SD=wb.create_sheet('Data'); SM=wb.create_sheet('Phương pháp')

# ---------- Phương pháp (inputs) ----------
title(SM,'Phương pháp, giả định & ánh xạ cột','Ô chữ xanh dương = giá trị đầu vào có thể chỉnh; mọi bảng khác tự tính lại theo các ô này.')
SM['A4']='Ngày tham chiếu (cắt dữ liệu)'; SM['C4']=REF.to_pydatetime(); SM['C4'].number_format=DATE; SM['C4'].font=BLUE
SM['D4']='File gốc lưu ngày 2026-11-13; phân tích lấy 2026-11-14 làm "hôm nay".'
SM['A5']='Định biên CN may xưởng APIC'; SM['C5']=960; SM['C5'].font=BLUE; SM['D5']='30 line thường (1A, 1B … 15A, 15B) × 32 CN. Line nhỏ 小ロットライン (vd 14A1, 14A2) = 16 CN, 2 line nhỏ = 1 line thường. AP1A/AP1B/AP1C cũ đã gộp thành APIC. Tham số 縫直 trong file gốc = 947.'
SM['A9']='Chi phí nhân công / công-ngày (USD)'; SM['C9']=46.29017951425554; SM['C9'].font=BLUE; SM['C9'].number_format=USD2; SM['D9']='Lấy từ file gốc AT5: 工賃単価 × 総人員 / 縫直 / 25 ngày = (760×1442)/947/25.'
SM['A10']='Mức tải mục tiêu'; SM['C10']=0.9; SM['C10'].font=BLUE; SM['C10'].number_format='0%'; SM['D10']='Dùng để tính công-ngày cần bổ sung.'
SM['A6']='Ngưỡng "Nghiêm trọng" (ngày trễ ETD)'; SM['C6']=-30; SM['C6'].font=BLUE
SM['A7']='Ngưỡng "Cao" (ngày trễ ETD)'; SM['C7']=-8; SM['C7'].font=BLUE
SM['A8']='Ngưỡng lô nhỏ (pcs)'; SM['C8']=300; SM['C8'].font=BLUE
for r in range(4,11): SM.cell(r,1).font=F(bold=True,size=10); SM.cell(r,4).font=F(size=9,italic=True,color='52514E')
SM['A12']='Ánh xạ cột (sheet gốc "PLANTABLE (原価) 2026年", header dòng 9, dữ liệu từ dòng 10)'; SM['A12'].font=F(bold=True,size=11)
mp=[('Status','A','Mùa vụ (chuẩn hoá: bỏ dấu ", khoảng trắng)'),('Factory','B','Xưởng (chuẩn hoá ký tự full-width)'),('LINE / ﾗｲﾝ数','C / E','Line và số công nhân/line'),
('BRAND','G','Brand (gộp "TNF\\n別注" → TNF + cờ 別注; "HH" → HELLY HANSEN)'),('STYLE # / ITEM','I / M','Mã hàng / tên hàng'),('QTTY','Q','Số lượng đơn'),
('ｺﾒﾝﾄ (công thức =U−AJ)','S','Lệch ETD = 工場ETD yêu cầu − ETD kế hoạch. Âm = kế hoạch trễ hơn yêu cầu. Giá trị |x|>365 bị loại (do 工場ETD trống).'),
('工場ETD','U','ETD xưởng được yêu cầu'),('ETA HCM / 投入可能日','W / X','Ngày NVL về HCM / ngày có thể đưa vào SX (=W+21)'),
('sample sent / reply / giai đoạn','Z / AA / AB','Trạng thái mẫu (PPS, SMS, 1st…)'),('TIME','AC','SAM – tổng thời gian chuẩn (giây/sp)'),
('START / END / ETD','AH / AI / AJ','Lịch may; ETD kế hoạch = END+10'),('標準原価1','AN–AY','Giá vốn gia công theo điều kiện đặt hàng (lô = tổng SL của style)'),
('標準原価2','AZ–BK','Giá vốn gia công theo lô sản xuất thực tế (gộp theo 集約CODE)'),('(không tên)','BL','Không có header – không dùng'),('END+38 / chênh lệch','BO / BP','Chỉ có dữ liệu cho 234 dòng 2024–25 – không dùng')]
hdr(SM,13,['Tên cột gốc','Cột','Ý nghĩa / xử lý'])
for i,(a,b,c) in enumerate(mp,14):
    SM.cell(i,1,a); SM.cell(i,2,b); SM.cell(i,3,c); SM.cell(i,3).alignment=Alignment(wrap_text=True)
body(SM,14,13+len(mp),1,3)
r=15+len(mp)
SM.cell(r,1,'Định nghĩa chỉ số').font=F(bold=True,size=11)
defs=[('Giờ chuẩn','Bằng QTY × SAM / 3600'),('CP gia công 1 / 2','Bằng QTY × đơn giá gia công (USD/pcs) của giá vốn 1 / 2'),
('Chi phí chia lô','Bằng CP2 − CP1: phần chi phí tăng thêm vì lô SX thực tế nhỏ hơn lô đặt hàng (kém hiệu suất học việc)'),
('Hiệu suất lô','Bằng 平均日産 / 日産MAX (giá vốn 2) – lô càng nhỏ càng không đạt tốc độ tối đa'),
('Tỷ lệ sử dụng ngày công (APIC)','Bằng Σ(ﾗｲﾝ数 × ngày làm việc T2–T7 từ START→END) / (định biên × ngày làm việc). Chưa trừ nghỉ lễ/Tết.'),
('Đơn đang/ sắp SX','END ≥ ngày tham chiếu, loại các dòng "mùa cũ nhưng lịch SX muộn"'),
('Hàng SMS','Mẫu chào hàng / mẫu triển lãm hỗ trợ chuyền mẫu (style/hợp đồng/item chứa SMS, hoặc ghi chú SMS / 展示会サンプル / 展示会投入; trừ "SMSサンプル対応"; SL ≤500). Tách khỏi hàng sản xuất: không tính vào sản lượng, chi phí, lô, rủi ro; vẫn tính chiếm line. Xem bảng riêng ở sheet 2 và 1C.'),
('Dòng chấm','Ô QTY nền chấm = dòng thứ 2 của chuyền lớn (vd 4AB, 5AB). Không cộng SL, giờ chuẩn, chi phí; chỉ dùng để tính chiếm line. Lô 原価1/原価2 được tính lại sau khi loại dòng chấm (điều chỉnh theo cùng công thức của file gốc).'),
('Số style, lô trung vị, ngày công kế hoạch','Tính bằng Python từ dữ liệu sạch (giá trị cố định); các bảng còn lại dùng công thức SUMIFS/COUNTIFS trên sheet Data')]
for i,(a,b) in enumerate(defs,r+1):
    SM.cell(i,1,a).font=F(bold=True,size=10); SM.cell(i,3,b).font=BODY; SM.cell(i,3).alignment=Alignment(wrap_text=True)
widths(SM,[34,12,95,60])

# ---------- Data ----------
cols=[('Dòng gốc',8),('Mùa',7),('Xưởng',7),('Line',5),('Part',5),('Số CN',6),('Brand',16),('Brand gốc',14),('別注',5),('Contract',12),
('Style',14),('Item',26),('QTY',8),('SAM (s/sp)',9),('Giờ chuẩn',10),('START',11),('END',11),('ETD KH',11),('工場ETD',11),('Lệch ETD (ngày)',9),
('NVL về HCM',11),('Có thể đưa vào SX',11),('Giai đoạn mẫu',9),('Mẫu gửi',11),('Mẫu phản hồi',11),('Lô SX (GV2)',9),('Lô đặt (GV1)',9),
('Đơn giá GC1 (USD/pcs)',10),('Đơn giá GC2 (USD/pcs)',10),('CP GC1 (USD)',12),('CP GC2 (USD)',12),('Chi phí chia lô (USD)',12),('Nhóm lô',12),
('Hiệu suất lô',8),('Đang/sắp SX',8),('Cờ chất lượng DL',30),('Loại trừ (mùa cũ)',8),('Mã xưởng cũ',8),('Số ngày may',8),('Nhóm SAM (thời gian chế tạo)',14),('Nhóm ngày may',12)]
hdr(SD,1,[c for c,_ in cols],h=42); widths(SD,[w for _,w in cols])
def v(x):
    if x is None: return None
    if isinstance(x,float) and np.isnan(x): return None
    if isinstance(x,pd.Timestamp): return None if pd.isna(x) else x.to_pydatetime()
    if x is pd.NaT: return None
    if isinstance(x,(dt.datetime,dt.date)): return x
    return x
def sv(x):
    if isinstance(x,(dt.datetime,pd.Timestamp)): return v(pd.Timestamp(x)) if x.year>2000 else None
    if isinstance(x,str): return x.replace('\n',' ')
    return None
N=len(o)
for i,r in enumerate(o.itertuples(),2):
    vals=[r.row,r.season,r.factory,v(r.line),r.part,v(r.workers),r.brand,r.brand_raw,'Y' if r.bespoke else '',r.contract,r.style,r.item,v(r.qty),v(r.sam),
          f'=IF(OR(M{i}="",N{i}=""),"",M{i}*N{i}/3600)',v(r.start),v(r.end),v(r.etd_plan),v(r.fac_etd),v(r.gap),v(r.mat_eta),
          v(r.input_ok) if pd.notna(r.input_ok) else r.input_txt,r.smp_stage,sv(r.smp_sent),sv(r.smp_reply),v(r.c2_qty),v(r.c1_qty),v(r.c1_upc),v(r.c2_upc),
          f'=IF(OR(M{i}="",AB{i}=""),"",M{i}*AB{i})',f'=IF(OR(M{i}="",AC{i}=""),"",M{i}*AC{i})',f'=IF(OR(AD{i}="",AE{i}=""),"",AE{i}-AD{i})',
          f'=IF(Z{i}="","",IF(Z{i}<=300,"1. ≤300",IF(Z{i}<=500,"2. 301–500",IF(Z{i}<=1000,"3. 501–1.000",IF(Z{i}<=2000,"4. 1.001–2.000",IF(Z{i}<=5000,"5. 2.001–5.000","6. >5.000"))))))',
          v(r.c2_avg/r.c2_max) if pd.notna(r.c2_avg) and r.c2_max else None,
          f'=IF(AND(Q{i}<>"",Q{i}>=\'Phương pháp\'!$C$4,AJ{i}=""),"Có","")' if not r.stale else f'=IF(AND(Q{i}<>"",Q{i}>=\'Phương pháp\'!$C$4,AJ{i}=""),"Có","")',
          r.dq or None,'Y' if r.stale else None,r.factory_raw,
          f'=IF(AND(ISNUMBER(P{i}),ISNUMBER(Q{i})),Q{i}-P{i}+1,"")',
          f'=IF(N{i}="","",IF(N{i}/60<60,"1. <60 phút",IF(N{i}/60<120,"2. 60–119 phút",IF(N{i}/60<180,"3. 120–179 phút",IF(N{i}/60<240,"4. 180–239 phút","5. ≥240 phút")))))',
          f'=IF(AM{i}="","",IF(AM{i}<=5,"1. ≤5 ngày",IF(AM{i}<=10,"2. 6–10 ngày",IF(AM{i}<=20,"3. 11–20 ngày",IF(AM{i}<=40,"4. 21–40 ngày","5. >40 ngày")))))']
    for j,x in enumerate(vals,1):
        c=SD.cell(i,j,x); c.font=BODY
    for j in (16,17,18,19,21,22,24,25): SD.cell(i,j).number_format=DATE
    for j in (13,14,15,26,27,30,31,32): SD.cell(i,j).number_format=NUM
    for j in (28,29): SD.cell(i,j).number_format=USD2
    SD.cell(i,34).number_format='0%'
# fix: forward flag must ignore rows with DQ stale flag -> use column AJ (cờ DQ) contains "Mùa cũ"
for i in range(2,N+2):
    SD.cell(i,35).value=f'=IF(AND(Q{i}<>"",Q{i}>=\'Phương pháp\'!$C$4,AK{i}<>"Y"),"Có","")'
SD.freeze_panes='C2'; SD.auto_filter.ref=f'A1:AO{N+1}'
LAST=N+1
R=lambda col: f"Data!${col}$2:${col}${LAST}"

# ---------- 2. Mùa vụ ----------
title(S2,'Tổng hợp theo mùa vụ – hàng sản xuất','Nguồn: sheet Data (SUMIFS/COUNTIFS). Chỉ hàng sản xuất: đã bỏ dòng chấm và tách hàng SMS (bảng riêng bên dưới). Số style & lô trung vị tính sẵn từ dữ liệu sạch.')
h2=['Mùa','Số dòng đơn','Số style','Sản lượng (pcs)','Tăng/giảm vs cùng mùa năm trước','Giờ chuẩn','SAM TB (phút/sp)','CP GC1 – theo đặt hàng (USD)','CP GC2 – theo lô SX (USD)','Chi phí chia lô (USD)','% chi phí chia lô','Đơn giá GC2 (USD/pcs)','Lô SX trung vị (pcs)','Tỷ trọng hàng 別注']
hdr(S2,4,h2,h=45)
for k,sn in enumerate(ORDER):
    r=5+k; sub=o[o.season==sn]
    S2.cell(r,1,sn)
    S2.cell(r,2,f'=COUNTIFS({R("B")},$A{r})')
    S2.cell(r,3,int(sub['style'].nunique()))
    S2.cell(r,4,f'=SUMIFS({R("M")},{R("B")},$A{r})')
    S2.cell(r,5,f'=IFERROR(D{r}/D{r-2}-1,"")' if k>=2 else 'n/a')
    S2.cell(r,6,f'=SUMIFS({R("O")},{R("B")},$A{r})')
    S2.cell(r,7,f'=IFERROR(F{r}*60/D{r},"")')
    S2.cell(r,8,f'=SUMIFS({R("AD")},{R("B")},$A{r})')
    S2.cell(r,9,f'=SUMIFS({R("AE")},{R("B")},$A{r})')
    S2.cell(r,10,f'=I{r}-H{r}'); S2.cell(r,11,f'=IFERROR(J{r}/H{r},"")')
    S2.cell(r,12,f'=IFERROR(I{r}/SUMIFS({R("M")},{R("B")},$A{r},{R("AE")},"<>"),"")')
    S2.cell(r,13,float(sub.c2_qty.median()))
    S2.cell(r,14,f'=IFERROR(SUMIFS({R("M")},{R("B")},$A{r},{R("I")},"Y")/D{r},"")')
tr=5+len(ORDER)
S2.cell(tr,1,'Tổng'); 
for c,fn in [(2,'SUM'),(3,None),(4,'SUM'),(6,'SUM'),(8,'SUM'),(9,'SUM'),(10,'SUM')]:
    if fn: S2.cell(tr,c,f'=SUM({CL(c)}5:{CL(c)}{tr-1})')
S2.cell(tr,3,int(o['style'].nunique())); S2.cell(tr,7,f'=F{tr}*60/D{tr}'); S2.cell(tr,11,f'=J{tr}/H{tr}')
body(S2,5,tr,1,14)
for r in range(5,tr+1):
    for c,fm in [(2,NUM),(3,NUM),(4,NUM),(5,PCT),(6,NUM),(7,'0.0'),(8,USD),(9,USD),(10,USD),(11,PCT),(12,USD2),(13,NUM),(14,PCT)]: S2.cell(r,c).number_format=fm
    S2.cell(r,1).font=F(bold=True,size=10)
for c in range(1,15): S2.cell(tr,c).fill=TOT_FILL; S2.cell(tr,c).font=F(bold=True,size=10)
S2.cell(tr+1,1,'Ghi chú: 27FW mới bắt đầu nhận đơn (chưa đầy đủ). "Cùng mùa năm trước" = so 27SS với 26SS, 26FW với 25FW…').font=F(italic=True,size=9,color='52514E')
# yearly
yr=tr+3; S2.cell(yr,1,'Theo năm (SS + FW)').font=F(bold=True,size=11)
hdr(S2,yr+1,['Năm','Sản lượng (pcs)','Tăng/giảm','Giờ chuẩn','CP GC2 (USD)','Chi phí chia lô (USD)'])
for k,y in enumerate(['23','24','25','26']):
    r=yr+2+k; ss=5+ORDER.index(y+'SS'); fw=ss+1
    S2.cell(r,1,'20'+y); S2.cell(r,2,f'=D{ss}+D{fw}'); S2.cell(r,3,f'=IFERROR(B{r}/B{r-1}-1,"")' if k else 'n/a')
    S2.cell(r,4,f'=F{ss}+F{fw}'); S2.cell(r,5,f'=I{ss}+I{fw}'); S2.cell(r,6,f'=J{ss}+J{fw}')
    for c,fm in [(2,NUM),(3,PCT),(4,NUM),(5,USD),(6,USD)]: S2.cell(r,c).number_format=fm
body(S2,yr+2,yr+5,1,6)
_r=yr+8; S2.cell(_r,1,'Hàng SMS hỗ trợ chuyền mẫu (tách riêng, không tính trong các bảng trên)').font=F(bold=True,size=11)
hdr(S2,_r+1,['Mùa','Số dòng SMS','SL SMS (pcs)','Công-ngày dùng cho SMS','Brand chính'],h=30)
_g=O_SMS.groupby('season')
for k_,sn in enumerate(ORDER):
    rr_=_r+2+k_; d_=O_SMS[O_SMS.season==sn]
    S2.cell(rr_,1,sn).font=F(bold=True,size=10); S2.cell(rr_,2,len(d_)); S2.cell(rr_,3,float(d_.qty.sum())); S2.cell(rr_,4,float(d_.cd.sum()))
    S2.cell(rr_,5,', '.join(d_.groupby('brand').qty.sum().sort_values(ascending=False).head(3).index) if len(d_) else '')
    for c_ in (2,3,4): S2.cell(rr_,c_).number_format=NUM
rr_=_r+2+len(ORDER); S2.cell(rr_,1,'Tổng'); S2.cell(rr_,2,f'=SUM(B{_r+2}:B{rr_-1})'); S2.cell(rr_,3,f'=SUM(C{_r+2}:C{rr_-1})'); S2.cell(rr_,4,f'=SUM(D{_r+2}:D{rr_-1})')
for c_ in (2,3,4): S2.cell(rr_,c_).number_format=NUM
body(S2,_r+2,rr_,1,5)
for c_ in range(1,6): S2.cell(rr_,c_).fill=TOT_FILL; S2.cell(rr_,c_).font=F(bold=True,size=10)
S2.cell(rr_+1,1,'Nhận diện SMS: style/hợp đồng/item chứa "SMS", hoặc ghi chú có SMS / 展示会サンプル / 展示会投入 (trừ "SMSサンプル対応"), SL ≤500 pcs.').font=F(italic=True,size=9,color='52514E')
widths(S2,[9,9,8,13,13,12,10,15,15,14,10,11,10,10])
S2.freeze_panes='B5'
ch=BarChart(); ch.type='col'; ch.title='Sản lượng theo mùa (pcs)'; ch.y_axis.title='pcs'; ch.y_axis.numFmt='#,##0'
ch.add_data(Reference(S2,min_col=4,min_row=4,max_row=tr-1),titles_from_data=True); ch.set_categories(Reference(S2,min_col=1,min_row=5,max_row=tr-1))
ch.series[0].graphicalProperties.solidFill='2A78D6'; ch.series[0].graphicalProperties.line.noFill=True; ch.legend=None; ch.gapWidth=60
ch.height=7.5; ch.width=17; chart_fix(ch); S2.add_chart(ch,f'H{yr}')
ch2=BarChart(); ch2.type='col'; ch2.title='Chi phí chia lô theo mùa (USD)'; ch2.y_axis.numFmt='$#,##0'; ch2.y_axis.title='USD'
ch2.add_data(Reference(S2,min_col=10,min_row=4,max_row=tr-1),titles_from_data=True); ch2.set_categories(Reference(S2,min_col=1,min_row=5,max_row=tr-1))
ch2.series[0].graphicalProperties.solidFill='EB6834'; ch2.series[0].graphicalProperties.line.noFill=True; ch2.legend=None; ch2.gapWidth=60
ch2.height=7.5; ch2.width=17; chart_fix(ch2); S2.add_chart(ch2,f'H{yr+16}')

# ---------- 3. Khách hàng ----------
title(S3,'Cơ cấu khách hàng / brand','Sản lượng (pcs) theo brand × mùa. "Khác" = phần còn lại. Brand trống = '+f'{K["blank"]:.0f}k pcs (xem sheet 7).')
tot=o.groupby('brand').qty.sum().sort_values(ascending=False)
top=[b for b in tot.index if b!='(trống)'][:10]
h3=['Brand']+ORDER+['Tổng','Tỷ trọng','Tăng/giảm 27SS vs 26SS','Đơn giá GC2 (USD/pcs)','SAM TB (phút/sp)']
hdr(S3,4,h3,h=45)
for k,b in enumerate(top+['(trống)']):
    r=5+k; S3.cell(r,1,b)
    for j,sn in enumerate(ORDER): S3.cell(r,2+j,f'=SUMIFS({R("M")},{R("G")},$A{r},{R("B")},{CL(2+j)}$4)')
    S3.cell(r,12,f'=SUM(B{r}:K{r})')
rk=5+len(top)+1; S3.cell(rk,1,'Khác')
for j in range(len(ORDER)):
    c=CL(2+j); S3.cell(rk,2+j,f"=SUMIFS({R('M')},{R('B')},{c}$4)-SUM({c}5:{c}{rk-1})")
S3.cell(rk,12,f'=SUM(B{rk}:K{rk})')
rt=rk+1; S3.cell(rt,1,'Tổng')
for j in range(len(ORDER)+1): c=CL(2+j); S3.cell(rt,2+j,f'=SUM({c}5:{c}{rk})')
for r in range(5,rt+1):
    S3.cell(r,13,f'=IFERROR(L{r}/$L${rt},"")'); S3.cell(r,14,f'=IFERROR(J{r}/H{r}-1,"")')
    if r<rk:
        S3.cell(r,15,f'=IFERROR(SUMIFS({R("AE")},{R("G")},$A{r})/SUMIFS({R("M")},{R("G")},$A{r},{R("AE")},"<>"),"")')
        S3.cell(r,16,f'=IFERROR(SUMIFS({R("O")},{R("G")},$A{r})*60/SUMIFS({R("M")},{R("G")},$A{r},{R("O")},"<>"),"")')
    for c in range(2,13): S3.cell(r,c).number_format=NUM
    S3.cell(r,13).number_format=PCT; S3.cell(r,14).number_format=PCT; S3.cell(r,15).number_format=USD2; S3.cell(r,16).number_format='0'
body(S3,5,rt,1,16)
for c in range(1,17): S3.cell(rt,c).fill=TOT_FILL; S3.cell(rt,c).font=F(bold=True,size=10)
for r in range(5,rt+1): S3.cell(r,1).font=F(bold=True,size=10)
r=rt+2; S3.cell(r,1,'Chỉ số tập trung khách hàng').font=F(bold=True,size=11)
S3.cell(r+1,1,'Tỷ trọng TNF (gồm 別注)'); S3.cell(r+1,4,f'=L5/L{rt}'); S3.cell(r+1,4).number_format=PCT
S3.cell(r+2,1,'Tỷ trọng Top 3 brand'); S3.cell(r+2,4,f'=SUM(L5:L7)/L{rt}'); S3.cell(r+2,4).number_format=PCT
S3.cell(r+3,1,'Tỷ trọng TNF riêng mùa 27SS'); S3.cell(r+3,4,f'=J5/J{rt}'); S3.cell(r+3,4).number_format=PCT
S3.cell(r+4,1,'Tỷ trọng hàng 別注 (bespoke) 26FW'); S3.cell(r+4,4,f"='2. Mùa vụ'!N12"); S3.cell(r+4,4).font=GREEN; S3.cell(r+4,4).number_format=PCT
for k in range(1,5): S3.cell(r+k,1).font=BODY; S3.cell(r+k,4).font=F(bold=True,size=10) if k<4 else GREEN
widths(S3,[17]+[9]*10+[11,9,11,11,10]); S3.freeze_panes='B5'
ch3=BarChart(); ch3.type='bar'; ch3.title='Sản lượng tích luỹ 23SS–27FW theo brand (pcs)'; ch3.x_axis.title=None
ch3.add_data(Reference(S3,min_col=12,min_row=4,max_row=4+len(top)),titles_from_data=True); ch3.set_categories(Reference(S3,min_col=1,min_row=5,max_row=4+len(top)))
ch3.series[0].graphicalProperties.solidFill='2A78D6'; ch3.series[0].graphicalProperties.line.noFill=True; ch3.legend=None; ch3.x_axis.scaling.orientation='maxMin'; ch3.y_axis.numFmt='#,##0'; ch3.gapWidth=50
ch3.height=8; ch3.width=17; chart_fix(ch3); S3.add_chart(ch3,f'F{r}')

# ---------- 4. Công suất (ngày công) ----------
v_=pd.concat([o,O_DOT.assign(qty=0.0),O_SMS.assign(qty=0.0)],ignore_index=True)
v_=v_[(~v_.stale)&v_.start.notna()&v_.end.notna()]
v_=v_[(v_.end>=v_.start)&((v_.end-v_.start).dt.days<200)]
recs=[]
for r in v_.itertuples():
    days=pd.date_range(r.start,r.end); days=days[days.dayofweek<6]
    w=r.workers if pd.notna(r.workers) else 0
    for d in days: recs.append((r.factory,r.line,r.part,r.row,d,w,r.qty/len(days),(r.qty*(r.sam if pd.notna(r.sam) else 0))/3600/len(days)))
Ld=pd.DataFrame(recs,columns=['f','line','part','row','d','w','q','h']); Ld['m']=Ld.d.dt.to_period('M')
months=pd.period_range('2025-01','2027-06',freq='M')
_r2=[]
for x in Ld[Ld.f=='APIC'].itertuples():
    for hh in 'AB':
        if hw(x.part,hh)>0: _r2.append((x.line,hh,x.d,x.part if isinstance(x.part,str) else ''))
_O=pd.DataFrame(_r2,columns=['line','h','d','part'])
def _occ(ps):
    ps=set(ps)
    for p_ in ps:
        if p_=='' or p_=='SPL' or p_.startswith('AB') or p_ in ('A','B'): return 32
    return min(32,16*len([p_ for p_ in ps if p_[-1:].isdigit()]))
OC=_O.groupby(['line','h','d']).part.apply(_occ).rename('occ').reset_index(); OC['m']=OC.d.dt.to_period('M')
WD={m:sum(1 for d in pd.date_range(m.start_time,m.end_time) if d.dayofweek<6) for m in months}
title(S4,'Công suất xưởng APIC theo ngày công','Công-ngày có đơn: mỗi ngày làm việc (T2–T7), line thường (vd 1A) có đơn tính 32 CN, mỗi line nhỏ (vd 14A1) có đơn tính 16 CN; đơn AB chiếm cả 1A và 1B. Không cộng trùng khi nhiều đơn chạy song song trên cùng line. Công suất = định biên APIC (30 line × 32 CN) × ngày làm việc. AP1A/AP1B/AP1C cũ được gộp thành APIC. Đã loại các dòng "mùa cũ" còn treo lịch.')
h4=['Tháng','Ngày làm việc (T2–T7)','Công suất (công-ngày)','Công-ngày có đơn','Tỷ lệ sử dụng','Công-ngày trống','SL kế hoạch APIC (pcs)','Giờ chuẩn APIC','Hiệu suất yêu cầu (giờ chuẩn / giờ công)','Số mã hàng bắt đầu (đổi mã)','SL xưởng khác (pcs)','Tổng SL (pcs)','Trạng thái']
hdr(S4,4,h4,h=58)
for k,m in enumerate(months):
    r=5+k; a=Ld[(Ld.f=='APIC')&(Ld.m==m)]; b=Ld[(Ld.f!='APIC')&(Ld.m==m)]
    S4.cell(r,1,m.start_time.to_pydatetime()); S4.cell(r,1).number_format='mm/yyyy'
    S4.cell(r,2,WD[m]); S4.cell(r,3,f"=B{r}*'Phương pháp'!$C$5"); S4.cell(r,4,int(OC[OC.m==m].occ.sum()))
    S4.cell(r,5,f'=IFERROR(D{r}/C{r},"")'); S4.cell(r,6,f'=MAX(C{r}-D{r},0)')
    S4.cell(r,7,round(a.q.sum())); S4.cell(r,8,round(a.h.sum())); S4.cell(r,9,f'=IFERROR(H{r}/(D{r}*8),"")')
    S4.cell(r,10,f'=COUNTIFS({R("C")},"APIC",{R("P")},">="&A{r},{R("P")},"<"&EDATE(A{r},1),{R("AK")},"<>Y")')
    S4.cell(r,11,round(b.q.sum())); S4.cell(r,12,f'=G{r}+K{r}')
    S4.cell(r,13,f'=IF(E{r}>=0.85,"Đầy",IF(E{r}>=0.5,"Còn trống","Thiếu đơn"))')
    for c in (2,3,4,6,7,8,10,11,12): S4.cell(r,c).number_format=NUM
    S4.cell(r,5).number_format='0%'; S4.cell(r,9).number_format='0%'
r4=5+len(months)-1
body(S4,5,r4,1,13)
for txt,clr in [('Thiếu đơn','F6C9C9'),('Còn trống','FCE3B0'),('Vượt định biên','F9D3C2'),('Đầy','CFEFCF')]:
    S4.conditional_formatting.add(f'M5:M{r4}',FormulaRule(formula=[f'M5="{txt}"'],fill=PatternFill('solid',fgColor=clr)))
IX={m:5+k for k,m in enumerate(months)}
rr=r4+2; S4.cell(rr,1,'Tổng hợp công suất').font=F(bold=True,size=11)
P=pd.Period
rows4=[('Năng suất 2026 (pcs / công-ngày), T1–T10',f'=SUM(G{IX[P("2026-01")]}:G{IX[P("2026-10")]})/SUM(D{IX[P("2026-01")]}:D{IX[P("2026-10")]})','0.00'),
('Tỷ lệ sử dụng bình quân T1–T10/2026',f'=SUM(D{IX[P("2026-01")]}:D{IX[P("2026-10")]})/SUM(C{IX[P("2026-01")]}:C{IX[P("2026-10")]})','0%'),
('Tỷ lệ sử dụng bình quân T11/2026–T1/2027',f'=SUM(D{IX[P("2026-11")]}:D{IX[P("2027-01")]})/SUM(C{IX[P("2026-11")]}:C{IX[P("2027-01")]})','0%'),
('Tỷ lệ sử dụng T2/2027 (có Tết)',f'=E{IX[P("2027-02")]}','0%'),
('Tỷ lệ sử dụng bình quân T3–T6/2027',f'=SUM(D{IX[P("2027-03")]}:D{IX[P("2027-06")]})/SUM(C{IX[P("2027-03")]}:C{IX[P("2027-06")]})','0%'),
('Công-ngày cần bổ sung T3–T6/2027 để đạt mức tải mục tiêu',f"=MAX(SUM(C{IX[P('2027-03')]}:C{IX[P('2027-06')]})*'Phương pháp'!$C$10-SUM(D{IX[P('2027-03')]}:D{IX[P('2027-06')]}),0)",NUM),
('≈ Sản lượng cần bổ sung (pcs)',f'=F{rr+6}*F{rr+1}',NUM),
('≈ Chi phí nhân công chưa có đơn để phân bổ (USD)',f"=F{rr+6}*'Phương pháp'!$C$9",USD),
('Số đơn chạy song song bình quân / line / ngày (2026)',round(float(Ld[(Ld.f=='APIC')&(Ld.d>='2026-01-01')&(Ld.d<'2026-11-01')].groupby(['line','d']).row.nunique().mean()/2),2),'0.00')]
for k,(a,fm,nf) in enumerate(rows4,1):
    S4.cell(rr+k,1,a).font=BODY; S4.cell(rr+k,6,fm); S4.cell(rr+k,6).number_format=nf; S4.cell(rr+k,6).font=F(bold=True,size=10)
CAP={'prod':f"'4. Công suất'!F{rr+1}",'u26':f"'4. Công suất'!F{rr+2}",'u_now':f"'4. Công suất'!F{rr+3}",'u_feb':f"'4. Công suất'!F{rr+4}",'u_mj':f"'4. Công suất'!F{rr+5}",
     'wd_gap':f"'4. Công suất'!F{rr+6}",'pcs_gap':f"'4. Công suất'!F{rr+7}",'usd_gap':f"'4. Công suất'!F{rr+8}",'over':f"'4. Công suất'!F{rr+9}"}
S4.cell(rr+10,1,'Lưu ý: ﾗｲﾝ数 ở mỗi dòng là số CN của cả line, nên không cộng dồn khi nhiều đơn chạy song song. Tháng 2/2027 có Tết Nguyên đán (06/02/2027) nên ngày làm việc thực tế ít hơn.').font=F(italic=True,size=9,color='D03B3B')
widths(S4,[10,9,11,11,9,11,12,11,12,10,11,11,14]); S4.freeze_panes='B5'
lc=LineChart(); lc.title='APIC – tỷ lệ sử dụng ngày công theo tháng'; lc.y_axis.numFmt='0%'; lc.y_axis.title='Tỷ lệ sử dụng'; lc.y_axis.scaling.min=0; lc.y_axis.scaling.max=1.1
lc.add_data(Reference(S4,min_col=5,min_row=4,max_row=r4),titles_from_data=True); lc.set_categories(Reference(S4,min_col=1,min_row=5,max_row=r4))
lc.x_axis.number_format='mm/yy'; s=lc.series[0]; s.graphicalProperties.line.solidFill='2A78D6'; s.graphicalProperties.line.width=22000; s.smooth=False
lc.legend=None; lc.height=8; lc.width=20; chart_fix(lc); S4.add_chart(lc,'O4')
# numbers for text
def pv(m): return Ld[(Ld.f=='APIC')&(Ld.m==P(m))]
def occm(m): return OC[OC.m==P(m)].occ.sum()
def util(ms):
    return sum(occm(m) for m in ms)/sum(WD[P(m)]*960 for m in ms)
N_=dict(u26=util([f'2026-{i:02d}' for i in range(1,11)]),u_now=util(['2026-11','2026-12','2027-01']),u_feb=util(['2027-02']),u_mar=util(['2027-03']),u_mj=util(['2027-03','2027-04','2027-05','2027-06']))
_q=sum(pv(f'2026-{i:02d}').q.sum() for i in range(1,11)); _w=sum(occm(f'2026-{i:02d}') for i in range(1,11)); N_['prod']=_q/_w
N_['wd_gap']=max(sum(WD[P(m)]*960 for m in ['2027-03','2027-04','2027-05','2027-06'])*0.9-sum(occm(m) for m in ['2027-03','2027-04','2027-05','2027-06']),0)
N_['pcs_gap']=N_['wd_gap']*N_['prod']; N_['usd_gap']=N_['wd_gap']*46.29017951425554


# ---------- 5. Chi phí & Lô ----------
title(S5,'Chi phí gia công & quy mô lô','Đơn giá GC2 = giá vốn gia công theo lô SX thực tế. Lô càng nhỏ → chưa qua giai đoạn học việc → năng suất thấp, đơn giá cao.')
groups=['1. ≤300','2. 301–500','3. 501–1.000','4. 1.001–2.000','5. 2.001–5.000','6. >5.000']
hdr(S5,4,['Nhóm lô SX (pcs)','Số dòng đơn','% số dòng','Sản lượng (pcs)','% sản lượng','CP GC2 (USD)','Đơn giá GC2 (USD/pcs)','Hiệu suất lô TB (平均/MAX)','Đơn giá so với lô 2.001–5.000'],h=45)
for k,g in enumerate(groups):
    r=5+k; S5.cell(r,1,g)
    S5.cell(r,2,f'=COUNTIFS({R("AG")},$A{r})'); S5.cell(r,3,f'=B{r}/SUM($B$5:$B$10)')
    S5.cell(r,4,f'=SUMIFS({R("M")},{R("AG")},$A{r})'); S5.cell(r,5,f'=D{r}/SUM($D$5:$D$10)')
    S5.cell(r,6,f'=SUMIFS({R("AE")},{R("AG")},$A{r})'); S5.cell(r,7,f'=IFERROR(F{r}/D{r},"")')
    S5.cell(r,8,f'=IFERROR(AVERAGEIFS({R("AH")},{R("AG")},$A{r}),"")'); S5.cell(r,9,f'=IFERROR(G{r}/$G$9,"")')
    for c,fm in [(2,NUM),(3,PCT),(4,NUM),(5,PCT),(6,USD),(7,USD2),(8,'0%'),(9,'0.0"x"')]: S5.cell(r,c).number_format=fm
    S5.cell(r,1).font=F(bold=True,size=10)
body(S5,5,10,1,9)
S5['A12']='Xu hướng lô nhỏ theo mùa'; S5['A12'].font=F(bold=True,size=11)
hdr(S5,13,['Mùa','Số dòng lô ≤300','% dòng lô ≤300','Chi phí chia lô (USD)','% chi phí chia lô'],h=32)
for k,sn in enumerate(ORDER):
    r=14+k; S5.cell(r,1,sn)
    S5.cell(r,2,f'=COUNTIFS({R("B")},$A{r},{R("Z")},"<="&\'Phương pháp\'!$C$8)'); S5.cell(r,3,f"=IFERROR(B{r}/'2. Mùa vụ'!B{5+k},\"\")")
    S5.cell(r,4,f"='2. Mùa vụ'!J{5+k}"); S5.cell(r,5,f"='2. Mùa vụ'!K{5+k}")
    S5.cell(r,2).number_format=NUM; S5.cell(r,3).number_format=PCT; S5.cell(r,4).number_format=USD; S5.cell(r,5).number_format=PCT
    for c in (3,4,5): S5.cell(r,c).font=GREEN
body(S5,14,23,1,5)
widths(S5,[18,11,9,13,10,13,11,12,12])
bc=BarChart(); bc.type='col'; bc.title='Đơn giá gia công theo quy mô lô (USD/pcs)'; bc.y_axis.numFmt='$#,##0'; bc.y_axis.title='USD/pcs'
bc.add_data(Reference(S5,min_col=7,min_row=4,max_row=10),titles_from_data=True); bc.set_categories(Reference(S5,min_col=1,min_row=5,max_row=10))
bc.series[0].graphicalProperties.solidFill='2A78D6'; bc.series[0].graphicalProperties.line.noFill=True; bc.legend=None; bc.gapWidth=60
bc.dataLabels=DataLabelList(); bc.dataLabels.showVal=True; bc.dataLabels.showCatName=False; bc.dataLabels.showSerName=False; bc.dataLabels.showLegendKey=False; bc.dataLabels.showPercent=False; bc.dataLabels.numFmt='$#,##0.0'
bc.height=7.5; bc.width=16; chart_fix(bc); S5.add_chart(bc,'K4')

# ---------- 6. Rủi ro giao hàng ----------
fw=o[(o.end>=REF)&(~o.stale)].copy()
fw['mat_late']=fw.input_ok.notna()&fw.start.notna()&(fw.input_ok>fw.start)
fw['mat_unk']=fw.input_txt.fillna('').str.contains('未定')
fw['smp_ok']=fw.smp_reply.map(lambda x: isinstance(x,(dt.datetime,pd.Timestamp)) or (isinstance(x,str) and x.strip() in ('●','継続','並行進行でOK','パターンの通り進行OK')))
def sev(g): return 0 if pd.isna(g) else (3 if g<=-30 else 2 if g<=-8 else 1 if g<0 else 0)
fw['score']=fw.gap.map(sev)*10+fw.mat_late*5+fw.mat_unk*5+(~fw.smp_ok)*2+fw.fac_etd.isna()*3
fw=fw.sort_values(['score','start'],ascending=[False,True])
title(S6,'Rủi ro giao hàng – đơn đang/ sắp sản xuất','Đơn có END ≥ ngày tham chiếu. Lệch ETD = 工場ETD yêu cầu − ETD kế hoạch (âm = trễ). Ngưỡng ở sheet Phương pháp.')
S6['A4']='Tổng quan'; S6['A4'].font=F(bold=True,size=11)
start_r=14
lst=f'$A${start_r+1}:$A${start_r+len(fw)}'
def rng(c): return f'${c}${start_r+1}:${c}${start_r+len(fw)}'
kp=[('Số đơn đang/ sắp SX',f'=COUNTA({rng("A")})',f'=SUM({rng("I")})'),
    ('Mức "Nghiêm trọng" (trễ ≥30 ngày)',f'=COUNTIFS({rng("O")},"Nghiêm trọng")',f'=SUMIFS({rng("I")},{rng("O")},"Nghiêm trọng")'),
    ('Mức "Cao" (trễ 8–29 ngày)',f'=COUNTIFS({rng("O")},"Cao")',f'=SUMIFS({rng("I")},{rng("O")},"Cao")'),
    ('Mức "Theo dõi" (trễ 1–7 ngày)',f'=COUNTIFS({rng("O")},"Theo dõi")',f'=SUMIFS({rng("I")},{rng("O")},"Theo dõi")'),
    ('Thiếu 工場ETD',f'=COUNTIFS({rng("O")},"Thiếu 工場ETD")',f'=SUMIFS({rng("I")},{rng("O")},"Thiếu 工場ETD")'),
    ('NVL về sau ngày START',f'=COUNTIFS({rng("P")},"Trễ")',f'=SUMIFS({rng("I")},{rng("P")},"Trễ")'),
    ('NVL chưa xác định (未定)',f'=COUNTIFS({rng("P")},"Chưa xác định")',f'=SUMIFS({rng("I")},{rng("P")},"Chưa xác định")'),
    ('Mẫu chưa có phản hồi/duyệt',f'=COUNTIFS({rng("Q")},"Chưa duyệt")',f'=SUMIFS({rng("I")},{rng("Q")},"Chưa duyệt")')]
hdr(S6,5,['Chỉ số','','','Số đơn','Sản lượng (pcs)'],h=20)
for k,(a,b,c) in enumerate(kp):
    r=6+k; S6.cell(r,1,a); S6.cell(r,4,b); S6.cell(r,5,c); S6.cell(r,4).number_format=NUM; S6.cell(r,5).number_format=NUM
body(S6,6,13,1,5)
h6=['Dòng gốc','Mùa','Xưởng','Line','Brand','Style','Item','Giai đoạn mẫu','QTY','START','END','ETD KH','工場ETD','Lệch ETD (ngày)','Mức ETD','NVL','Mẫu','Ngày có thể đưa vào SX','Ghi chú gốc (rút gọn)','Giá trị gia công (USD)']
hdr(S6,start_r,h6,h=36)
for i,r in enumerate(fw.itertuples(),start_r+1):
    vals=[r.row,r.season,r.factory,v(r.line),r.brand,r.style,r.item,r.smp_stage,v(r.qty),v(r.start),v(r.end),v(r.etd_plan),v(r.fac_etd),v(r.gap),
          f"=IF(M{i}=\"\",\"Thiếu 工場ETD\",IF(N{i}=\"\",\"Kiểm tra\",IF(N{i}<='Phương pháp'!$C$6,\"Nghiêm trọng\",IF(N{i}<='Phương pháp'!$C$7,\"Cao\",IF(N{i}<0,\"Theo dõi\",\"Đúng hạn\")))))",
          'Chưa xác định' if r.mat_unk else f'=IF(AND(ISNUMBER(R{i}),ISNUMBER(J{i})),IF(R{i}>J{i},"Trễ","OK"),"")',
          'OK' if r.smp_ok else 'Chưa duyệt',
          v(r.input_ok) if pd.notna(r.input_ok) else r.input_txt,(r.comment if isinstance(r.comment,str) else '')[:120],v(r.qty*r.c2_upc) if pd.notna(r.c2_upc) and pd.notna(r.qty) else None]
    for j,x in enumerate(vals,1): S6.cell(i,j,x).font=BODY
    for j in (10,11,12,13,18): S6.cell(i,j).number_format=DATE
    S6.cell(i,9).number_format=NUM; S6.cell(i,20).number_format=USD
end6=start_r+len(fw)
for col,txt,clr in [('O','Nghiêm trọng','F6C9C9'),('O','Cao','F9D3C2'),('O','Theo dõi','FCE3B0'),('O','Thiếu 工場ETD','E0E0E0'),('P','Trễ','F9D3C2'),('P','Chưa xác định','FCE3B0'),('Q','Chưa duyệt','FCE3B0')]:
    S6.conditional_formatting.add(f'{col}{start_r+1}:{col}{end6}',FormulaRule(formula=[f'{col}{start_r+1}="{txt}"'],fill=PatternFill('solid',fgColor=clr)))
S6.freeze_panes=f'A{start_r+1}'; S6.auto_filter.ref=f'A{start_r}:T{end6}'
widths(S6,[8,6,6,5,13,14,24,8,8,10,10,10,10,8,12,12,10,11,50,12])
S6.column_dimensions['A'].width=30

# ---------- 7. Chất lượng DL ----------
title(S7,'Chất lượng dữ liệu – các điểm cần làm sạch tại file gốc','Ảnh hưởng trực tiếp đến độ tin cậy của báo cáo. Đề xuất giao người phụ trách chuẩn hoá.')
raw=pd.read_pickle(os.path.join(WORKDIR,'raw.pkl')); rawd=raw[raw.Q.notna()|raw.I.notna()]
errc=int(sum(1 for c in rawd.columns for x in rawd[c] if isinstance(x,str) and x.startswith('#')))
seas_raw=rawd.A.dropna().astype(str).str.strip().value_counts()
stg=o.smp_stage.dropna().nunique()
issues=[('Brand trống',f'{(o.brand=="(trống)").sum()} dòng / {int(o[o.brand=="(trống)"].qty.sum()):,} pcs','Cao','Không phân bổ được doanh thu/chi phí theo khách hàng (gồm đơn HH26-103, mẫu SMS).','Bắt buộc nhập Brand; dùng Data Validation.'),
('Mã mùa vụ không thống nhất',f'{len(seas_raw)} cách ghi cho 10 mùa (vd "26FW\\"" và "26FW")','Trung bình','Pivot/lọc trong file gốc bị tách đôi một mùa.','Danh sách thả xuống cố định cho cột Status.'),
('Brand ghi lẫn ghi chú',f'{o.brand_raw.nunique()} giá trị Brand gốc (vd "TNF↵別注", "HH↵別注")','Trung bình','Khó tổng hợp; 別注 nên là cột riêng.','Tách cột "別注 (Y/N)".'),
("Dòng chấm (chuyền lớn ghi 2 lần)",f'{len(O_DOT)} dòng / {int(O_DOT.qty_raw.sum()):,} pcs có ô QTY nền chấm (chủ yếu dòng thứ 2 của chuyền lớn AB)','Cao','Nếu cộng sẽ gấp đôi SL. Công thức gốc SUMIF theo style (標準原価1) vẫn cộng cả dòng chấm nên lô bị phóng đại.',f'Báo cáo: không cộng SL dòng chấm, chỉ dùng để tính chiếm line; tính lại lô 原価1 cho {N_ADJ1} dòng. File gốc: đặt QTY dòng thứ 2 = 0 hoặc loại khỏi SUMIF.'),
('Mã xưởng cũ AP1A/AP1B/AP1C','Đã gộp thành APIC trong báo cáo (cột "Mã xưởng cũ" ở sheet Data)','Thấp','Dễ hiểu nhầm là 2 xưởng khác nhau khi tổng hợp công suất.','Đổi cột Factory trong file gốc thành APIC.'),
('Ký tự full-width','Factory "AP２", giai đoạn mẫu "ＰＰＳ", "F23 " có khoảng trắng','Thấp','Lọc/đếm sai.','Chuẩn hoá danh mục.'),
('Ô lỗi công thức',f'{errc} ô lỗi (VALUE, N/A…)','Trung bình','Tổng/trung bình trong file gốc có thể sai.','Bọc IFERROR / sửa nguồn.'),
('Dòng mùa cũ còn treo lịch',f'{int(o.stale.sum())} dòng 23SS ở line 7, lịch đến 02/2027 (dòng gốc {", ".join(map(str,o[o.stale].row.tolist()))}); đã loại khỏi phân tích','Cao','START = END dòng trước + 1 nên dòng thừa cuối khối line bị đẩy vào tương lai → làm sai tải line.','Xoá hoặc đánh dấu huỷ các dòng không còn sản xuất.'),
('Mã mùa không khớp lịch SX',f'{int(o.chk_season.sum())} dòng 26SS/26FW có START muộn (vd 26FW chạy 02–03/2027; dòng gốc {", ".join(map(str,o[o.chk_season].row.tolist()))})','Trung bình','Có thể là đơn bổ sung hoặc ghi sai mùa; vẫn được tính trong báo cáo.','Xác nhận mã mùa đúng với Sales.'),
('Lệch ETD (cột S) vô nghĩa',f'{int(o.gap_bad.sum())} dòng có giá trị ±40.000+ ngày','Trung bình','Do 工場ETD trống nhưng công thức vẫn trừ.','Sửa công thức S: trả "" khi U trống.'),
('Thiếu 工場ETD (đơn tương lai)',f'{int(((o.end>=REF)&(~o.stale)&o.fac_etd.isna()).sum())} đơn đang/sắp SX','Cao','Không kiểm soát được hạn giao.','Sales cập nhật 工場ETD trước khi xếp line.'),
('Cột BL không có tiêu đề',f'{int(o.BL.notna().sum())} dòng có số liệu','Thấp','Không rõ ý nghĩa → không dùng.','Đặt tên cột hoặc xoá.'),
('Cột BO/BP (END+38, chênh lệch)','Chỉ 234 dòng (2024–25) có dữ liệu','Thấp','Không theo dõi được kho nhận hàng cho mùa mới.','Bỏ hoặc kéo công thức cho toàn bộ.'),
('Giai đoạn mẫu nhiều mã',f'{stg} mã khác nhau (PPS, SMS, 1st, 暫定, 想定, F25, S26…)','Thấp','Khó đo tiến độ mẫu.','Chuẩn hoá: 1st → SMS → PPS → BULK.'),
('Cột mẫu gửi/phản hồi trộn kiểu','Ngày, "●", "継続", ghi chú tự do','Trung bình','Không đo được lead time duyệt mẫu.','Tách cột ngày & cột trạng thái.'),
('Đơn giá GC bất thường',f'{int((o.c2_upc>200).sum())} dòng > 200 USD/pcs (SL rất nhỏ / mẫu)','Thấp','Làm lệch trung bình nếu không gia quyền theo SL.','Tách đơn mẫu khỏi đơn sản xuất.')]
hdr(S7,4,['Vấn đề','Quy mô','Mức ảnh hưởng','Tác động','Đề xuất xử lý'])
for k,row in enumerate(issues):
    for j,x in enumerate(row,1):
        c=S7.cell(5+k,j,x); c.alignment=Alignment(wrap_text=True,vertical='top')
body(S7,5,4+len(issues),1,5)
widths(S7,[28,40,12,50,40])



# ---------- 1. Tóm tắt điều hành ----------
RK=dict(n=len(fw),q=fw.qty.sum()/1000,n8=int((fw.gap<=-8).sum()),q8=fw[fw.gap<=-8].qty.sum()/1000,p8=fw[fw.gap<=-8].qty.sum()/fw.qty.sum(),n30=int((fw.gap<=-30).sum()),
        nml=int(fw.mat_late.sum()),qml=fw[fw.mat_late].qty.sum()/1000,nmu=int(fw.mat_unk.sum()),nsm=int((~fw.smp_ok).sum()),nfe=int(fw.fac_etd.isna().sum()))
title(S1,'BÁO CÁO ĐIỀU HÀNH – KẾ HOẠCH SẢN XUẤT (PLANTABLE)','Nguồn: APIC P-Plan 2026 (bản 2026-11-13), sheet "PLANTABLE (原価) 2026年" – 3.494 dòng đơn, 23SS → 27FW. Ngày tham chiếu: 14/11/2026.')
S1.merge_cells('A1:F1')
S1['A4']='A. CHỈ SỐ CHÍNH'; S1['A4'].font=F(bold=True,size=12,color='1C5CAB')
hdr(S1,5,['Chỉ số','Giá trị','So sánh / ghi chú','','Đọc nhanh'],h=22)
S1.merge_cells('C5:D5')
y26=yr+5
k=[('Sản lượng năm 2026 (26SS + 26FW)',f"='2. Mùa vụ'!B{y26}",NUM,f"='2. Mùa vụ'!C{y26}",PCT,'so với 2025'),
   ('Sản lượng mùa 27SS',"='2. Mùa vụ'!D13",NUM,"='2. Mùa vụ'!E13",PCT,'so với 26SS'),
   ('Tỷ trọng TNF trong tổng sản lượng',f"='3. Khách hàng'!D{rt+3}",PCT,f"='3. Khách hàng'!D{rt+5}",PCT,'riêng mùa 27SS'),
   ('Tỷ lệ sử dụng ngày công APIC T11/2026 – T1/2027',"="+CAP['u_now'],'0%',"="+CAP['u_feb'],'0%','riêng T2/2027 (Tết)'),
   ('SL cần bổ sung T3–T6/2027 để đạt tải 90% (pcs)',"="+CAP['pcs_gap'],NUM,"="+CAP['usd_gap'],USD,'chi phí nhân công chưa có đơn'),
   ('Chi phí chia lô tích luỹ 23SS–27FW (USD)',f"='2. Mùa vụ'!J{tr}",USD,f"='2. Mùa vụ'!K13",PCT,'tỷ lệ mùa 27SS (đang giảm)'),
   ('Đơn giá GC lô ≤300 pcs so với lô 2.001–5.000',"='5. Chi phí & Lô'!I5",'0.0"x"',"='5. Chi phí & Lô'!C5",PCT,'tỷ trọng số dòng đơn lô ≤300'),
   ('Đơn sắp SX trễ ETD ≥ 8 ngày (pcs)',"='6. Rủi ro giao hàng'!E7+'6. Rủi ro giao hàng'!E8",NUM,"=('6. Rủi ro giao hàng'!E7+'6. Rủi ro giao hàng'!E8)/'6. Rủi ro giao hàng'!E6",PCT,'trên tổng SL đang/sắp SX'),
   ('Đơn có NVL về sau ngày START (pcs)',"='6. Rủi ro giao hàng'!E11",NUM,"='6. Rủi ro giao hàng'!D11",NUM,'số đơn')]
for i,(a,b,bf,c,cf,d) in enumerate(k,6):
    S1.cell(i,1,a).font=F(size=10,bold=True); S1.cell(i,2,b); S1.cell(i,2).number_format=bf; S1.cell(i,2).font=F(size=11,bold=True,color='008000')
    S1.cell(i,3,c); S1.cell(i,3).number_format=cf; S1.cell(i,3).font=GREEN; S1.cell(i,4,d).font=F(size=9,italic=True,color='52514E')
reads=['Ổn định, tăng nhẹ','Giảm – cần bù','Phụ thuộc cao','Kín đến T1/2027','Lỗ hổng lớn','Cải thiện',f'Lô nhỏ đắt gấp ~{V(K["lot_x"])} lần','Rủi ro cao','Cần xử lý NVL']
st=['good','warning','serious','warning','critical','good','serious','critical','serious']
FILL={'good':'CFEFCF','warning':'FCE3B0','serious':'F9D3C2','critical':'F6C9C9'}
ICON={'good':'✔ ','warning':'▲ ','serious':'▲ ','critical':'● '}
for i,(t,s_) in enumerate(zip(reads,st),6):
    c=S1.cell(i,5,ICON[s_]+t); c.fill=PatternFill('solid',fgColor=FILL[s_]); c.font=F(size=10,bold=True)
body(S1,6,14,1,5)

S1['A17']='B. PHÁT HIỆN CHÍNH'; S1['A17'].font=F(bold=True,size=12,color='1C5CAB')
finds=[
('1. Quy mô đi ngang, mùa 27SS giảm.',f'Sản lượng năm 2024 đạt {K["y24"]/1000:.0f}k pcs, 2025 là {K["y25"]/1000:.0f}k pcs ({K["g25"]:+.0%}), 2026 là {K["y26"]/1000:.0f}k pcs ({K["g26"]:+.0%}). Mùa 27SS hiện có {K["s27"]/1000:.0f}k pcs ({K["g27"]:+.0%} so với 26SS). Tổng giờ chuẩn mỗi năm khoảng {V(K["h_lo"])}–{V(K["h_hi"])} triệu giờ vì hàng FW phức tạp hơn nhiều (SAM ~{K["samFW_lo"]:.0f}–{K["samFW_hi"]:.0f} phút/sp so với ~{K["samSS_lo"]:.0f}–{K["samSS_hi"]:.0f} phút/sp của hàng SS). Vì vậy nên đo năng lực bằng giờ chuẩn thay vì số pcs. Số liệu đã bỏ {K["ndot"]} dòng chấm ({K["qdot"]:.0f}k pcs) để không cộng 2 lần đơn chuyền lớn, và tách riêng {K["nsms"]} dòng hàng SMS ({K["qsms_vn"]} pcs) hỗ trợ chuyền mẫu.'),
('2. Phụ thuộc lớn vào một khách hàng.',f'TNF chiếm {K["tnf"]:.0%} tổng sản lượng và {K["tnf27"]:.0%} mùa 27SS. Top 3 brand chiếm {K["top3"]:.0%}. MONTBELL giảm mạnh: {K["mb"][0]:.0f}k pcs (24FW) → {K["mb"][1]:.0f}k (25FW) → {K["mb"][2]:.0f}k (26FW) → {K["mb"][3]:.0f}k ở 27SS. LL.BEAN đã dừng từ 25SS. Tăng trưởng 27SS so với cùng kỳ 26SS: HELLY HANSEN {K["hh"]:+.0%}, ALLTERRAIN 81 {K["at"]:+.0%}; ngược lại GOLDWIN {K["gw"]:+.0%}. Hàng 別注 tăng từ 0% lên {K["bes26"]:.0%} sản lượng (26FW).'),
('3. Xưởng APIC kín lịch đến T1/2027, sau đó là khoảng trống lớn.',f'Kế hoạch dùng khoảng {N_["u_now"]:.0%} ngày công từ T11/2026 đến T1/2027 (bình quân 2026 là {N_["u26"]:.0%}). Tháng 2/2027 có Tết nên còn {N_["u_feb"]:.0%}. Từ T3/2027 chỉ còn {N_["u_mar"]:.0%}, và T4–T6 gần như chưa có đơn. Để đạt mức tải 90% trong T3–T6/2027 cần thêm khoảng {N_["pcs_gap"]/1000:,.0f}k pcs. Nếu không lấp được, khoảng {VN1(N_["usd_gap"]/1e6)} triệu USD chi phí nhân công sẽ không có đơn để phân bổ. Hiện 27FW mới có {K["s27fw"]/1000:.0f}k pcs, so với {K["s26fw"]/1000:.0f}k pcs của 26FW.'),
('4. Chia nhỏ lô làm tăng chi phí.',f'Lô ≤300 pcs chiếm {K["lot_rows"]:.0%} số dòng đơn nhưng chỉ {K["lot_qty"]:.0%} sản lượng. Đơn giá gia công của lô này là {V(K["upc_s"])} USD/pcs, gấp {V(K["lot_x"])} lần lô 2.001–5.000 ({V(K["upc_m"])} USD/pcs), vì chưa qua đường cong học việc (chỉ đạt ~{K["eff_s"]:.0%} tốc độ tối đa). Phần "chi phí chia lô" (giá vốn theo lô sản xuất thực tế so với lô đặt hàng) cộng dồn khoảng {V(K["split"])} triệu USD qua 10 mùa. Tỷ lệ này là {K["sp25"]:.0%} ở 25SS, {K["sp26fw"]:.0%} ở 26FW và {K["sp27"]:.0%} ở 27SS. Riêng 26FW tỷ lệ lô nhỏ lại tăng lên {K["small26fw"]:.0%} số dòng.'),
('5. Rủi ro giao hàng tập trung ở 27SS.',f'Có {RK["n"]} đơn ({RK["q"]:.0f}k pcs) đang hoặc sắp sản xuất. Trong đó {RK["n8"]} đơn ({RK["q8"]:.0f}k pcs, ~{RK["p8"]:.0%}) có ETD kế hoạch trễ hơn 工場ETD từ 8 ngày trở lên, và {RK["n30"]} đơn trễ từ 30 ngày trở lên. {RK["nml"]} đơn ({RK["qml"]:.0f}k pcs) có NVL về sau ngày bắt đầu may. {RK["nsm"]} đơn chưa có phản hồi hoặc duyệt mẫu. {RK["nfe"]} đơn chưa có 工場ETD.'+' '+f'Tỷ lệ đơn có ETD kế hoạch trễ của 27SS là {K["late27"]:.0%}, so với {K["late_lo"]:.0%}–{K["late_hi"]:.0%} ở các mùa 24FW–26FW.'),
('6. Dữ liệu chưa đủ tin cậy để ra quyết định nhanh.',f'{K["blank"]:.0f}k pcs không có brand. Đơn chuyền lớn (AB) ghi 2 lần, dòng thứ 2 tô chấm: {K["ndot"]} dòng / {K["qdot"]:.0f}k pcs, dễ bị cộng gấp đôi.'+' Mã mùa vụ ghi không thống nhất. File gốc có 139 ô lỗi công thức. 9 dòng 23SS cũ vẫn còn treo lịch đến 02/2027, làm sai tải line 7 (đã loại khỏi phân tích). Cột lệch ETD cho giá trị vô nghĩa khi 工場ETD trống. Chi tiết ở sheet 7.'),
('7. 受注進行表 (tự dựng từ PLANTABLE): 3 tháng tới kín nhưng phần lớn tạm tính; các tháng sau còn trống.',f'{K["j_w1"]} phủ {K["j_t1"]:.0%} công suất nhưng chỉ {K["j_c1"]:.0%} là đơn đã chốt (đã chạy hoặc NVL OK). {K["j_w2"]} kể cả tạm tính chỉ phủ {K["j_t2"]:.0%} (đỉnh {K["j_pk"]:.0%} ngày {K["j_pkd"]:%d/%m}). Chi tiết và bảng 受注進行表 ở sheet 4A.')]
r=18
for a,b in finds:
    S1.cell(r,1,a).font=F(bold=True,size=10); S1.cell(r,2,b).font=BODY; S1.merge_cells(start_row=r,start_column=2,end_row=r,end_column=5)
    S1.cell(r,2).alignment=Alignment(wrap_text=True,vertical='top'); S1.cell(r,1).alignment=Alignment(wrap_text=True,vertical='top')
    S1.row_dimensions[r].height=max(45,14*(len(b)//78+1)); r+=1
r+=1; S1.cell(r,1,'C. KHUYẾN NGHỊ HÀNH ĐỘNG').font=F(bold=True,size=12,color='1C5CAB'); r+=1
hdr(S1,r,['Hành động','Nội dung','Phụ trách (đề xuất)','Thời hạn','Tác động kỳ vọng'],h=22); r+=1
acts=[('Lấp công suất T3–T6/2027','Chốt đơn 27FW sớm, ưu tiên HELLY HANSEN, GOLDWIN và brand mới. Cân nhắc đưa đơn 27SS đang trễ sang T3 để giảm tải T1–T2.','Sales / Merchandising','Trước 15/12/2026',f'Tránh khoảng {N_["pcs_gap"]/1000:,.0f}k pcs công suất bị bỏ trống (~{VN1(N_["usd_gap"]/1e6)} triệu USD chi phí nhân công)'),
(f'Xử lý {RK["n8"]} đơn trễ ETD ≥ 8 ngày','Họp xử lý theo danh sách ở sheet 6, ưu tiên mức "Nghiêm trọng". Các lựa chọn: tăng line hoặc tăng ca, tách lô, hoặc đàm phán lùi 工場ETD.','Kế hoạch SX + Sales','Tuần 47–48/2026',f'Bảo vệ ~{RK["q8"]:.0f}k pcs giao đúng hạn'),
('Đẩy NVL & duyệt mẫu',f'Có {RK["nml"]} đơn NVL về sau ngày START, {RK["nmu"]} đơn NVL 未定 và {RK["nsm"]} đơn chưa duyệt mẫu.'+' Lập checklist "sẵn sàng sản xuất" 4 tuần trước START.','Mua hàng / QA mẫu','Hàng tuần','Giảm dừng line, đổi lịch'),
('Chính sách lô tối thiểu','Đặt MOQ hoặc phụ phí cho lô ≤300 pcs. Gộp lô cùng style hoặc cùng 集約CODE. Theo dõi % chi phí chia lô theo mùa (mục tiêu ≤10%).','Ban điều hành + Sales','Áp dụng từ 27FW',f'Giảm đơn giá GC lô nhỏ (đang gấp {V(K["lot_x"])} lần)'),
('Giảm rủi ro tập trung khách hàng','Mục tiêu brand ngoài TNF đạt ≥ 30% sản lượng. Tìm hiểu lý do MONTBELL giảm và có kế hoạch giữ khách.','Ban điều hành','Kế hoạch 2027','Ổn định doanh thu'),
('Chuẩn hoá dữ liệu PLANTABLE','Dùng danh mục thả xuống cho mùa, brand, factory. Tách cột 別注. Sửa công thức cột S. Xoá các dòng treo. Bắt buộc nhập 工場ETD.','Kế hoạch SX / IT','30/11/2026','Báo cáo tự động, tin cậy')]
for a in acts:
    for j,x in enumerate(a,1):
        c=S1.cell(r,j,x); c.alignment=Alignment(wrap_text=True,vertical='top'); c.font=F(size=10,bold=(j==1))
    S1.row_dimensions[r].height=48; r+=1
body(S1,r-len(acts),r-1,1,5)
r+=1; S1.cell(r,1,'Lưu ý: rủi ro giao hàng đánh giá theo 工場ETD trong file PLANTABLE (Lệch ETD = 工場ETD − ETD kế hoạch). Tỷ lệ sử dụng chưa trừ ngày lễ.').font=F(italic=True,size=9,color='52514E')
widths(S1,[38,16,16,24,26]); S1.sheet_view.showGridLines=False
S1.column_dimensions['B'].width=18; S1.column_dimensions['C'].width=14; S1.column_dimensions['D'].width=30; S1.column_dimensions['E'].width=40


# ================= ROLE SHEETS =================
SC=wb.create_sheet('1A. CEO'); SG=wb.create_sheet('1B. GĐ Nhà máy'); SQ=wb.create_sheet('1C. QL Sản xuất'); SX=wb.create_sheet('0. Mục lục')
NOTE=lambda ws,r,txt: (ws.cell(r,1,txt), setattr(ws.cell(r,1),'font',F(italic=True,size=9,color='52514E')))
def section(ws,r,txt): ws.cell(r,1,txt).font=F(bold=True,size=12,color='1C5CAB')
STF={'good':('✔ ','CFEFCF'),'warning':('▲ ','FCE3B0'),'serious':('▲ ','F9D3C2'),'critical':('● ','F6C9C9')}
def kpi_table(ws,r0,items,note=(4,4),stat=(5,5)):
    heads=[(1,1,'Chỉ số'),(2,2,'Giá trị'),(3,3,'So sánh'),(note[0],note[1],'Ghi chú'),(stat[0],stat[1],'Đánh giá')]
    for a,b,h_ in heads:
        for c in range(a,b+1):
            x=ws.cell(r0,c); x.fill=HDR_FILL; x.border=BOX
        x=ws.cell(r0,a,h_); x.font=HDR_FONT; x.alignment=Alignment(horizontal='center',vertical='center')
        if b>a: ws.merge_cells(start_row=r0,start_column=a,end_row=r0,end_column=b)
    ws.row_dimensions[r0].height=22
    for i,(a,b,bf,c,cf,d,lab,st) in enumerate(items,r0+1):
        ws.cell(i,1,a).font=F(size=10,bold=True); ws.cell(i,2,b).number_format=bf; ws.cell(i,2).font=F(size=11,bold=True,color='008000')
        if c is not None: ws.cell(i,3,c).number_format=cf; ws.cell(i,3).font=GREEN
        ws.cell(i,note[0],d).font=F(size=9,italic=True,color='52514E')
        ic,fl=STF[st]; x=ws.cell(i,stat[0],ic+lab); x.font=F(size=10,bold=True)
        for cc_ in range(stat[0],stat[1]+1): ws.cell(i,cc_).fill=PatternFill('solid',fgColor=fl)
        for cc_ in range(1,stat[1]+1): ws.cell(i,cc_).border=BOX
        if note[1]>note[0]: ws.merge_cells(start_row=i,start_column=note[0],end_row=i,end_column=note[1])
        if stat[1]>stat[0]: ws.merge_cells(start_row=i,start_column=stat[0],end_row=i,end_column=stat[1])
    return r0+len(items)
def text_rows(ws,r,rows,ncol=5,w=95):
    for a,b in rows:
        ws.cell(r,1,a).font=F(bold=True,size=10); ws.cell(r,1).alignment=Alignment(wrap_text=True,vertical='top')
        ws.cell(r,2,b).font=BODY; ws.cell(r,2).alignment=Alignment(wrap_text=True,vertical='top'); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=ncol)
        ws.row_dimensions[r].height=max(30,14*(len(b)//w+1)); r+=1
    return r
def table_rows(ws,r,head,rows,h=22,heights=40,spans=None):
    spans=spans or [(j,j) for j in range(1,len(head)+1)]
    for (a,b),h_ in zip(spans,head):
        for c in range(a,b+1): ws.cell(r,c).fill=HDR_FILL; ws.cell(r,c).border=BOX
        x=ws.cell(r,a,h_); x.font=HDR_FONT; x.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
        if b>a: ws.merge_cells(start_row=r,start_column=a,end_row=r,end_column=b)
    ws.row_dimensions[r].height=h; r+=1
    for row in rows:
        for j,((a,b),x) in enumerate(zip(spans,row)):
            c=ws.cell(r,a,x); c.alignment=Alignment(wrap_text=True,vertical='top'); c.font=F(size=10,bold=(j==0))
            for cc_ in range(a,b+1): ws.cell(r,cc_).border=BOX
            if b>a: ws.merge_cells(start_row=r,start_column=a,end_row=r,end_column=b)
        ws.row_dimensions[r].height=heights; r+=1
    return r
S6R=lambda c: f"'6. Rủi ro giao hàng'!${c}${start_r+1}:${c}${end6}"
sev_hi=fw[fw.gap<=-8]; risk_usd=float((sev_hi.qty*sev_hi.c2_upc).sum())

# ---------- 1A. CEO ----------
title(SC,'TRANG CEO – BỨC TRANH KINH DOANH & QUYẾT ĐỊNH CẦN ĐƯA RA','Đọc trong 3 phút. Số liệu liên kết trực tiếp với các sheet phân tích; ngày tham chiếu 14/11/2026.')
section(SC,4,'Bức tranh 30 giây')
SC.cell(5,1,f'Sản lượng đi ngang quanh {V(min(K["y24"],K["y25"],K["y26"])/1e6,2)}–{V(max(K["y24"],K["y25"],K["y26"])/1e6,2)} triệu pcs/năm; mùa 27SS {K["g27"]:+.0%} so với 26SS. Nhà máy kín lịch đến T1/2027, nhưng từ T3/2027 gần như chưa có đơn: cần thêm khoảng {N_["pcs_gap"]/1000:,.0f}k pcs, nếu không ~{VN1(N_["usd_gap"]/1e6)} triệu USD chi phí nhân công không có đơn để phân bổ. TNF chiếm {K["tnf"]:.0%} sản lượng. Lô nhỏ đắt gấp {V(K["lot_x"])} lần. Khoảng {RK["p8"]:.0%} sản lượng sắp sản xuất có ETD kế hoạch trễ từ 8 ngày trở lên.')
SC.merge_cells('A5:E5'); SC['A5'].alignment=Alignment(wrap_text=True,vertical='top'); SC['A5'].font=F(size=11); SC.row_dimensions[5].height=62
section(SC,7,'Chỉ số sức khoẻ kinh doanh')
y26=yr+5
ce=[('Sản lượng năm 2026 (pcs)',f"='2. Mùa vụ'!B{y26}",NUM,f"='2. Mùa vụ'!C{y26}",PCT,'so với 2025','Ổn định','good'),
('Sản lượng mùa 27SS (pcs)',"='2. Mùa vụ'!D13",NUM,"='2. Mùa vụ'!E13",PCT,'so với 26SS','Giảm','warning'),
('Chi phí gia công năm 2026 (USD)',f"='2. Mùa vụ'!E{y26}",USD,f"='2. Mùa vụ'!E{y26}/'2. Mùa vụ'!B{y26}",USD2,'USD/pcs bình quân','Theo dõi','good'),
('Tỷ lệ sử dụng nhà máy T11/2026–T1/2027',"="+CAP['u_now'],'0%',"="+CAP['u26'],'0%','bình quân T1–T10/2026','Kín lịch','good'),
('Tỷ lệ sử dụng nhà máy T3–T6/2027',"="+CAP['u_mj'],'0%',"="+CAP['pcs_gap'],NUM,'pcs cần bổ sung để đạt 90%','Thiếu đơn','critical'),
('Chi phí nhân công chưa có đơn T3–T6/2027 (USD)',"="+CAP['usd_gap'],USD,None,None,'nếu không lấp được công suất','Rủi ro tài chính','critical'),
('Tỷ trọng TNF trong sản lượng',f"='3. Khách hàng'!D{rt+3}",PCT,f"='3. Khách hàng'!D{rt+4}",PCT,'Top 3 brand','Tập trung cao','serious'),
('Chi phí chia lô mùa 27SS (USD)',"='2. Mùa vụ'!J13",USD,"='2. Mùa vụ'!K13",PCT,'% trên chi phí gia công','Đang cải thiện','good'),
('Sản lượng sắp SX trễ ETD ≥ 8 ngày (pcs)',"='6. Rủi ro giao hàng'!E7+'6. Rủi ro giao hàng'!E8",NUM,f'=SUMIFS({S6R("T")},{S6R("O")},"Nghiêm trọng")+SUMIFS({S6R("T")},{S6R("O")},"Cao")',USD,'giá trị gia công bị ảnh hưởng','Rủi ro uy tín','critical')]
r=kpi_table(SC,8,ce)+2
section(SC,r,'Xu hướng 4 năm'); r+=1
hdr(SC,r,['Năm','Sản lượng (pcs)','Tăng/giảm','Chi phí gia công (USD)','Chi phí chia lô (USD)'],h=22)
for k in range(4):
    rr_=r+1+k; src=yr+2+k
    SC.cell(rr_,1,f"='2. Mùa vụ'!A{src}"); SC.cell(rr_,2,f"='2. Mùa vụ'!B{src}"); SC.cell(rr_,3,f"='2. Mùa vụ'!C{src}"); SC.cell(rr_,4,f"='2. Mùa vụ'!E{src}"); SC.cell(rr_,5,f"='2. Mùa vụ'!F{src}")
    for c,fm in [(2,NUM),(3,PCT),(4,USD),(5,USD)]: SC.cell(rr_,c).number_format=fm
    for c in range(1,6): SC.cell(rr_,c).font=GREEN
body(SC,r+1,r+4,1,5); trend_r=r; r+=6
section(SC,r,'Danh mục khách hàng (pcs)'); r+=1
hdr(SC,r,['Brand','2025 (SS+FW)','2026 (SS+FW)','27SS','Xu hướng 27SS vs 26SS'],h=22)
brands=['TNF','MONTBELL','HELLY HANSEN','GOLDWIN','ALLTERRAIN 81','TETON BROS']
bi={b:5+k for k,b in enumerate(top+['(trống)'])}
for k,b in enumerate(brands):
    rr_=r+1+k; br_=bi.get(b)
    SC.cell(rr_,1,b)
    if br_:
        SC.cell(rr_,2,f"='3. Khách hàng'!F{br_}+'3. Khách hàng'!G{br_}"); SC.cell(rr_,3,f"='3. Khách hàng'!H{br_}+'3. Khách hàng'!I{br_}")
        SC.cell(rr_,4,f"='3. Khách hàng'!J{br_}"); SC.cell(rr_,5,f"=IFERROR('3. Khách hàng'!J{br_}/'3. Khách hàng'!H{br_}-1,\"Mới\")")
    for c,fm in [(2,NUM),(3,NUM),(4,NUM),(5,PCT)]: SC.cell(rr_,c).number_format=fm; SC.cell(rr_,c).font=GREEN
    SC.cell(rr_,1).font=F(bold=True,size=10)
body(SC,r+1,r+len(brands),1,5)
SC.conditional_formatting.add(f'E{r+1}:E{r+len(brands)}',FormulaRule(formula=[f'AND(ISNUMBER(E{r+1}),E{r+1}<-0.2)'],fill=PatternFill('solid',fgColor='F6C9C9')))
SC.conditional_formatting.add(f'E{r+1}:E{r+len(brands)}',FormulaRule(formula=[f'AND(ISNUMBER(E{r+1}),E{r+1}>0.2)'],fill=PatternFill('solid',fgColor='CFEFCF')))
r+=len(brands)+2
section(SC,r,'Quyết định cần CEO phê duyệt'); r+=1
dec=[('1. Lấp công suất T3–T6/2027',f'Duyệt mục tiêu chốt khoảng {N_["pcs_gap"]/1000:,.0f}k pcs đơn 27FW trước 15/12/2026. Cho phép giá linh hoạt (giảm giá gia công) cho đơn lấp tháng thấp điểm, ưu tiên brand đang tăng (HELLY HANSEN, ALLTERRAIN 81) và lấy lại sản lượng GOLDWIN.','Phương án dự phòng: tạm điều chuyển nhân lực, cho nghỉ phép bù, hoặc nhận gia công ngắn hạn.'),
('2. Chính sách lô tối thiểu',f'Áp dụng MOQ hoặc phụ phí cho lô ≤300 pcs từ 27FW. Lô này chiếm {K["lot_rows"]:.0%} số đơn nhưng chỉ {K["lot_qty"]:.0%} sản lượng, với đơn giá gia công gấp {V(K["lot_x"])} lần.','Mục tiêu: chi phí chia lô ≤10% chi phí gia công.'),
('3. Đa dạng hoá khách hàng','Đặt mục tiêu brand ngoài TNF ≥30% sản lượng năm 2027. Giao Sales tìm hiểu lý do MONTBELL giảm về 0 và lập kế hoạch giữ khách.',f'Rủi ro: một khách hàng chiếm {K["tnf"]:.0%} ({K["tnf27"]:.0%} riêng 27SS).'),
('4. Xử lý đơn 27SS trễ',f'Phê duyệt ngân sách tăng ca hoặc thêm line để kéo các đơn trễ ≥8 ngày (~{sev_hi.qty.sum()/1000:,.0f}k pcs, ~{VN1(risk_usd/1e6)} triệu USD giá trị gia công) về đúng hạn. Song song, đàm phán lùi 工場ETD với khách.','So sánh chi phí tăng ca với rủi ro phạt hoặc mất đơn.'),
('5. Chuẩn hoá dữ liệu kế hoạch','Giao Kế hoạch SX chuẩn hoá file PLANTABLE (danh mục, công thức, xoá dòng treo) để báo cáo này cập nhật hàng tuần.','Chi phí thấp, giúp ra quyết định nhanh.')]
r=table_rows(SC,r,['Quyết định','Nội dung đề xuất','Ghi chú / phương án'],dec,heights=48,spans=[(1,1),(2,4),(5,5)])
widths(SC,[40,40,26,24,30]); SC.sheet_view.showGridLines=False
cc=BarChart(); cc.type='col'; cc.title='Sản lượng theo năm (pcs)'; cc.y_axis.numFmt='#,##0'
cc.add_data(Reference(SC,min_col=2,min_row=trend_r,max_row=trend_r+4),titles_from_data=True); cc.set_categories(Reference(SC,min_col=1,min_row=trend_r+1,max_row=trend_r+4))
cc.series[0].graphicalProperties.solidFill='2A78D6'; cc.series[0].graphicalProperties.line.noFill=True; cc.legend=None; cc.gapWidth=60; cc.height=7; cc.width=13; chart_fix(cc); SC.add_chart(cc,'G7')

# ---------- 1B. GĐ Nhà máy ----------
title(SG,'TRANG GIÁM ĐỐC NHÀ MÁY – CÔNG SUẤT, NHÂN LỰC, HIỆU SUẤT','Xưởng APIC: 30 line thường × 32 CN = 960 CN may (line nhỏ 小ロットライン = 16 CN). Ngày tham chiếu 14/11/2026.')
section(SG,4,'Chỉ số vận hành')
i26a,i26b=IX[P('2026-01')],IX[P('2026-10')]
ge=[('Tỷ lệ sử dụng ngày công T11/2026–T1/2027',"="+CAP['u_now'],'0%',"="+CAP['u26'],'0%','bình quân T1–T10/2026','Kín lịch','good'),
('Tỷ lệ sử dụng T2/2027 (Tết)',"="+CAP['u_feb'],'0%',None,None,'Tết Nguyên đán 06/02/2027','Lập kế hoạch Tết','warning'),
('Tỷ lệ sử dụng T3–T6/2027',"="+CAP['u_mj'],'0%',"="+CAP['wd_gap'],NUM,'công-ngày cần bổ sung','Thiếu đơn','critical'),
('Hiệu suất yêu cầu bình quân 2026 (giờ chuẩn / giờ công)',f"=SUM('4. Công suất'!H{i26a}:H{i26b})/(SUM('4. Công suất'!D{i26a}:D{i26b})*8)",'0%',None,None,'giờ chuẩn theo SAM / giờ công kế hoạch','Theo dõi','warning'),
('Số mã hàng bắt đầu bình quân / tháng 2026',f"=AVERAGE('4. Công suất'!J{i26a}:J{i26b})",'0',f"=AVERAGE('4. Công suất'!J{i26a}:J{i26b})*32/'Phương pháp'!$C$5",'0.0','lần đổi mã / line / tháng','Đổi mã nhiều','serious'),
('Số đơn chạy song song / line / ngày (2026)',"="+CAP['over'],'0.00',None,None,'nhiều đơn cùng lúc trên 1 line','Theo dõi','warning'),
('Hiệu suất lô ≤300 pcs (平均日産/MAX)',"='5. Chi phí & Lô'!H5",'0%',"='5. Chi phí & Lô'!H9",'0%','lô 2.001–5.000','Mất năng suất','serious')]
r=kpi_table(SG,5,ge,note=(4,7),stat=(8,10))+2
section(SG,r,'Bản đồ tải theo line (tỷ lệ sử dụng ngày công theo tháng)'); r+=1
NOTE(SG,r,'Line thường = 32 CN (sửa được ở cột B). Line nhỏ 小ロットライン (vd 14A1, 14A2) = 16 CN, được tính vào line mẹ (14A). Đơn chạy chung AB chia đều cho A và B. Tỷ lệ = công-ngày có đơn / (32 × ngày làm việc). Line nhỏ có đơn tính 16 CN.'); r+=1
hm_months=list(pd.period_range('2026-11','2027-06',freq='M'))
hdr(SG,r,['Line','Định biên (CN)']+[m.strftime('%m/%Y') for m in hm_months]+['Đơn còn lại','SL còn lại (pcs)','Line trống từ ngày','Cấu hình đơn sắp tới'],h=32)
hr0=r+1
lines=[(float(l),h) for l in range(1,16) for h in ('A','B')]
fwd_all=o[(~o.stale)&(o.end>=REF)]
for k,(ln,hh) in enumerate(lines):
    rr_=hr0+k; d=Ld[(Ld.f=='APIC')&(Ld.line==ln)].copy(); d['k']=d.part.map(lambda p: hw(p,hh))
    SG.cell(rr_,1,f'{int(ln)}{hh}'); SG.cell(rr_,2,32); SG.cell(rr_,2).font=BLUE
    for j,m in enumerate(hm_months):
        wsum=int(OC[(OC.line==ln)&(OC.h==hh)&(OC.m==m)].occ.sum())
        c=SG.cell(rr_,3+j,f'={wsum}/($B{rr_}*{WD[m]})' if wsum>0 else None); c.number_format='0%'
    ff=fwd_all[(fwd_all.factory=='APIC')&(fwd_all.line==ln)]; ff=ff[ff.part.map(lambda p: hw(p,hh))>0]
    SG.cell(rr_,3+len(hm_months),len(ff)); SG.cell(rr_,4+len(hm_months),float(ff.qty.sum())); SG.cell(rr_,4+len(hm_months)).number_format=NUM
    le=ff.end.max(); SG.cell(rr_,5+len(hm_months),(le+pd.Timedelta(days=1)).to_pydatetime() if pd.notna(le) else None); SG.cell(rr_,5+len(hm_months)).number_format=DATE
    ps=sorted(set(str(p) for p in ff.part.dropna()))
    small=[f'{int(ln)}{p}' for p in ps if p[-1].isdigit() and not p.startswith('AB')]
    cfg='Line nhỏ: '+', '.join(small) if small and all(p[-1].isdigit() or p.startswith('AB') for p in ps) else ('Thường + line nhỏ: '+', '.join(small) if small else ('Thường' if ps else ''))
    SG.cell(rr_,6+len(hm_months),cfg).font=F(size=9)
    SG.cell(rr_,1).font=F(bold=True,size=10)
hr1=hr0+len(lines)-1
body(SG,hr0,hr1,1,6+len(hm_months))
from openpyxl.formatting.rule import ColorScaleRule
SG.conditional_formatting.add(f'C{hr0}:{CL(2+len(hm_months))}{hr1}',ColorScaleRule(start_type='num',start_value=0,start_color='FFFFFF',end_type='num',end_value=1.1,end_color='3987E5'))
SG.conditional_formatting.add(f'{CL(5+len(hm_months))}{hr0}:{CL(5+len(hm_months))}{hr1}',FormulaRule(formula=[f"AND(ISNUMBER({CL(5+len(hm_months))}{hr0}),{CL(5+len(hm_months))}{hr0}<DATE(2027,3,15))"],fill=PatternFill('solid',fgColor='F6C9C9')))
r=hr1+1; NOTE(SG,r,'Cột "Line trống từ ngày" tô đỏ = line hết đơn trước 15/03/2027, cần giao đơn 27FW sớm nhất.'); r+=1
r+=1
section(SG,r,'Hiệu quả theo cụm line (A+B) – các đơn bắt đầu T1–T10/2026'); r+=1
hdr(SG,r,['Cụm line','Số đơn','Sản lượng (pcs)','Giờ chuẩn','SAM TB (phút/sp)','Lô TB (pcs)','CP gia công (USD)','Đơn giá GC (USD/pcs)','Hiệu suất lô TB'],h=32)
pr0=r+1
D1,D2='DATE(2026,1,1)','DATE(2026,11,1)'
for k in range(15):
    rr_=pr0+k; ln=k+1
    cond=f'{R("C")},"APIC",{R("D")},{ln},{R("P")},">="&{D1},{R("P")},"<"&{D2},{R("AK")},"<>Y"'
    SG.cell(rr_,1,ln); SG.cell(rr_,2,f'=COUNTIFS({cond})'); SG.cell(rr_,3,f'=SUMIFS({R("M")},{cond})'); SG.cell(rr_,4,f'=SUMIFS({R("O")},{cond})')
    SG.cell(rr_,5,f'=IFERROR(D{rr_}*60/C{rr_},"")'); SG.cell(rr_,6,f'=IFERROR(C{rr_}/B{rr_},"")'); SG.cell(rr_,7,f'=SUMIFS({R("AE")},{cond})')
    SG.cell(rr_,8,f'=IFERROR(G{rr_}/C{rr_},"")'); SG.cell(rr_,9,f'=IFERROR(AVERAGEIFS({R("AH")},{cond}),"")')
    for c,fm in [(2,NUM),(3,NUM),(4,NUM),(5,'0'),(6,NUM),(7,USD),(8,USD2),(9,'0%')]: SG.cell(rr_,c).number_format=fm
    SG.cell(rr_,1).font=F(bold=True,size=10)
pr1=pr0+14; SG.cell(pr1+1,1,'Tổng / TB')
for c in (2,3,4,7): SG.cell(pr1+1,c,f'=SUM({CL(c)}{pr0}:{CL(c)}{pr1})')
SG.cell(pr1+1,5,f'=D{pr1+1}*60/C{pr1+1}'); SG.cell(pr1+1,6,f'=C{pr1+1}/B{pr1+1}'); SG.cell(pr1+1,8,f'=G{pr1+1}/C{pr1+1}'); SG.cell(pr1+1,9,f'=AVERAGE(I{pr0}:I{pr1})')
for c,fm in [(2,NUM),(3,NUM),(4,NUM),(5,'0'),(6,NUM),(7,USD),(8,USD2),(9,'0%')]: SG.cell(pr1+1,c).number_format=fm
body(SG,pr0,pr1+1,1,9)
for c in range(1,10): SG.cell(pr1+1,c).fill=TOT_FILL; SG.cell(pr1+1,c).font=F(bold=True,size=10)
SG.conditional_formatting.add(f'H{pr0}:H{pr1}',ColorScaleRule(start_type='min',start_color='FFFFFF',end_type='max',end_color='EB6834'))
r=pr1+2; NOTE(SG,r,'Đơn giá GC cao thường do SAM cao (hàng phức tạp) hoặc lô nhỏ. Cụm 14 chạy toàn line nhỏ (14A1, 14A2, 14B1, 14B2); các cụm 3, 9, 13, 15 cũng thường chia line nhỏ.'); r+=2
section(SG,r,'Việc Giám đốc nhà máy cần làm'); r+=1
gact=[('Lập kế hoạch nhân lực T3–T6/2027',f'Với tải hiện tại chỉ {N_["u_mar"]:.0%} ở T3 và gần 0 từ T4, cần chuẩn bị kịch bản: (a) có đơn 27FW, (b) thiếu đơn thì bố trí nghỉ phép, đào tạo đa kỹ năng, chuyển line. Báo cáo CEO trước 30/11.','Trước 30/11/2026'),
('Kế hoạch Tết 2027',f'T11/2026–T1/2027 tải ~{N_["u_now"]:.0%}, còn T2 bị Tết cắt ngày. Chốt lịch nghỉ, tăng ca trước Tết, ưu tiên đơn có 工場ETD trước Tết.','Trước 15/12/2026'),
('Giảm số lần đổi mã','Khoảng 2 lần đổi mã / line / tháng (mỗi line 32 CN). Gom lô cùng style hoặc cùng 集約CODE vào một line. Dành line nhỏ 小ロットライン (14A1, 14A2…) cho lô ≤300.','Liên tục'),
('Kiểm soát đơn chạy song song','Nhiều line chạy 2 đơn cùng lúc (song song AB, A với A1/A2). Theo dõi năng suất thực tế từng đơn và cân đối nhân lực giữa các đơn.','Hàng tháng'),
('Năng suất lô nhỏ',f'Lô ≤300 chỉ đạt ~{K["eff_s"]:.0%} tốc độ tối đa.'+' Áp dụng mẫu chuẩn hoá, cử tổ trưởng kèm line và giảm thời gian học việc.','Từ 27FW')]
r=table_rows(SG,r,['Việc','Nội dung','Thời hạn'],gact,heights=36,spans=[(1,1),(2,10),(11,13)])
widths(SG,[44,13]+[11]*len(hm_months)+[9,12,13,30]); SG.sheet_view.showGridLines=False

# ---------- 1C. QL Sản xuất ----------
title(SQ,'TRANG QUẢN LÝ SẢN XUẤT – LỊCH LINE & ĐỘ SẴN SÀNG ĐƠN HÀNG','Dùng cho họp sản xuất hằng tuần. Lịch theo tuần (bắt đầu thứ Hai) và danh sách đơn bắt đầu trong 6 tuần tới.')
H=6; W0=pd.Timestamp('2026-11-16')
nx=fw[(fw.start>=REF)&(fw.start<REF+pd.Timedelta(weeks=H))].sort_values(['start','factory','line'])
section(SQ,4,'Tình trạng 6 tuần tới')
lr0=4+6+2+2+1+len(lines)+1+2  # readiness list header row
def RL(c): return f'${c}${lr0+1}:${c}${lr0+len(nx)}'
qe=[('Số đơn bắt đầu trong 6 tuần',f'=COUNTA({RL("C")})',NUM,f'=SUM({RL("F")})',NUM,'pcs','Khối lượng','good'),
('Đơn sẵn sàng',f'=COUNTIF({RL("Q")},"Sẵn sàng")',NUM,f'=IFERROR(B6/B5,"")',PCT,'trên tổng số đơn','Mục tiêu ≥90%','warning'),
('Đơn vướng NVL (trễ / chưa xác định)',f'=COUNTIF({RL("K")},"Trễ")+COUNTIF({RL("K")},"Chưa xác định")',NUM,f'=SUMIF({RL("K")},"Trễ",{RL("F")})+SUMIF({RL("K")},"Chưa xác định",{RL("F")})',NUM,'pcs','Đôn đốc mua hàng','critical'),
('Đơn chưa duyệt mẫu',f'=COUNTIF({RL("L")},"Chưa duyệt")',NUM,f'=SUMIF({RL("L")},"Chưa duyệt",{RL("F")})',NUM,'pcs','Đôn đốc QA / khách','serious'),
('Đơn ETD trễ ≥8 ngày',f'=COUNTIF({RL("P")},"Nghiêm trọng")+COUNTIF({RL("P")},"Cao")',NUM,f'=SUMIF({RL("P")},"Nghiêm trọng",{RL("F")})+SUMIF({RL("P")},"Cao",{RL("F")})',NUM,'pcs','Cân đối lại line','critical'),
('Đơn thiếu 工場ETD',f'=COUNTIF({RL("P")},"Thiếu 工場ETD")',NUM,None,None,'yêu cầu Sales nhập','Bổ sung dữ liệu','warning')]
r=kpi_table(SQ,4,qe,note=(4,6),stat=(7,9))
# fix kpi references to B5/B6 rows: kpi rows start at 5
r+=2
section(SQ,r,f'Lịch line theo tuần (16 tuần từ {W0:%d/%m/%Y})'); r+=1
NOTE(SQ,r,'Mỗi ô hiển thị style có SL lớn nhất đang chạy trong tuần trên line đó ([A1] = line nhỏ 小ロットライン; +n = thêm n đơn khác). Đỏ = có đơn trễ ETD ≥30 ngày hoặc NVL về sau START. Cam = trễ 8–29 ngày hoặc NVL 未定. Xanh nhạt = đang chạy bình thường. Tím = chỉ chạy hàng SMS (mẫu chào hàng, hỗ trợ chuyền mẫu). Xám = line trống.'); r+=1
weeks=[W0+pd.Timedelta(weeks=i) for i in range(16)]
hdr(SQ,r,['Line']+[f'{w:%d/%m}' for w in weeks],h=22); g0=r+1
fl=fw.copy(); fl['crit']=(fl.gap<=-30)|fl.mat_late; fl['high']=((fl.gap<=-8)&(fl.gap>-30))|fl.mat_unk
_fd=O_DOT[(O_DOT.end>=REF)&(~O_DOT.stale)].copy(); _fd['crit']=False; _fd['high']=False; _fd['qty']=_fd.qty_raw
_fs=O_SMS[(O_SMS.end>=REF)&(~O_SMS.stale)].copy(); _fs['crit']=False; _fs['high']=False
fl['sms']=False; _fd['sms']=False; _fs['sms']=True
fl=pd.concat([fl,_fd,_fs],ignore_index=True)
for k,(ln,hh) in enumerate(lines):
    rr_=g0+k; SQ.cell(rr_,1,f'{int(ln)}{hh}').font=F(bold=True,size=10)
    sub=fl[(fl.factory=='APIC')&(fl.line==ln)]; sub=sub[sub.part.map(lambda p: hw(p,hh))>0]
    for j,ws_ in enumerate(weeks):
        we=ws_+pd.Timedelta(days=6); act=sub[(sub.start<=we)&(sub.end>=ws_)]
        c=SQ.cell(rr_,2+j); c.font=F(size=8); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=BOX
        if len(act)==0: c.value='trống'; c.fill=PatternFill('solid',fgColor='EDEDED'); c.font=F(size=8,color='8A8A8A'); continue
        tp=act.sort_values('qty',ascending=False).iloc[0]; pp=str(tp['part']) if isinstance(tp['part'],str) and tp['part'][-1].isdigit() else ''; c.value=(f'[{pp}] ' if pp else '')+str(tp['style'])[:12]+(f' +{len(act)-1}' if len(act)>1 else '')
        c.fill=PatternFill('solid',fgColor='F6C9C9' if act.crit.any() else 'FCE3B0' if act.high.any() else 'DCE9F9')
        if act.sms.fillna(False).all(): c.value='[SMS] '+str(tp['style'])[:10]; c.fill=PatternFill('solid',fgColor='E4DEF7')
    SQ.cell(rr_,1).border=BOX; SQ.row_dimensions[rr_].height=24
r=g0+len(lines)+1
section(SQ,r,'Danh sách đơn bắt đầu trong 6 tuần tới – kiểm tra sẵn sàng'); r+=1
NOTE(SQ,r,'CUT = START − 10 ngày. NVL "OK" khi ngày có thể đưa vào SX ≤ START. Kết luận tự cập nhật khi sửa ngày hoặc trạng thái.'); r+=1
assert lr0==r, (lr0,r)
head=['Xưởng','Line','Part','Brand','Style','QTY','CUT','START','END','Ngày NVL vào SX được','NVL','Mẫu','工場ETD','ETD KH','Lệch ETD (ngày)','Mức ETD','Kết luận','Item','Dòng gốc']
# reorder so letters match RL(): C=Part? -> define explicit mapping below
head=['Line','Cụm line','Style','Brand','Part','QTY','CUT','START','END','Ngày NVL vào SX được','NVL','Mẫu','工場ETD','ETD KH','Lệch ETD (ngày)','Mức ETD','Kết luận','Item','Dòng gốc']
hdr(SQ,lr0,head,h=36)
for i,x in enumerate(nx.itertuples(),lr0+1):
    vals=[f'{int(x.line)}{x.part}' if pd.notna(x.line) and isinstance(x.part,str) else x.factory,v(x.line),x.style,x.brand,x.part,v(x.qty),f'=H{i}-10',v(x.start),v(x.end),v(x.input_ok) if pd.notna(x.input_ok) else x.input_txt,
          'Chưa xác định' if x.mat_unk else ('OK' if isinstance(x.input_txt,str) and 'OK' in x.input_txt.upper() else f'=IF(ISNUMBER(J{i}),IF(J{i}<=H{i},"OK","Trễ"),"Chưa có ngày")'),'OK' if x.smp_ok else 'Chưa duyệt',
          v(x.fac_etd),v(x.etd_plan),f'=IF(OR(M{i}="",N{i}=""),"",M{i}-N{i})',
          f"=IF(M{i}=\"\",\"Thiếu 工場ETD\",IF(O{i}<='Phương pháp'!$C$6,\"Nghiêm trọng\",IF(O{i}<='Phương pháp'!$C$7,\"Cao\",IF(O{i}<0,\"Theo dõi\",\"Đúng hạn\"))))",
          f'=IF(AND(OR(K{i}="OK",K{i}="Chưa có ngày"),L{i}="OK",OR(P{i}="Đúng hạn",P{i}="Theo dõi")),"Sẵn sàng",TRIM(IF(OR(K{i}="Trễ",K{i}="Chưa xác định"),"NVL; ","")&IF(L{i}<>"OK","Mẫu; ","")&IF(OR(P{i}="Nghiêm trọng",P{i}="Cao"),"ETD; ","")&IF(P{i}="Thiếu 工場ETD","工場ETD","")))',
          x.item,x.row]
    for j,val in enumerate(vals,1): SQ.cell(i,j,val).font=BODY; SQ.cell(i,j).border=BOX
    for j in (7,8,9,10,13,14): SQ.cell(i,j).number_format=DATE
    SQ.cell(i,6).number_format=NUM
lr1=lr0+len(nx)
for col,txt,clr in [('Q','Sẵn sàng','CFEFCF'),('K','Trễ','F9D3C2'),('K','Chưa xác định','FCE3B0'),('L','Chưa duyệt','FCE3B0'),('P','Nghiêm trọng','F6C9C9'),('P','Cao','F9D3C2'),('P','Thiếu 工場ETD','E0E0E0')]:
    SQ.conditional_formatting.add(f'{col}{lr0+1}:{col}{lr1}',FormulaRule(formula=[f'{col}{lr0+1}="{txt}"'],fill=PatternFill('solid',fgColor=clr)))
SQ.conditional_formatting.add(f'Q{lr0+1}:Q{lr1}',FormulaRule(formula=[f'AND(Q{lr0+1}<>"",Q{lr0+1}<>"Sẵn sàng")'],fill=PatternFill('solid',fgColor='F6C9C9')))
SQ.auto_filter.ref=f'A{lr0}:S{lr1}'
r=lr1+2; section(SQ,r,'Hàng SMS hỗ trợ chuyền mẫu – đang/sắp chạy (tách riêng, không tính là hàng sản xuất)'); r+=1
NOTE(SQ,r,'SMS = mẫu chào hàng / mẫu triển lãm (展示会サンプル) chạy trên line nhỏ để hỗ trợ chuyền mẫu. Không cộng vào sản lượng, chi phí, rủi ro giao hàng của hàng sản xuất; vẫn tính chiếm line.'); r+=1
_sm=O_SMS[(O_SMS.end>=REF)&(~O_SMS.stale)].sort_values('start')
hdr(SQ,r,['Line','Brand','Mã SMS / style','SL (pcs)','START','END','Nội dung (ghi chú gốc)'],h=24); _s0=r+1
for i,x in enumerate(_sm.itertuples(),_s0):
    vals=[f'{int(x.line)}{x.part}' if pd.notna(x.line) and isinstance(x.part,str) else '',x.brand,x.style,v(x.qty),v(x.start),v(x.end),(x.comment if isinstance(x.comment,str) else '')[:140]]
    for j,val in enumerate(vals,1): SQ.cell(i,j,val).font=BODY; SQ.cell(i,j).border=BOX
    SQ.cell(i,5).number_format=DATE; SQ.cell(i,6).number_format=DATE; SQ.cell(i,4).number_format=NUM
    SQ.merge_cells(start_row=i,start_column=7,end_row=i,end_column=17); SQ.cell(i,7).alignment=Alignment(wrap_text=False)
r=_s0+len(_sm)+1
section(SQ,r,'Nhịp làm việc đề xuất cho Quản lý sản xuất'); r+=1
qact=[('Thứ Hai','Họp line: rà lịch tuần trên Gantt và xử lý các ô đỏ/cam. Cập nhật START/END thực tế vào PLANTABLE.'),
('Thứ Tư','Kiểm tra "sẵn sàng" cho các đơn có CUT trong 10 ngày tới (NVL, mẫu PPS, rập, phụ liệu). Đơn nào chưa sẵn sàng thì báo Mua hàng/QA.'),
('Thứ Sáu','Gửi Sales danh sách đơn ETD trễ ≥8 ngày và đơn thiếu 工場ETD để thống nhất ngày giao mới.'),
('Hằng ngày','Theo dõi sản lượng ngày so với 予定日産数. Line đạt dưới 80% kế hoạch 2 ngày liên tiếp thì báo GĐ nhà máy.')]
r=text_rows(SQ,r,qact,ncol=10,w=120)
widths(SQ,[16]+[10]*16+[10,10]); SQ.column_dimensions['A'].width=24; SQ.column_dimensions['Q'].width=20; SQ.column_dimensions['R'].width=24; SQ.sheet_view.showGridLines=False
SQ.freeze_panes='B5' if False else None
QL_LR0=lr0

# ---------- 0. Mục lục ----------
title(SX,'MỤC LỤC – BÁO CÁO ĐIỀU HÀNH KẾ HOẠCH SẢN XUẤT','Nguồn: APIC P-Plan 2026 (2026-11-13), sheet PLANTABLE (原価) 2026年. Ngày tham chiếu 14/11/2026. Bấm vào tên sheet để mở.')
hdr(SX,4,['Người đọc','Đọc trước','Đọc thêm khi cần','Câu hỏi được trả lời'],h=22)
idx=[('CEO','1A. CEO','1. Tóm tắt điều hành; 2A. So sánh mùa; 3. Khách hàng','Kinh doanh đang đi về đâu? Rủi ro tài chính nào? Cần quyết định gì?'),
('Giám đốc nhà máy','1B. GĐ Nhà máy','4. Công suất; 4A. 受注進行表; 2A. So sánh mùa','Nhà máy đủ hay thiếu đơn? Line nào trống, khi nào? Nhân lực, Tết, hiệu suất?'),
('Quản lý sản xuất','1C. QL Sản xuất','6. Rủi ro giao hàng; 7. Chất lượng DL','Tuần này/6 tuần tới chạy gì, đơn nào chưa sẵn sàng, đơn nào trễ?'),
('Tất cả','1. Tóm tắt điều hành','Phương pháp; Data','7 phát hiện chính, 6 khuyến nghị; định nghĩa chỉ số và dữ liệu sạch.')]
for k,(a,b,c,d) in enumerate(idx,5):
    SX.cell(k,1,a).font=F(bold=True,size=11); x=SX.cell(k,2,b); x.hyperlink=f"#'{b}'!A1"; x.font=F(size=11,color='2A78D6',underline='single')
    SX.cell(k,3,c).font=BODY; SX.cell(k,4,d).font=BODY
    for j in range(1,5): SX.cell(k,j).alignment=Alignment(wrap_text=True,vertical='center'); SX.cell(k,j).border=BOX
    SX.row_dimensions[k].height=36
SX['A11']='Danh sách sheet'; SX['A11'].font=F(bold=True,size=11)
desc={'1. Tóm tắt điều hành':'KPI tổng hợp, phát hiện chính, khuyến nghị','1A. CEO':'Bức tranh kinh doanh, khách hàng, quyết định cần phê duyệt','1B. GĐ Nhà máy':'Tải theo line, nhân lực, hiệu suất line, Tết','1C. QL Sản xuất':'Lịch line 16 tuần, checklist sẵn sàng 6 tuần tới',
'2. Mùa vụ':'Sản lượng, giờ chuẩn, chi phí theo mùa & năm','2A. So sánh mùa':'So với mùa trước & cùng kỳ: tổng quan, brand, quy mô lô, SAM, ngày may','3. Khách hàng':'Brand × mùa, tập trung khách hàng','4. Công suất':'Ngày công kế hoạch vs định biên theo tháng','4A. 受注進行表':'受注進行表 tự dựng từ PLANTABLE: bảng line × nửa tháng, độ phủ đã chốt / tạm tính, xu hướng theo bản chụp','5. Chi phí & Lô':'Đơn giá gia công theo quy mô lô',
'6. Rủi ro giao hàng':'Toàn bộ đơn đang/sắp SX với cờ rủi ro','7. Chất lượng DL':'Lỗi dữ liệu cần sửa ở file gốc','Data':'Dữ liệu đã làm sạch (nguồn cho mọi công thức)','Phương pháp':'Giả định có thể chỉnh (ô xanh), ánh xạ cột, định nghĩa'}
for k,(n,dsc) in enumerate(desc.items(),12):
    x=SX.cell(k,1,n); x.hyperlink=f"#'{n}'!A1"; x.font=F(size=10,color='2A78D6',underline='single'); SX.cell(k,2,dsc).font=BODY
    SX.merge_cells(start_row=k,start_column=2,end_row=k,end_column=4)
SX.cell(27,1,'Cách cập nhật: dán dữ liệu PLANTABLE mới → chạy lại script phân tích (hoặc nhờ Claude) → các sheet tự tính lại theo sheet Data và các ô xanh ở sheet Phương pháp.').font=F(italic=True,size=9,color='52514E')
widths(SX,[24,24,38,60]); SX.sheet_view.showGridLines=False


# ---------- 8. So sánh mùa (thống kê đa chiều) ----------
S8=wb.create_sheet('2A. So sánh mùa')
title(S8,'Thống kê & so sánh theo mùa, nhãn hàng, quy mô lô, thời gian chế tạo','Mùa trước = mùa liền kề trước (27SS so với 26FW). Cùng kỳ = cùng mùa năm trước (27SS so với 26SS). Chỉ hàng sản xuất: đã bỏ dòng chấm (chuyền lớn ghi 2 lần) và hàng SMS. Đổi 3 ô chữ xanh để so sánh mùa khác.')
S8['A4']='Mùa phân tích'; S8['C4']='27SS'; S8['A5']='Mùa trước'; S8['C5']='26FW'; S8['A6']='Cùng kỳ năm trước'; S8['C6']='26SS'
for r_ in (4,5,6): S8.cell(r_,1).font=F(bold=True,size=10); S8.cell(r_,3).font=F(bold=True,size=11,color='0000FF'); S8.cell(r_,3).fill=PatternFill('solid',fgColor='FFF2CC'); S8.cell(r_,3).border=BOX
S8['D4']='Gõ mã mùa: 23SS … 27FW'; S8['D4'].font=F(italic=True,size=9,color='52514E')
M_,P_,C_='$C$4','$C$5','$C$6'
def SQ_(crit,extra=''): return f'SUMIFS({R("M")},{R("B")},{crit}{extra})'
def CN_(crit,extra=''): return f'COUNTIFS({R("B")},{crit}{extra})'
# ---- A. tổng quan 3 mùa ----
r=8; section(S8,r,'A. So sánh tổng quan 3 mùa'); r+=1
hdr(S8,r,['Chỉ số','Mùa phân tích','Mùa trước','Cùng kỳ','Δ vs mùa trước','Δ vs cùng kỳ'],h=30)
S8.cell(r,2,f'={M_}'); S8.cell(r,3,f'={P_}'); S8.cell(r,4,f'={C_}')
for c_ in (2,3,4): S8.cell(r,c_).font=HDR_FONT
mets=[('Số dòng đơn',lambda s: f'={CN_(s)}',NUM),
('Sản lượng (pcs)',lambda s: f'={SQ_(s)}',NUM),
('Giờ chuẩn',lambda s: f'=SUMIFS({R("O")},{R("B")},{s})',NUM),
('SAM TB – thời gian chế tạo (phút/sp)',lambda s: f'=IFERROR(SUMIFS({R("O")},{R("B")},{s})*60/SUMIFS({R("M")},{R("B")},{s},{R("N")},"<>"),"")','0.0'),
('Lô TB (pcs / dòng đơn)',lambda s: f'=IFERROR({SQ_(s)}/{CN_(s)},"")',NUM),
('Số ngày may TB / đơn',lambda s: f'=IFERROR(AVERAGEIFS({R("AM")},{R("B")},{s}),"")','0.0'),
('Chi phí gia công (USD)',lambda s: f'=SUMIFS({R("AE")},{R("B")},{s})',USD),
('Đơn giá GC (USD/pcs)',lambda s: f'=IFERROR(SUMIFS({R("AE")},{R("B")},{s})/SUMIFS({R("M")},{R("B")},{s},{R("AE")},"<>"),"")',USD2),
('Chi phí chia lô (USD)',lambda s: f'=SUMIFS({R("AF")},{R("B")},{s})',USD),
('% SL lô ≤300',lambda s: f'=IFERROR({SQ_(s,","+R("AG")+",\"1. ≤300\"")}/{SQ_(s)},"")',PCT),
('% SL hàng 別注',lambda s: f'=IFERROR({SQ_(s,","+R("I")+",\"Y\"")}/{SQ_(s)},"")',PCT),
('% SL TNF',lambda s: f'=IFERROR({SQ_(s,","+R("G")+",\"TNF\"")}/{SQ_(s)},"")',PCT)]
a0=r+1
for k_,(lab,fn,nf) in enumerate(mets):
    rr_=a0+k_; S8.cell(rr_,1,lab).font=F(bold=True,size=10)
    for j_,cr in enumerate((M_,P_,C_)): S8.cell(rr_,2+j_,fn(cr)).number_format=nf
    if nf==PCT:
        S8.cell(rr_,5,f'=IFERROR(B{rr_}-C{rr_},"")'); S8.cell(rr_,6,f'=IFERROR(B{rr_}-D{rr_},"")'); S8.cell(rr_,5).number_format='+0.0%;-0.0%;0.0%'; S8.cell(rr_,6).number_format='+0.0%;-0.0%;0.0%'
    else:
        S8.cell(rr_,5,f'=IFERROR(B{rr_}/C{rr_}-1,"")'); S8.cell(rr_,6,f'=IFERROR(B{rr_}/D{rr_}-1,"")'); S8.cell(rr_,5).number_format='+0.0%;-0.0%;0.0%'; S8.cell(rr_,6).number_format='+0.0%;-0.0%;0.0%'
a1=a0+len(mets)-1; body(S8,a0,a1,1,6)
for col in ('E','F'):
    S8.conditional_formatting.add(f'{col}{a0}:{col}{a1}',FormulaRule(formula=[f'AND(ISNUMBER({col}{a0}),{col}{a0}<-0.1)'],font=Font(color='D03B3B',bold=True)))
    S8.conditional_formatting.add(f'{col}{a0}:{col}{a1}',FormulaRule(formula=[f'AND(ISNUMBER({col}{a0}),{col}{a0}>0.1)'],font=Font(color='008000',bold=True)))
r=a1+1; NOTE(S8,r,'Δ của các dòng % là chênh lệch điểm phần trăm. Đỏ = giảm >10%, xanh = tăng >10% (với chi phí/đơn giá, tăng là bất lợi).'); r+=2
# ---- B. theo nhãn hàng ----
section(S8,r,'B. Theo nhãn hàng (brand)'); r+=1
hb=['Brand (top 12 của 25SS–27FW)','SL mùa phân tích','SL mùa trước','SL cùng kỳ','Δ vs mùa trước','Δ vs cùng kỳ','Tỷ trọng (mùa phân tích)','SAM TB (phút) – phân tích','SAM TB (phút) – cùng kỳ','Đơn giá GC – phân tích','Đơn giá GC – cùng kỳ','Lô TB – phân tích','Lô TB – cùng kỳ']
hdr(S8,r,hb,h=45); b0=r+1
top12=[b for b in o[o.season.isin(['25SS','25FW','26SS','26FW','27SS','27FW'])].groupby('brand').qty.sum().sort_values(ascending=False).index if b!='(trống)'][:12]+['(trống)']
for k_,b in enumerate(top12):
    rr_=b0+k_; S8.cell(rr_,1,b).font=F(bold=True,size=10)
    g=f',{R("G")},$A{rr_}'
    for j_,cr in enumerate((M_,P_,C_)): S8.cell(rr_,2+j_,f'={SQ_(cr,g)}').number_format=NUM
    S8.cell(rr_,5,f'=IFERROR(B{rr_}/C{rr_}-1,IF(B{rr_}>0,"Mới",""))'); S8.cell(rr_,6,f'=IFERROR(B{rr_}/D{rr_}-1,IF(B{rr_}>0,"Mới",""))')
    S8.cell(rr_,7,f'=IFERROR(B{rr_}/{SQ_(M_)},"")')
    for j_,cr in ((8,M_),(9,C_)): S8.cell(rr_,j_,f'=IFERROR(SUMIFS({R("O")},{R("B")},{cr}{g})*60/SUMIFS({R("M")},{R("B")},{cr}{g},{R("N")},"<>"),"")').number_format='0'
    for j_,cr in ((10,M_),(11,C_)): S8.cell(rr_,j_,f'=IFERROR(SUMIFS({R("AE")},{R("B")},{cr}{g})/SUMIFS({R("M")},{R("B")},{cr}{g},{R("AE")},"<>"),"")').number_format=USD2
    for j_,cr in ((12,M_),(13,C_)): S8.cell(rr_,j_,f'=IFERROR({SQ_(cr,g)}/{CN_(cr,g)},"")').number_format=NUM
    for j_ in (5,6,7): S8.cell(rr_,j_).number_format='+0%;-0%;0%' if j_<7 else '0.0%'
rr_=b0+len(top12); S8.cell(rr_,1,'Khác').font=F(bold=True,size=10)
for j_,c_ in ((2,'B'),(3,'C'),(4,'D')): S8.cell(rr_,j_,f'={SQ_([M_,P_,C_][j_-2])}-SUM({c_}{b0}:{c_}{rr_-1})').number_format=NUM
S8.cell(rr_,5,f'=IFERROR(B{rr_}/C{rr_}-1,"")').number_format='+0%;-0%;0%'; S8.cell(rr_,6,f'=IFERROR(B{rr_}/D{rr_}-1,"")').number_format='+0%;-0%;0%'; S8.cell(rr_,7,f'=IFERROR(B{rr_}/{SQ_(M_)},"")').number_format='0.0%'
b1=rr_; body(S8,b0,b1,1,13)
for col in ('E','F'):
    S8.conditional_formatting.add(f'{col}{b0}:{col}{b1}',FormulaRule(formula=[f'AND(ISNUMBER({col}{b0}),{col}{b0}<-0.2)'],fill=PatternFill('solid',fgColor='F6C9C9')))
    S8.conditional_formatting.add(f'{col}{b0}:{col}{b1}',FormulaRule(formula=[f'AND(ISNUMBER({col}{b0}),{col}{b0}>0.2)'],fill=PatternFill('solid',fgColor='CFEFCF')))
r=b1+2
# ---- C/D/E: bands for 3 seasons ----
def band_block(r,ttl,col,bands,note=None):
    section(S8,r,ttl); r+=1
    if note: NOTE(S8,r,note); r+=1
    hdr(S8,r,['Nhóm','SL mùa phân tích','SL mùa trước','SL cùng kỳ','% SL (phân tích)','% SL (mùa trước)','% SL (cùng kỳ)','Δ SL vs mùa trước','Δ SL vs cùng kỳ','Số đơn (phân tích)','Số đơn (cùng kỳ)','Đơn giá GC – phân tích','Đơn giá GC – cùng kỳ'],h=45)
    x0=r+1
    for k_,bd in enumerate(bands):
        rr_=x0+k_; S8.cell(rr_,1,bd).font=F(bold=True,size=10); g=f',{R(col)},$A{rr_}'
        for j_,cr in enumerate((M_,P_,C_)): S8.cell(rr_,2+j_,f'={SQ_(cr,g)}').number_format=NUM
        for j_,(c_,cr) in enumerate((('B',M_),('C',P_),('D',C_))): S8.cell(rr_,5+j_,f'=IFERROR({c_}{rr_}/{SQ_(cr)},"")').number_format='0.0%'
        S8.cell(rr_,8,f'=IFERROR(B{rr_}/C{rr_}-1,"")').number_format='+0%;-0%;0%'; S8.cell(rr_,9,f'=IFERROR(B{rr_}/D{rr_}-1,"")').number_format='+0%;-0%;0%'
        S8.cell(rr_,10,f'={CN_(M_,g)}').number_format=NUM; S8.cell(rr_,11,f'={CN_(C_,g)}').number_format=NUM
        for j_,cr in ((12,M_),(13,C_)): S8.cell(rr_,j_,f'=IFERROR(SUMIFS({R("AE")},{R("B")},{cr}{g})/SUMIFS({R("M")},{R("B")},{cr}{g},{R("AE")},"<>"),"")').number_format=USD2
    x1=x0+len(bands)-1; body(S8,x0,x1,1,13)
    S8.conditional_formatting.add(f'E{x0}:G{x1}',ColorScaleRule(start_type='num',start_value=0,start_color='FFFFFF',end_type='max',end_color='86B6EF'))
    return x1+2
LOTB=['1. ≤300','2. 301–500','3. 501–1.000','4. 1.001–2.000','5. 2.001–5.000','6. >5.000']
SAMB=['1. <60 phút','2. 60–119 phút','3. 120–179 phút','4. 180–239 phút','5. ≥240 phút']
DAYB=['1. ≤5 ngày','2. 6–10 ngày','3. 11–20 ngày','4. 21–40 ngày','5. >40 ngày']
r=band_block(r,'C. Theo quy mô lô sản xuất','AG',LOTB,'Lô = lô sản xuất thực tế (集約, giá vốn 2).')
r=band_block(r,'D. Theo thời gian chế tạo sản phẩm (SAM, phút/sp)','AN',SAMB,'SAM = tổng thời gian chuẩn may 1 sản phẩm (cột TIME). Hàng phức tạp (jacket kỹ thuật) thường trên 180 phút.')
r=band_block(r,'E. Theo thời gian may của đơn (START → END, ngày)','AO',DAYB)
# ---- F. xu hướng mọi mùa ----
section(S8,r,'F. Xu hướng tất cả các mùa (tỷ trọng sản lượng)'); r+=1
def trend_block(r,lab,col,bands):
    hdr(S8,r,[lab]+ORDER,h=22); y0=r+1
    for k_,bd in enumerate(bands):
        rr_=y0+k_; S8.cell(rr_,1,bd).font=F(bold=True,size=10)
        for j_,sn in enumerate(ORDER):
            S8.cell(rr_,2+j_,f'=IFERROR(SUMIFS({R("M")},{R("B")},"{sn}",{R(col)},$A{rr_})/SUMIFS({R("M")},{R("B")},"{sn}"),"")').number_format='0%'
    y1=y0+len(bands)-1; body(S8,y0,y1,1,1+len(ORDER))
    S8.conditional_formatting.add(f'B{y0}:{CL(1+len(ORDER))}{y1}',ColorScaleRule(start_type='num',start_value=0,start_color='FFFFFF',end_type='max',end_color='86B6EF'))
    return y1+2
r=trend_block(r,'Quy mô lô',  'AG',LOTB)
r=trend_block(r,'SAM (phút/sp)','AN',SAMB)
hdr(S8,r,['Chỉ số']+ORDER,h=22); z0=r+1
zrows=[('SL (pcs)',lambda sn: f'=SUMIFS({R("M")},{R("B")},"{sn}")',NUM),('Δ vs mùa trước',None,'+0%;-0%;0%'),('Δ vs cùng kỳ',None,'+0%;-0%;0%'),
('SAM TB (phút/sp)',lambda sn: f'=IFERROR(SUMIFS({R("O")},{R("B")},"{sn}")*60/SUMIFS({R("M")},{R("B")},"{sn}",{R("N")},"<>"),"")','0'),
('Lô TB (pcs)',lambda sn: f'=IFERROR(SUMIFS({R("M")},{R("B")},"{sn}")/COUNTIFS({R("B")},"{sn}"),"")',NUM),
('Đơn giá GC (USD/pcs)',lambda sn: f'=IFERROR(SUMIFS({R("AE")},{R("B")},"{sn}")/SUMIFS({R("M")},{R("B")},"{sn}",{R("AE")},"<>"),"")',USD2),
('Số ngày may TB',lambda sn: f'=IFERROR(AVERAGEIFS({R("AM")},{R("B")},"{sn}"),"")','0.0')]
for k_,(lab,fn,nf) in enumerate(zrows):
    rr_=z0+k_; S8.cell(rr_,1,lab).font=F(bold=True,size=10)
    for j_,sn in enumerate(ORDER):
        c=CL(2+j_)
        if lab=='Δ vs mùa trước': v_=f'=IFERROR({c}{z0}/{CL(1+j_)}{z0}-1,"")' if j_>=1 else '–'
        elif lab=='Δ vs cùng kỳ': v_=f'=IFERROR({c}{z0}/{CL(j_)}{z0}-1,"")' if j_>=2 else '–'
        else: v_=fn(sn)
        S8.cell(rr_,2+j_,v_).number_format=nf
z1=z0+len(zrows)-1; body(S8,z0,z1,1,1+len(ORDER))
r=z1+1; NOTE(S8,r,'27FW mới bắt đầu nhận đơn nên chưa so sánh được.')
widths(S8,[30,13,13,13,12,12,12,12,12,12,12,12,12]); S8.freeze_panes='B7'; S8.sheet_view.showGridLines=False

# ---------- 4A. 受注進行表 – tự dựng từ PLANTABLE ----------
SJ=wb.create_sheet('4A. 受注進行表')
_m0=REF.to_period('M')
title(SJ,f'受注進行表 – tiến độ nhận đơn (tự dựng từ PLANTABLE, ngày tham chiếu {REF:%d/%m/%Y})',
      'Không cần file 受注進行表 riêng. Mỗi đơn PLANTABLE = 1 thanh trên line (START→END). Đã chốt (ước tính) = đã bắt đầu may hoặc NVL đã OK (生地 OK & 資材 OK trong ghi chú mới nhất); còn lại = tạm tính (暫定, chữ đỏ). Đơn vị line-ngày = 1 line thường (32 CN) × 1 ngày làm việc; line nhỏ = 0,5; công suất 30 line.')
# ---- KPI rows 4..9; A. monthly table
rA=12; section(SJ,rA,'A. Độ phủ đơn hàng theo tháng')
hdr(SJ,rA+1,['Tháng','Ngày làm việc','Công suất (line-ngày)','Đã chốt','Tạm tính (暫定)','Còn trống','% đã chốt','% gồm tạm tính','Số đơn bắt đầu','SL bắt đầu (pcs)'],h=40)
a0=rA+2
for k_,x in enumerate(JMON.itertuples()):
    rr_=a0+k_; m=x.m
    SJ.cell(rr_,1,m.start_time.to_pydatetime()).number_format='mm/yyyy'; SJ.cell(rr_,2,x.wd); SJ.cell(rr_,3,f'=30*B{rr_}')
    SJ.cell(rr_,4,round(x.conf,1)); SJ.cell(rr_,5,round(x.tot-x.conf,1)); SJ.cell(rr_,6,f'=MAX(C{rr_}-D{rr_}-E{rr_},0)')
    SJ.cell(rr_,7,f'=IFERROR(D{rr_}/C{rr_},"")'); SJ.cell(rr_,8,f'=IFERROR((D{rr_}+E{rr_})/C{rr_},"")')
    SJ.cell(rr_,9,f'=COUNTIFS({R("P")},">="&A{rr_},{R("P")},"<"&EDATE(A{rr_},1),{R("AK")},"<>Y")'); SJ.cell(rr_,10,f'=SUMIFS({R("M")},{R("P")},">="&A{rr_},{R("P")},"<"&EDATE(A{rr_},1),{R("AK")},"<>Y")')
    for c_ in (2,3,4,5,6,9,10): SJ.cell(rr_,c_).number_format='#,##0;-#,##0;-'
    for c_ in (7,8): SJ.cell(rr_,c_).number_format='0%'
a1=a0+len(JMON)-1; body(SJ,a0,a1,1,10)
SJ.conditional_formatting.add(f'H{a0}:H{a1}',FormulaRule(formula=[f'AND(ISNUMBER(H{a0}),H{a0}<0.5)'],fill=PatternFill('solid',fgColor='F6C9C9')))
SJ.conditional_formatting.add(f'G{a0}:G{a1}',FormulaRule(formula=[f'AND(ISNUMBER(G{a0}),G{a0}<0.5,H{a0}>=0.8)'],fill=PatternFill('solid',fgColor='FCE3B0')))
_ix=lambda p: a0+list(JMON.m).index(p)
bc=BarChart(); bc.type='col'; bc.grouping='stacked'; bc.overlap=100; bc.title='Độ phủ đơn hàng theo tháng (line-ngày)'; bc.y_axis.numFmt='#,##0'
bc.add_data(Reference(SJ,min_col=4,max_col=6,min_row=rA+1,max_row=a1),titles_from_data=True); bc.set_categories(Reference(SJ,min_col=1,min_row=a0,max_row=a1))
for s_,col_ in zip(bc.series,('2A78D6','EB6834','D9D9D9')): s_.graphicalProperties.solidFill=col_; s_.graphicalProperties.line.noFill=True
bc.x_axis.number_format='mm/yy'; bc.gapWidth=50; bc.height=8; bc.width=18; chart_fix(bc); SJ.add_chart(bc,'L12')
_w1=[_m0+1,_m0+3]; _w2=[_m0+4,_m0+7]
_wl=lambda w: f'T{w[0].month}/{w[0].year}–T{w[1].month}/{w[1].year}'
# ---- B. GRID 受注進行表
rB=a1+3; section(SJ,rB,'B. Bảng 受注進行表 (line × nửa tháng) – nhãn = mùa_brand chiếm nhiều nhất; chữ đỏ = tạm tính; * = line nhỏ 小ロット')
NOTE(SJ,rB+1,'Nền xanh dương = hàng SS, xanh lá = hàng FW, tím = SMS (mẫu chào hàng), trắng = trống. Ô gộp = cùng một nhóm đơn chạy liên tục.')
hdr(SJ,rB+2,['Line']+[p[2] for p in JPER],h=30); g0=rB+3
FILLS={'S':'DCE9F9','F':'D9F2D9'}
for k_,(ln,hh) in enumerate([(l,h) for l in range(1,16) for h in 'AB']):
    rr_=g0+k_; SJ.cell(rr_,1,f'{ln}{hh}').font=F(bold=True,size=10); SJ.cell(rr_,1).border=BOX
    run=None
    for j_ in range(len(JPER)+1):
        val=JGRID.get((ln,hh,j_)) if j_<len(JPER) else None
        key=(val[0],val[1]<0.5) if val else None
        if run and (key!=run[0]):
            c0_,c1_=run[1],j_-1
            if c1_>c0_: SJ.merge_cells(start_row=rr_,start_column=2+c0_,end_row=rr_,end_column=2+c1_)
            run=None
        if val and run is None: run=(key,j_)
        if j_<len(JPER):
            c=SJ.cell(rr_,2+j_); c.border=BOX; c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
            if val:
                lab,cs,sm,cov,nl=val
                c.value=lab+('*' if sm else '')+(' +' if nl>1 else '')
                fc='E4DEF7' if '_SMS' in lab else FILLS.get(lab[:1],'EDEDED')
                c.fill=PatternFill('solid',fgColor=fc); c.font=F(size=8,bold=True,color='D03B3B' if cs<0.5 else '1F1F1F')
            else:
                c.value=None
    SJ.row_dimensions[rr_].height=18
g1=g0+29
for j_ in range(len(JPER)): SJ.column_dimensions[CL(2+j_)].width=9.5
# ---- C. season x brand
rC=g1+2; section(SJ,rC,'C. Đơn theo mùa × brand – các tháng tới (line-ngày, gồm tạm tính)')
mcols=list(JSB.columns.drop(['tot','tent']))
hdr(SJ,rC+1,['Mùa / brand']+[m.strftime('%m/%Y') for m in mcols]+['Tổng','Trong đó tạm tính','% tạm tính'],h=32)
b0=rC+2; _sb=JSB[JSB.tot>=1].head(16)
for k_,(lab,x) in enumerate(_sb.iterrows()):
    rr_=b0+k_; SJ.cell(rr_,1,lab).font=F(bold=True,size=10)
    for j_,m in enumerate(mcols): SJ.cell(rr_,2+j_,round(float(x[m]),1) if x[m]>=0.5 else None).number_format='#,##0'
    n_=len(mcols); SJ.cell(rr_,2+n_,f'=SUM(B{rr_}:{CL(1+n_)}{rr_})').number_format='#,##0'; SJ.cell(rr_,3+n_,round(float(x.tent),1)).number_format='#,##0'
    SJ.cell(rr_,4+n_,f'=IFERROR({CL(3+n_)}{rr_}/{CL(2+n_)}{rr_},"")').number_format='0%'
b1=b0+max(len(_sb),1)-1; body(SJ,b0,b1,1,4+len(mcols))
SJ.conditional_formatting.add(f'B{b0}:{CL(1+len(mcols))}{b1}',ColorScaleRule(start_type='num',start_value=0,start_color='FFFFFF',end_type='max',end_color='86B6EF'))
# ---- D. per line
rD=b1+2; section(SJ,rD,'D. Từng line (32 CN): đang chạy, đã chốt đến, hết đơn')
hdr(SJ,rD+1,['Line','Đang chạy (tuần tham chiếu)','Đã chốt đến','Tạm tính đến (hết đơn)','Line trống từ','Đơn cuối cùng','Line nhỏ (小ロット)'],h=32)
c0=rD+2
for k_,x in enumerate(JLN.itertuples()):
    rr_=c0+k_
    vals=[x.line,x.cur,x.conf_end.to_pydatetime() if pd.notna(x.conf_end) else None,x.tent_end.to_pydatetime() if pd.notna(x.tent_end) else None,f'=IF(ISNUMBER(D{rr_}),D{rr_}+1,"")',x.last,x.small]
    for j_,val in enumerate(vals,1): SJ.cell(rr_,j_,val)
    for j_ in (3,4,5): SJ.cell(rr_,j_).number_format=DATE
    SJ.cell(rr_,1).font=F(bold=True,size=10)
c1=c0+len(JLN)-1; body(SJ,c0,c1,1,7)
_lim=(_w2[0].start_time+pd.Timedelta(days=14))
SJ.conditional_formatting.add(f'E{c0}:E{c1}',FormulaRule(formula=[f'AND(ISNUMBER(E{c0}),E{c0}<DATE({_lim.year},{_lim.month},{_lim.day}))'],fill=PatternFill('solid',fgColor='F6C9C9')))
# ---- E. trend
rE=c1+2; section(SJ,rE,f'E. Xu hướng qua các bản chụp (độ phủ {_wl(_w1)} và {_wl(_w2)})')
NOTE(SJ,rE+1,'Lịch sử lưu ở bao-cao/lich-su/do_phu_don_theo_ban_chup.csv. Các mốc trước 13/11/2026 lấy từ file 受注進行表; từ nay mỗi lần chạy báo cáo với PLANTABLE mới sẽ thêm 1 mốc (nguồn PLANTABLE).')
hdr(SJ,rE+2,['Ngày bản chụp','Nguồn',f'{_wl(_w1)}: gồm tạm tính',f'{_wl(_w1)}: đã chốt',f'{_wl(_w2)}: gồm tạm tính',f'{_wl(_w2)}: đã chốt'],h=45)
d0=rE+3
for k_,x in enumerate(JTR.itertuples()):
    rr_=d0+k_; SJ.cell(rr_,1,x.date.to_pydatetime()).number_format=DATE; SJ.cell(rr_,2,x.src)
    for j_,v_ in enumerate((x.t1,x.c1,x.t2,x.c2)): SJ.cell(rr_,3+j_,round(float(v_),4) if pd.notna(v_) else None).number_format='0%'
d1=d0+len(JTR)-1; body(SJ,d0,d1,1,6)
lc2=LineChart(); lc2.title='Độ phủ đơn (gồm tạm tính) theo bản chụp'; lc2.y_axis.numFmt='0%'; lc2.y_axis.scaling.min=0; lc2.y_axis.scaling.max=1
lc2.add_data(Reference(SJ,min_col=3,min_row=rE+2,max_row=d1),titles_from_data=True); lc2.add_data(Reference(SJ,min_col=5,min_row=rE+2,max_row=d1),titles_from_data=True)
lc2.set_categories(Reference(SJ,min_col=1,min_row=d0,max_row=d1)); lc2.x_axis.number_format='dd/mm'
for s_,col_ in zip(lc2.series,('2A78D6','EB6834')): s_.graphicalProperties.line.solidFill=col_; s_.graphicalProperties.line.width=22000; s_.smooth=False
lc2.height=7.5; lc2.width=18; chart_fix(lc2); SJ.add_chart(lc2,f'H{rE}')
# ---- F. brand mix by year (line-days)
rF=d1+2; section(SJ,rF,'F. Cơ cấu brand theo năm (tỷ trọng line-ngày, từ PLANTABLE)')
yrs=list(JHB.columns); hdr(SJ,rF+1,['Brand']+[str(y) for y in yrs],h=22); e0=rF+2
_hs=JHB.div(JHB.sum(),axis=1); topb=list(JHB.sum(axis=1).sort_values(ascending=False).index[:10])
for k_,b in enumerate(topb+['Khác']):
    rr_=e0+k_; SJ.cell(rr_,1,b).font=F(bold=True,size=10)
    for j_,y in enumerate(yrs):
        v_=float(_hs.loc[b,y]) if b!='Khác' else float(1-_hs.loc[topb,y].sum())
        SJ.cell(rr_,2+j_,round(v_,4)).number_format='0%'
e1=e0+len(topb); body(SJ,e0,e1,1,1+len(yrs))
SJ.conditional_formatting.add(f'B{e0}:{CL(1+len(yrs))}{e1}',ColorScaleRule(start_type='num',start_value=0,start_color='FFFFFF',end_type='max',end_color='86B6EF'))
# ---- KPI
_i1a,_i1b,_i2a,_i2b=_ix(_w1[0]),_ix(_w1[1]),_ix(_w2[0]),_ix(_w2[1])
_prev=JTR[JTR.date<=REF-pd.Timedelta(days=21)]
je=[(f'Độ phủ {_wl(_w1)} (gồm tạm tính)',f'=SUM(D{_i1a}:E{_i1b})/SUM(C{_i1a}:C{_i1b})','0%',f'=SUM(D{_i1a}:D{_i1b})/SUM(C{_i1a}:C{_i1b})','0%','riêng phần đã chốt','Kín nhưng chủ yếu tạm tính','warning'),
(f'Độ phủ {_wl(_w2)} (gồm tạm tính)',f'=SUM(D{_i2a}:E{_i2b})/SUM(C{_i2a}:C{_i2b})','0%',float(_prev.t2.iloc[-1]) if len(_prev) and pd.notna(_prev.t2.iloc[-1]) else None,'0%',f'bản chụp {(_prev.date.iloc[-1] if len(_prev) else REF):%d/%m/%Y}','Thiếu đơn','critical'),
(f'Line-ngày còn trống {_wl(_w2)}',f'=SUM(F{_i2a}:F{_i2b})',NUM,f"=SUM(F{_i2a}:F{_i2b})*32*"+CAP['prod'],NUM,'≈ pcs (theo năng suất 2026)','Cần đơn mới','critical'),
(f'Số line (32 CN) hết đơn trước {_lim:%d/%m/%Y}',f'=COUNTIFS($E${c0}:$E${c1},"<"&DATE({_lim.year},{_lim.month},{_lim.day}))','0',None,None,'trên 30 line','Ưu tiên giao đơn','serious'),
(f'Tỷ trọng tạm tính trong {_wl(_w1)}',f'=SUM(E{_i1a}:E{_i1b})/SUM(D{_i1a}:E{_i1b})','0%',None,None,'đơn chưa chốt (chữ đỏ)','Cần chốt đơn','serious')]
kpi_table(SJ,4,je,note=(4,7),stat=(8,10))
# ---- G. findings
rG=e1+2; section(SJ,rG,'G. Nhận định & khuyến nghị')
JN=dict(t1=K['j_t1'],c1=K['j_c1'],t2=K['j_t2'],nfree=int((JLN.tent_end<_lim).sum()),lim=_lim,w1=_wl(_w1),w2=_wl(_w2),pk=K['j_pk'],pkd=K['j_pkd'])
jf=[('1. 3 tháng tới kín lịch nhưng phần lớn còn tạm tính.',f'{JN["w1"]} đã kín khoảng {JN["t1"]:.0%} công suất, nhưng chỉ {JN["c1"]:.0%} là đơn đã chốt (đã chạy hoặc NVL OK). Phần còn lại là tạm tính, cần chốt sớm để mua NVL và giữ lịch.'),
('2. Các tháng sau còn thiếu đơn.',f'{JN["w2"]} kể cả tạm tính mới phủ {JN["t2"]:.0%} công suất (đỉnh {JN["pk"]:.0%} ngày {JN["pkd"]:%d/%m}). Theo dõi bảng E mỗi tuần.'),
('3. Line hết đơn.',f'{JN["nfree"]}/30 line hết cả đơn tạm tính trước {JN["lim"]:%d/%m/%Y} (bảng D, ô đỏ) – ưu tiên giao đơn mới cho các line này.'),
('4. Khuyến nghị.','(a) Sales chốt các đơn đang tạm tính trong 3 tháng tới; (b) mục tiêu độ phủ các tháng 4–7 ≥60% sau 4 tuần và ≥90% sau 8 tuần; (c) nếu độ phủ không tăng sau 2 bản chụp liên tiếp thì kích hoạt phương án nhân lực (sheet 1B).')]
text_rows(SJ,rG+1,jf,ncol=11,w=130)
SJ.column_dimensions['A'].width=34; SJ.sheet_view.showGridLines=False; SJ.freeze_panes=None

for ws_ in (S1,S2,S3,S4,S5,S6,S7,S8,SC,SG,SQ,SX,SJ):
    ws_.page_setup.orientation='landscape'; ws_.page_setup.fitToWidth=1; ws_.page_setup.fitToHeight=0; ws_.sheet_properties.pageSetUpPr.fitToPage=True
wb._sheets=[SX,S1,SC,SG,SQ,S2,S8,S3,S4,SJ,S5,S6,S7,SD,SM]
pass
wb.calculation.fullCalcOnLoad=True
wb.save(OUT); print('saved', OUT, len(fw))

# ================= JAPANESE VERSION =================
import importlib.util, os
_spec=importlib.util.spec_from_file_location('ja_tr',os.path.join(os.path.dirname(os.path.abspath(__file__)),'ja_tr.py'))
ja_tr=importlib.util.module_from_spec(_spec); _spec.loader.exec_module(ja_tr)
ja_tr.translate_wb(wb,dict(JN=JN,N_=N_,RK=RK,K=K,N_ADJ1=N_ADJ1,q8k=sev_hi.qty.sum()/1000,risk_usd=risk_usd))
OUT_JP=OUT.replace('.xlsx','_JP.xlsx').replace('Bao_cao_dieu_hanh','Keiei_Report')
for _n,_w in [('4A. 受注進行表',{'A':34}),('1. エグゼクティブサマリー',{'A':54}),('1A. CEO',{'A':50}),('1B. 工場長',{'A':50,'N':34}),('1C. 生産管理',{'A':40}),('2. シーズン別',{'A':12}),('前提・定義',{'A':40}),('7. データ品質',{'A':34})]:
    for _c,_x in _w.items(): wb[_n].column_dimensions[_c].width=_x
wb.calculation.fullCalcOnLoad=True
wb.save(OUT_JP); print('saved', OUT_JP)
