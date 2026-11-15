
# ---------- 0A. Thuật ngữ & định nghĩa ----------
_tn_spec=importlib.util.spec_from_file_location('thuat_ngu',os.path.join(_HERE,'thuat_ngu.py')); TN=importlib.util.module_from_spec(_tn_spec); _tn_spec.loader.exec_module(TN)
SGl=wb.create_sheet('0A. Thuật ngữ')
title(SGl,'Thuật ngữ & định nghĩa','Giải thích các khái niệm dùng trong báo cáo. Ô chữ xanh dương ở phần B có thể sửa để thử.')
r=4; section(SGl,r,'A. Chi phí chia lô là gì?'); r+=1
for vi_,ja_ in TN.GIAI_THICH:
    c=SGl.cell(r,1,vi_); c.font=F(size=10); c.alignment=Alignment(wrap_text=True,vertical='top'); SGl.merge_cells(start_row=r,start_column=1,end_row=r,end_column=12); SGl.row_dimensions[r].height=32; r+=1
r+=1; section(SGl,r,'B. Ví dụ tính theo công thức giá vốn của PLANTABLE'); r+=1
NOTE(SGl,r,'Một đơn hàng được may thành 1 lô lớn hoặc chia nhỏ thành nhiều lô. Mỗi lô phải trả lại chi phí học việc và chạy ở hiệu suất thấp hơn, nên tổng chi phí tăng.'); r+=1
inp0=r
for lab,val,fmt,font in [('SAM (giây/sp)',10000,'#,##0',BLUE),('Số công nhân / line',32,'0',BLUE),('Chi phí nhân công / công-ngày (USD)',"='Phương pháp'!$C$9",USD2,GREEN),('Số lượng đơn hàng (pcs)',5000,'#,##0',BLUE)]:
    SGl.cell(r,1,lab).font=F(bold=True,size=10); c=SGl.cell(r,3,val); c.number_format=fmt; c.font=F(size=10,bold=True,color=font.color.rgb if font.color else None); c.border=BOX; r+=1
SAMc,CNc,RTc,ORc=f'$C${inp0}',f'$C${inp0+1}',f'$C${inp0+2}',f'$C${inp0+3}'
r+=1
hdr(SGl,r,['Phương án','Lô SX (pcs)','Số lô','稼働率 (hệ số hiệu suất)','習熟日数 (ngày học việc)','日産MAX (pcs/ngày)','Số ngày may / lô','Chi phí / lô (USD)','Đơn giá GC (USD/pcs)','Tổng chi phí (USD)','Chi phí chia lô (USD)','% tăng so với 1 lô'],h=45)
t0=r+1
for k_,(lab,lot) in enumerate([('1 lô (không chia)',f'={ORc}/1'),('Chia 2 lô',f'={ORc}/2'),('Chia 5 lô',f'={ORc}/5'),('Chia 10 lô',f'={ORc}/10'),('Lô 300 pcs',300)]):
    rr_=t0+k_
    SGl.cell(rr_,1,lab).font=F(bold=True,size=10)
    c=SGl.cell(rr_,2,lot); c.font=BLUE
    SGl.cell(rr_,3,f'={ORc}/B{rr_}')
    SGl.cell(rr_,4,f'=(3.51*LN(B{rr_})+76.97)/100')
    SGl.cell(rr_,5,f'=2.23*{SAMc}^0.11')
    SGl.cell(rr_,6,f'=ROUNDDOWN(28800/{SAMc}*{CNc}*D{rr_},0)')
    SGl.cell(rr_,7,f'=B{rr_}/F{rr_}+E{rr_}/2')
    SGl.cell(rr_,8,f'={RTc}*{CNc}*G{rr_}')
    SGl.cell(rr_,9,f'=H{rr_}/B{rr_}')
    SGl.cell(rr_,10,f'=I{rr_}*{ORc}')
    SGl.cell(rr_,11,f'=J{rr_}-$J${t0}')
    SGl.cell(rr_,12,f'=IFERROR(J{rr_}/$J${t0}-1,"")')
    for c_,fm in [(2,'#,##0'),(3,'0.0'),(4,'0.000'),(5,'0.0'),(6,'#,##0'),(7,'0.0'),(8,USD),(9,USD2),(10,USD),(11,USD),(12,'0%')]: SGl.cell(rr_,c_).number_format=fm
t1=t0+4; body(SGl,t0,t1,1,12)
SGl.conditional_formatting.add(f'K{t0}:K{t1}',ColorScaleRule(start_type='num',start_value=0,start_color='FFFFFF',end_type='max',end_color='EB6834'))
r=t1+1; NOTE(SGl,r,'Công thức giống cột AR–AY / BD–BK của PLANTABLE: 稼働率 = (3,51×LN(lô)+76,97)/100; 習熟日数 = 2,23×SAM^0,11; 日産MAX = 28.800/SAM × CN × 稼働率; số ngày = lô/日産MAX + 習熟日数/2.'); r+=2
# ví dụ thực tế
_cs=[s for s in ORDER if len(o[o.season==s])>=100][-1]
_ex=o[(o.season==_cs)&(o.c1_qty>o.c2_qty)&o.c1_upc.notna()&o.c2_upc.notna()].copy(); _ex['sp']=_ex.qty*(_ex.c2_upc-_ex.c1_upc)
_x=_ex.sort_values('sp',ascending=False).iloc[0]
section(SGl,r,'C. Ví dụ thực tế từ PLANTABLE (đơn có chi phí chia lô lớn nhất mùa hiện tại)'); r+=1
x0=r
for lab,val,fmt in [('Mùa',_x.season,None),('Brand',_x.brand,None),('Style',_x['style'],None),('Item',_x['item'],None),('Dòng gốc',int(_x.row),'0'),('SL dòng đơn (pcs)',float(_x.qty),'#,##0'),
                    ('Lô đặt hàng – giá vốn 1 (pcs)',float(_x.c1_qty),'#,##0'),('Lô sản xuất – giá vốn 2 (pcs)',float(_x.c2_qty),'#,##0'),('SAM (giây/sp)',float(_x.sam),'#,##0'),
                    ('Đơn giá GC1 (USD/pcs)',float(_x.c1_upc),USD2),('Đơn giá GC2 (USD/pcs)',float(_x.c2_upc),USD2)]:
    SGl.cell(r,1,lab).font=F(bold=True,size=10); c=SGl.cell(r,3,val); c.font=F(size=10); c.border=BOX
    if fmt: c.number_format=fmt
    r+=1
SGl.cell(r,1,'Chi phí chia lô của dòng này (USD)').font=F(bold=True,size=10); SGl.cell(r,3,f'=C{x0+5}*(C{x0+10}-C{x0+9})').number_format=USD; SGl.cell(r,3).font=F(bold=True,size=10,color='D03B3B'); r+=1
SGl.cell(r,1,'% tăng đơn giá').font=F(bold=True,size=10); SGl.cell(r,3,f'=C{x0+10}/C{x0+9}-1').number_format='0%'; r+=2
# bảng thuật ngữ
section(SGl,r,'D. Bảng thuật ngữ'); r+=1
hdr(SGl,r,['Thuật ngữ','Tiếng Nhật','Định nghĩa','Cách tính / nguồn','Xem ở sheet'],h=24); g0=r+1
for k_,(tv,tj,dv,dj,fv,fj,sh) in enumerate(TN.THUAT_NGU):
    rr_=g0+k_
    for j_,val in enumerate([tv,tj,dv,fv,sh],1):
        c=SGl.cell(rr_,j_,val); c.alignment=Alignment(wrap_text=True,vertical='top'); c.font=F(size=10,bold=(j_==1)); c.border=BOX
    SGl.row_dimensions[rr_].height=max(30,14*(max(len(dv),len(fv)*0.8)//60+1))
GL_ROWS=(g0,g0+len(TN.THUAT_NGU)-1)
for j_,w in enumerate([34,26,60,60,12],1): SGl.column_dimensions[CL(j_)].width=w
for j_ in range(6,13): SGl.column_dimensions[CL(j_)].width=13
SGl.sheet_view.showGridLines=False
