import json,openpyxl
from copy import copy
from openpyxl.styles import Alignment,Font
from openpyxl.utils import get_column_letter as L
d=json.load(open('/tmp/netdata.json'));P=d['P'];E=d['E']
wb=openpyxl.load_workbook('DDF_Pooled_Fund_Programs.xlsx')
hc=wb['Programs']['A1']
def sheet(name,hdr,rows,widths,note):
    if name in wb.sheetnames: del wb[name]
    ws=wb.create_sheet(name)
    ws.append([note]); ws['A1'].font=Font(name='Arial',italic=True,size=10)
    ws.append(hdr)
    for c in ws[2]:
        c.font=copy(hc.font);c.fill=copy(hc.fill);c.alignment=Alignment(wrap_text=True,vertical='center')
    for r in rows: ws.append(r)
    for r in ws.iter_rows(min_row=3):
        for c in r: c.font=Font(name='Arial',size=10);c.alignment=Alignment(wrap_text=True,vertical='top')
    for i,w in enumerate(widths): ws.column_dimensions[L(i+1)].width=w
    ws.freeze_panes='A3'; ws.auto_filter.ref=f"A2:{L(len(hdr))}{ws.max_row}"
pairs={}
for p in P:
    f=p['f']
    for i in range(len(f)):
        for j in range(i+1,len(f)):
            k=tuple(sorted((f[i],f[j])))
            e=pairs.setdefault(k,{'progs':[],'w':0,'known':False}); e['progs'].append(p)
            if p['amt'] is not None: e['w']+=p['amt']; e['known']=True
rows=[]
for (a,b),e in pairs.items():
    ddf='DDF' in (a,b)
    if ddf: a,b='DDF',(b if a=='DDF' else a)
    ps=e['progs']
    rows.append([a,b,'DDF to funder' if ddf else 'Co-funder (same program, no DDF)',len(ps),
      '; '.join(sorted({p['pr'] for p in ps})),'; '.join(p['n'] for p in ps),
      '; '.join(str(y) for y in sorted({p['y'] for p in ps})),
      round(e['w'],1) if e['known'] else None,'Yes' if e['known'] else 'No: amount not found'])
rows.sort(key=lambda r:(r[2]!='DDF to funder',r[1] if r[2]=='DDF to funder' else r[0],r[1]))
sheet('Network edges',['Funder A','Funder B','Link type','Shared programs (count)','DDF program(s)','Shared programs','Years','Edge size: total size of shared programs ($M)','Amount found?'],rows,[22,22,26,12,22,60,16,18,20],
 "Edge list behind the network visualization. Edge size = sum of the total sizes of programs both funders share. It is not either funder's own contribution, which is mostly undisclosed.")
nodes={}
for p in P:
    for f in p['f']: nodes.setdefault(f,[]).append(p)
nr=[]
for f,ps in nodes.items():
    en=E.get(f)
    nr.append([f,en[0] if en else None,en[1] if en else 'No figure found',len(ps),'; '.join(sorted({p['pr'] for p in ps})),'; '.join(p['n'] for p in ps)])
nr.sort(key=lambda r:(r[0]!='DDF',-(r[1] or -1),r[0]))
sheet('Network nodes',['Funder','Endowment ($B)','Endowment basis','Programs (count)','DDF program(s)','Programs'],nr,[24,14,28,12,24,70],
 'Node list behind the network visualization. Node size = endowment. Faint nodes in the visualization have no figure.')
pr=[[p['n'],p['pr'],p['y'],p['amt'],p.get('basis'),p.get('src'),len(p['f']),'; '.join(p['f'])] for p in P]
sheet('Network programs',['Program','DDF program','Year','Total size ($M)','What the amount is','Source','Funders (count)','Funders'],pr,[44,18,8,14,44,50,12,90],
 'Programs used to build the network. The year filter uses the Year column. A blank size means the amount was not found.')
wb.save('DDF_Pooled_Fund_Programs.xlsx')
print(len(rows),len(nr),len(pr))
for r in rows[:3]+rows[-2:]: print(r[:5],r[7])
