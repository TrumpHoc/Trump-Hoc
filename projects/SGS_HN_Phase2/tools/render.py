import json, math, sys
import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt
src, out, x0,y0,x1,y1 = sys.argv[1], sys.argv[2], *map(float, sys.argv[3:7])
d=json.load(open(src)); B=d['blocks']
fig,ax=plt.subplots(figsize=(16,16*(y1-y0)/(x1-x0)))
def T(p,tr):
    x,y=p
    for (ix,iy,sx,sy,r) in tr:
        x,y=x*sx,y*sy; c,s=math.cos(r),math.sin(r); x,y=x*c-y*s+ix, x*s+y*c+iy
    return x,y
def draw(ents,tr,depth=0):
    for e in ents:
        t=e['type']
        try:
            if t=='LINE':
                a=T((e['startPoint']['x'],e['startPoint']['y']),tr); b=T((e['endPoint']['x'],e['endPoint']['y']),tr); ax.plot([a[0],b[0]],[a[1],b[1]],'k-',lw=0.4)
            elif t=='LWPOLYLINE':
                v=[T((p['x'],p['y']),tr) for p in e['vertices']]
                if e.get('flag',0)&1: v.append(v[0])
                ax.plot([p[0] for p in v],[p[1] for p in v],'k-',lw=0.4)
            elif t in('CIRCLE','ARC'):
                c=e['center']; r=e['radius']; a0=e.get('startAngle',0); a1=e.get('endAngle',2*math.pi)
                if a1<a0: a1+=2*math.pi
                pts=[T((c['x']+r*math.cos(a0+(a1-a0)*i/40), c['y']+r*math.sin(a0+(a1-a0)*i/40)),tr) for i in range(41)]
                ax.plot([p[0] for p in pts],[p[1] for p in pts],'k-',lw=0.4)
            elif t in('TEXT','MTEXT'):
                p=e.get('insertionPoint') or e.get('startPoint') or {}
                s=str(e.get('text') or e.get('contents') or '').replace('\\P','\n')
                import re; s=re.sub(r'\\[A-Za-z][^;]*;','',s).replace('{','').replace('}','')
                if p: q=T((p['x'],p['y']),tr); ax.text(q[0],q[1],s,fontsize=6,color='blue',va='top')
            elif t=='INSERT' and depth<3 and e['name'] in B:
                ip=e['insertionPoint']
                draw(B[e['name']],[(ip['x'],ip['y'],e.get('xScale',1),e.get('yScale',1),e.get('rotation',0))]+tr,depth+1)
        except Exception as ex: pass
draw(d['ents'],[])
ax.set_xlim(x0,x1); ax.set_ylim(y0,y1); ax.set_aspect('equal'); ax.axis('off')
plt.savefig(out,dpi=110,bbox_inches='tight')
