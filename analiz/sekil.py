import sys, numpy as np, openpyxl, matplotlib
matplotlib.use('Agg'); import matplotlib.pyplot as plt
wb=openpyxl.load_workbook(sys.argv[1]); ws=wb.active
X=np.array([[r[1],r[2],r[3]] for r in ws.iter_rows(min_row=2,values_only=True) if r[1] is not None],float)
out=sys.argv[2]
plt.rcParams.update({'font.size':10,'font.family':'DejaVu Sans'})
fig,ax=plt.subplots(1,2,figsize=(10,4.2),dpi=200)
labs=['KIBT –\nReferans','Primescan –\nReferans','Füzyon –\nReferans']
cols=['#B5651D','#2E6FAC','#3A8F5C']
bp=ax[0].boxplot([X[:,j] for j in range(3)],widths=.5,patch_artist=True,medianprops=dict(color='k'))
for b,c in zip(bp['boxes'],cols): b.set_facecolor(c); b.set_alpha(.35)
for i in range(15): ax[0].plot([1,2,3],X[i],color='grey',lw=.6,alpha=.6,zorder=1)
for j in range(3): ax[0].scatter(np.full(15,j+1),X[:,j],s=14,color=cols[j],zorder=3)
ax[0].set_xticks([1,2,3],labs); ax[0].set_ylabel('Yüzey sapması (mm)'); ax[0].set_title('a) Yöntemlere göre sapma değerleri (n = 15)',fontsize=10)
ax[0].spines[['top','right']].set_visible(False)
m=(X[:,1]+X[:,2])/2; d=X[:,1]-X[:,2]; md=d.mean(); sd=d.std(ddof=1)
ax[1].scatter(m,d,s=18,color='#2E6FAC')
for y,ls,t in [(md,'-',f'Ort. fark = {md:.3f}'),(md+1.96*sd,'--',f'+1.96 SS = {md+1.96*sd:.3f}'),(md-1.96*sd,'--',f'−1.96 SS = {md-1.96*sd:.3f}')]:
    ax[1].axhline(y,ls=ls,color='k',lw=.8); ax[1].text(m.max()+0.012,y+0.006,t,va='bottom',fontsize=8)
ax[1].axhline(0,color='grey',lw=.5,ls=':')
ax[1].set_xlim(m.min()-0.02,m.max()+0.17); ax[1].set_ylim(md-1.96*sd-0.03,md+1.96*sd+0.05)
ax[1].set_xlabel('İki yöntemin ortalaması (mm)'); ax[1].set_ylabel('Primescan − Füzyon (mm)')
ax[1].set_title('b) Bland–Altman: Primescan ve Füzyon model',fontsize=10); ax[1].spines[['top','right']].set_visible(False)
fig.tight_layout(); fig.savefig(out)
