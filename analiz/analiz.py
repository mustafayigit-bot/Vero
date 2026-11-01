import sys, json, numpy as np, openpyxl
from scipy import stats
from statsmodels.stats.power import TTestPower
np.set_printoptions(suppress=True)
wb=openpyxl.load_workbook(sys.argv[1]); ws=wb.active
rows=[r for r in ws.iter_rows(min_row=2,values_only=True) if r[1] is not None]
X=np.array([[r[1],r[2],r[3]] for r in rows],float)   # isimler kullanılmaz
n,k=X.shape; names=['CBCT','Primescan','Fuzyon']
R={}
print("n=",n)
for j,nm in enumerate(names):
    x=X[:,j]; m=x.mean(); s=x.std(ddof=1); se=s/np.sqrt(n); t=stats.t.ppf(.975,n-1)
    q1,md,q3=np.percentile(x,[25,50,75]); sw=stats.shapiro(x)
    print(f"{nm}: mean={m:.4f} sd={s:.4f} CI=({m-t*se:.4f},{m+t*se:.4f}) med={md:.4f} IQR=({q1:.4f},{q3:.4f}) min={x.min():.4f} max={x.max():.4f} SW W={sw.statistic:.3f} p={sw.pvalue:.3f}")
# Friedman
fr=stats.friedmanchisquare(*X.T); W=fr.statistic/(n*(k-1))
print(f"Friedman chi2={fr.statistic:.4f} p={fr.pvalue:.3e} W={W:.4f}")
rk=np.apply_along_axis(stats.rankdata,1,X); print("mean ranks",rk.mean(0), "rank sums",rk.sum(0))
# RM-ANOVA
gm=X.mean(); ss_c=n*((X.mean(0)-gm)**2).sum(); ss_s=k*((X.mean(1)-gm)**2).sum(); ss_t=((X-gm)**2).sum(); ss_e=ss_t-ss_c-ss_s
df1,df2=k-1,(k-1)*(n-1); F=(ss_c/df1)/(ss_e/df2); pe=ss_c/(ss_c+ss_e)
# Mauchly + GG epsilon
S=np.cov(X.T); C=np.array([[1,-1,0],[0,1,-1]],float); C=np.linalg.qr(C.T)[0].T
Sc=C@S@C.T; p=k-1; eps=np.trace(Sc)**2/(p*np.trace(Sc@Sc))
Wm=np.linalg.det(Sc)/(np.trace(Sc)/p)**p; d=1-(2*p*p+p+2)/(6*p*(n-1)); chiM=-(n-1)*d*np.log(Wm); dfM=p*(p+1)/2-1
print(f"RM-ANOVA F({df1},{df2})={F:.3f} p={stats.f.sf(F,df1,df2):.3e} partial eta2={pe:.4f} f={np.sqrt(pe/(1-pe)):.3f}")
print(f"Mauchly W={Wm:.4f} chi2={chiM:.3f} df={dfM} p={stats.chi2.sf(chiM,dfM):.4f} GG eps={eps:.4f} GG p={stats.f.sf(F,df1*eps,df2*eps):.3e} df=({df1*eps:.2f},{df2*eps:.2f})")
print("corr\n",np.corrcoef(X.T))
# pairwise
pairs=[(0,1),(0,2),(1,2)]
for a,b in pairs:
    dlt=X[:,a]-X[:,b]; md=dlt.mean(); sd=dlt.std(ddof=1); dz=md/sd; tt=stats.ttest_rel(X[:,a],X[:,b]); t=stats.t.ppf(.975,n-1)
    wx=stats.wilcoxon(X[:,a],X[:,b],method='exact'); wxa=stats.wilcoxon(X[:,a],X[:,b],method='approx')
    z=stats.norm.isf(wxa.pvalue/2); r=z/np.sqrt(n)
    # Dunn-type (SPSS Friedman pairwise): z=(Ri-Rj)/sqrt(k(k+1)/(6n))
    zd=(rk[:,a].mean()-rk[:,b].mean())/np.sqrt(k*(k+1)/(6*n)); pd=2*stats.norm.sf(abs(zd))
    pw=TTestPower().power(effect_size=abs(dz),nobs=n,alpha=0.05/3) if abs(dz)<50 else 1
    print(f"{names[a]}-{names[b]}: diff={md:.4f} sd={sd:.4f} CI=({md-t*sd/np.sqrt(n):.4f},{md+t*sd/np.sqrt(n):.4f}) dz={dz:.3f} t={tt.statistic:.3f} p={tt.pvalue:.3e} SWdiff p={stats.shapiro(dlt).pvalue:.3f} | Wilcoxon T={wx.statistic} p_exact={wx.pvalue:.5f} p_bonf={min(1,3*wx.pvalue):.5f} z={z:.3f} r={r:.3f} | SPSS-Dunn z={zd:.3f} p={pd:.4f} adj={min(1,3*pd):.4f} | power(paired t, a=.0167)={pw:.4f}")
