import sys, numpy as np, openpyxl
from scipy import stats
from statsmodels.stats.power import TTestPower, TTestIndPower
wb=openpyxl.load_workbook(sys.argv[1]); ws=wb.active
X=np.array([[r[1],r[2],r[3]] for r in ws.iter_rows(min_row=2,values_only=True) if r[1] is not None],float)
n,k=X.shape; rng=np.random.default_rng(20261101); B=20000
crit=stats.chi2.ppf(.95,2); print("crit chi2",crit)
print("noncentral chi2 power (lambda=22.53):",stats.ncx2.sf(crit,2,22.5333))
def fried_stat(Y):
    rk=np.apply_along_axis(stats.rankdata,-1,Y); m=Y.shape[-2]
    Rj=rk.sum(-2); return 12/(m*k*(k+1))*(Rj**2).sum(-1)-3*m*(k+1)
def fried_p(Y): return np.array([stats.friedmanchisquare(*y.T).pvalue for y in Y])
mu=X.mean(0); S=np.cov(X.T)
for m in [3,4,5,6,8,10,12,15]:
    idx=rng.integers(0,n,(B if m<15 else B,m)); Yb=X[idx]
    pb=(fried_p(Yb[:4000])<.05).mean()
    Ym=rng.multivariate_normal(mu,S,(4000,m)); pm=(fried_p(Ym)<.05).mean()
    print(f"n={m}: bootstrap power={pb:.3f}  MVN power={pm:.3f}")
# RM-ANOVA (G*Power style, within factors, corr from data, eps=GG)
f=4.054; lam_raw=f**2*n*k  # G*Power 'as in SPSS' uses lambda = f^2 * n * k? report conservative eps
for eps in [1,0.5235]:
    df1=(k-1)*eps; df2=(k-1)*(n-1)*eps; lam=f**2*n*k*eps
    print("RM ANOVA power eps",eps, stats.ncf.sf(stats.f.ppf(.95,df1,df2),df1,df2,lam))
# Prescan vs Fusion
d=X[:,1]-X[:,2]; md=d.mean(); sd=d.std(ddof=1); se=sd/np.sqrt(n); t90=stats.t.ppf(.95,n-1)
print("PS-F 90% CI",md-t90*se,md+t90*se)
for M in [0.05,0.1,0.3]:
    p1=stats.t.sf((md+M)/se,n-1); p2=stats.t.cdf((md-M)/se,n-1)
    # TOST power with true diff=0, sd=observed, n=15 (simulation)
    sims=rng.normal(0,sd,(50000,n)); mm=sims.mean(1); ss=sims.std(1,ddof=1)/np.sqrt(n)
    pw=((mm-t90*ss>-M)&(mm+t90*ss<M)).mean()
    # n needed for 80% TOST
    nn=None
    for m in range(3,60):
        s2=rng.normal(0,sd,(20000,m)); a=s2.mean(1); b=s2.std(1,ddof=1)/np.sqrt(m); tq=stats.t.ppf(.95,m-1)
        if ((a-tq*b>-M)&(a+tq*b<M)).mean()>=.8: nn=m;break
    print(f"margin {M}: TOST p={max(p1,p2):.2e} power(n=15)={pw:.3f} n80={nn}")
tp=TTestPower()
print("n to detect observed dz=0.027 (a=.05, 80%):",tp.solve_power(effect_size=0.0272,alpha=.05,power=.8))
print("observed power PS-F a=.05:",tp.power(0.0272,n,.05))
# a priori recalcs
ip=TTestIndPower()
print("ind one-tailed d1.2:",ip.solve_power(1.2,alpha=.05,power=.8,alternative='larger'))
print("ind two-tailed d1.2:",ip.solve_power(1.2,alpha=.05,power=.8))
for a in [.05,.05/3]:
    nn=tp.solve_power(1.2,alpha=a,power=.8); print(f"paired dz1.2 a={a:.4f}: n={nn:.2f}  /0.955={nn/0.955:.2f} /0.864={nn/0.864:.2f}")
for nn in [11,12,13,15]: print("paired power n",nn,"a=.0167:",tp.power(1.2,nn,.05/3),"a=.05:",tp.power(1.2,nn,.05))
print("min dz detectable n=15 a=.0167 80%:",tp.solve_power(nobs=15,alpha=.05/3,power=.8))
print("min dz detectable n=15 a=.05 80%:",tp.solve_power(nobs=15,alpha=.05,power=.8))
