"""Nihai istatistiksel analiz (n=15). Kullanım: python3 analiz_nihai.py veri.xlsx cikti_klasoru"""
import sys, json, os, numpy as np, openpyxl, matplotlib
matplotlib.use('Agg'); import matplotlib.pyplot as plt
from scipy import stats
from statsmodels.stats.power import TTestPower

src, out = sys.argv[1], sys.argv[2]; fd = os.path.join(out, 'sekiller'); os.makedirs(fd, exist_ok=True)
ws = openpyxl.load_workbook(src).active
X = np.array([[r[1], r[2], r[3]] for r in ws.iter_rows(min_row=2, values_only=True) if r[1] is not None], float)
n, k = X.shape; ids = [f'K{i+1}' for i in range(n)]
NM = ['KIBT', 'Primescan', 'Füzyon']; NML = ['KIBT –\nReferans', 'Primescan –\nReferans', 'Füzyon –\nReferans']
COL = ['#C0632B', '#2F6DB5', '#2E8B57']
rng = np.random.default_rng(20261101)
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 10, 'axes.spines.top': False, 'axes.spines.right': False,
                     'axes.titlesize': 10.5, 'axes.titleweight': 'bold', 'savefig.dpi': 220, 'savefig.bbox': 'tight'})
R = {'n': n}
def save(f, name): f.savefig(os.path.join(fd, name)); plt.close(f)

# ---------- tanımlayıcı ----------
desc = []
for j in range(k):
    x = X[:, j]; m, s = x.mean(), x.std(ddof=1); h = stats.t.ppf(.975, n-1)*s/np.sqrt(n)
    q1, md, q3 = np.percentile(x, [25, 50, 75]); sw = stats.shapiro(x)
    desc.append(dict(ad=NM[j], ort=m, ss=s, alt=m-h, ust=m+h, med=md, q1=q1, q3=q3, min=x.min(), max=x.max(), sw_w=sw.statistic, sw_p=sw.pvalue))
R['desc'] = desc

# Şekil 1: ortalama ± %95 GA
f, ax = plt.subplots(figsize=(6.2, 3.4))
for j, d in enumerate(desc):
    ax.bar(j, d['ort'], color=COL[j], alpha=.8, width=.6)
    ax.errorbar(j, d['ort'], yerr=[[d['ort']-d['alt']], [d['ust']-d['ort']]], color='k', capsize=6, lw=1.2)
    ax.text(j, d['ust']+0.06, f"{d['ort']:.3f} ± {d['ss']:.3f}", ha='center', fontsize=9)
ax.set_xticks(range(3), NML); ax.set_ylabel('Ortalama sapma (mm)'); ax.set_ylim(0, 2.6)
ax.set_title('Yöntemlere göre ortalama sapma (± %95 GA)'); save(f, 's01_ortalama.png')

# Şekil 2: kutu + noktalar + 0.3 mm çizgisi
f, ax = plt.subplots(figsize=(6.2, 3.8))
bp = ax.boxplot([X[:, j] for j in range(3)], positions=range(3), widths=.5, patch_artist=True, showfliers=False, medianprops=dict(color='k', lw=1.5))
for b, c in zip(bp['boxes'], COL): b.set_facecolor(c); b.set_alpha(.3)
for j in range(3): ax.scatter(j + rng.uniform(-.08, .08, n), X[:, j], color=COL[j], s=22, zorder=3, edgecolor='white', lw=.5)
ax.set_xticks(range(3), NML); ax.set_ylabel('Sapma (mm)'); ax.set_title('Sapma değerlerinin dağılımı (kutu grafiği ve bireysel değerler)')
save(f, 's02_kutu.png')

# Şekil 3: katılımcı bazlı gruplanmış çubuk
f, ax = plt.subplots(figsize=(6.6, 3.4)); w = .27; xx = np.arange(n)
for j in range(3): ax.bar(xx + (j-1)*w, X[:, j], w, color=COL[j], label=NM[j], alpha=.85)
ax.set_xticks(xx, ids, fontsize=8); ax.set_ylabel('Sapma (mm)'); ax.legend(frameon=False, ncol=3, loc='upper center', bbox_to_anchor=(.5, 1.12), fontsize=9)
ax.set_title('Katılımcılara göre sapma değerleri', pad=22); save(f, 's03_katilimci.png')

# Şekil 4: profil (spaghetti) + log ölçek
f, axs = plt.subplots(1, 2, figsize=(6.6, 3.3))
for i in range(n): axs[0].plot(range(3), X[i], color='grey', lw=.8, alpha=.6, marker='o', ms=3)
axs[0].plot(range(3), X.mean(0), color='k', lw=2.2, marker='o', label='Ortalama'); axs[0].legend(frameon=False, fontsize=8)
axs[0].set_xticks(range(3), NML, fontsize=8); axs[0].set_ylabel('Sapma (mm)'); axs[0].set_title('a) Katılımcı profilleri')
for i in range(n): axs[1].plot([0, 1], X[i, 1:], color='grey', lw=.8, alpha=.7, marker='o', ms=3)
axs[1].plot([0, 1], X[:, 1:].mean(0), color='k', lw=2.2, marker='o')
axs[1].set_xticks([0, 1], NML[1:], fontsize=8); axs[1].set_xlim(-.3, 1.3); axs[1].set_title('b) Primescan ve Füzyon (yakın görünüm)')
f.tight_layout(); save(f, 's04_profil.png')

# Şekil 5: normallik Q-Q
f, axs = plt.subplots(1, 3, figsize=(6.8, 2.6))
for j in range(3):
    (osm, osr), (sl, ic, _) = stats.probplot(X[:, j])
    axs[j].scatter(osm, osr, color=COL[j], s=16); axs[j].plot(osm, sl*osm+ic, color='k', lw=1)
    axs[j].set_title(f"{NM[j]}\nShapiro–Wilk p = {desc[j]['sw_p']:.3f}", fontsize=9); axs[j].set_xlabel('Teorik kantil', fontsize=8)
axs[0].set_ylabel('Gözlenen (mm)', fontsize=8); f.tight_layout(); save(f, 's05_qq.png')

# ---------- Friedman ----------
rk = np.apply_along_axis(stats.rankdata, 1, X); fr = stats.friedmanchisquare(*X.T); Wk = fr.statistic/(n*(k-1))
R['friedman'] = dict(chi2=fr.statistic, p=fr.pvalue, W=Wk, ort_sira=list(rk.mean(0)), sira_top=list(rk.sum(0)))
f, ax = plt.subplots(figsize=(6.2, 3.0))
ax.barh(range(3)[::-1], rk.mean(0), color=COL, alpha=.85, height=.55)
for j in range(3): ax.text(rk.mean(0)[j]+.04, 2-j, f'{rk.mean(0)[j]:.2f}', va='center')
ax.set_yticks(range(3)[::-1], [s.replace('\n', ' ') for s in NML]); ax.set_xlim(0, 3.4); ax.set_xlabel('Ortalama sıra (1 = en düşük sapma)')
ax.set_title(f"Friedman ortalama sıraları — χ²(2) = {fr.statistic:.2f}; p < 0.001; W = {Wk:.3f}"); save(f, 's06_friedman.png')

# ---------- ikili ----------
def hl(d):
    wa = (d[:, None]+d[None, :])/2; return np.median(wa[np.triu_indices(len(d))])
pairs = [(0, 1), (0, 2), (1, 2)]; pw = []
for a, b in pairs:
    d = X[:, a]-X[:, b]; md, sd = d.mean(), d.std(ddof=1); h = stats.t.ppf(.975, n-1)*sd/np.sqrt(n)
    wx = stats.wilcoxon(X[:, a], X[:, b], method='exact'); z = stats.norm.isf(stats.wilcoxon(X[:, a], X[:, b], method='approx').pvalue/2)
    pw.append(dict(ad=f'{NM[a]} – {NM[b]}', ort=md, ss=sd, alt=md-h, ust=md+h, med=float(np.median(d)), hl=hl(d), T=float(wx.statistic), p=wx.pvalue,
                   p_bonf=min(1, 3*wx.pvalue), z=z, r=z/np.sqrt(n), dz=md/sd, ust_say=int((d > 0).sum())))
R['ikili'] = pw
f, ax = plt.subplots(figsize=(6.4, 2.8))
for i, p in enumerate(pw):
    y = 2-i; ax.errorbar(p['ort'], y, xerr=[[p['ort']-p['alt']], [p['ust']-p['ort']]], fmt='o', color='k', capsize=5, ms=6)
    lab = 'p < 0.001' if p['p_bonf'] < .001 else f"p = {p['p_bonf']:.3f}"
    ax.text(p['ust']+0.05, y, f"{p['ort']:.3f} mm; {lab}", va='center', fontsize=9)
ax.axvline(0, color='grey', ls='--', lw=.8); ax.set_yticks([2, 1, 0], [p['ad'] for p in pw]); ax.set_xlim(-0.2, 2.6)
ax.set_xlabel('Ortalama fark (mm) ve %95 GA'); ax.set_title('İkili karşılaştırmalar (Wilcoxon, Bonferroni düzeltmeli p)'); save(f, 's07_ikili.png')

# ---------- Primescan – Füzyon: uyum ----------
a, b = X[:, 1], X[:, 2]; d = a-b; md, sd = d.mean(), d.std(ddof=1)
sp = stats.spearmanr(a, b); pr = stats.pearsonr(a, b)
# ICC(A,1) iki yönlü rastgele, mutlak uyum
Y = X[:, 1:]; gm = Y.mean(); msr = 2*((Y.mean(1)-gm)**2).sum()/(n-1); msc = n*((Y.mean(0)-gm)**2).sum()/1
sse = ((Y-gm)**2).sum()-msr*(n-1)-msc; mse = sse/(n-1)
icc = (msr-mse)/(msr+mse+2*(msc-mse)/n)
# ICC GA (McGraw & Wong, A,1)
Fr = msr/mse; a_ = 2*icc/(n*(1-icc)); b_ = 1+2*icc*(n-1)/(n*(1-icc))
v = (a_*msc+b_*mse)**2/((a_*msc)**2/1+(b_*mse)**2/(n-1))
FL = stats.f.ppf(.975, n-1, v); FU = stats.f.ppf(.975, v, n-1)
import pandas as pd, pingouin as pg  # ICC(A,1) güven aralığı pingouin ile (McGraw & Wong)
_df = pd.DataFrame([(i, m, Y[i, j]) for i in range(n) for j, m in enumerate(['P', 'F'])], columns=['id', 'm', 'v'])
_ic = pg.intraclass_corr(_df, 'id', 'm', 'v').set_index('Type').loc['ICC(A,1)']
icc = float(_ic['ICC']); iccL, iccU = map(float, _ic['CI95'])
loa = (md-1.96*sd, md+1.96*sd); se_loa = sd*np.sqrt(3/n); tq = stats.t.ppf(.975, n-1)
t90 = stats.t.ppf(.95, n-1); ci90 = (md-t90*sd/np.sqrt(n), md+t90*sd/np.sqrt(n))
tost = {}
for M in [0.05, 0.10, 0.30]:
    se = sd/np.sqrt(n); p1 = stats.t.sf((md+M)/se, n-1); p2 = stats.t.cdf((md-M)/se, n-1); tost[str(M)] = max(p1, p2)
R['uyum'] = dict(ort=md, ss=sd, loa=loa, loa_ga=[loa[0]-tq*se_loa, loa[0]+tq*se_loa, loa[1]-tq*se_loa, loa[1]+tq*se_loa],
                 spearman=sp.statistic, spearman_p=sp.pvalue, pearson=pr.statistic, icc=icc, icc_alt=iccL, icc_ust=iccU,
                 ci90=ci90, tost=tost, ba_trend_p=stats.linregress((a+b)/2, d).pvalue, ust_say=int((d > 0).sum()))

f, ax = plt.subplots(figsize=(4.4, 4.0))
lim = (0.15, 0.68); ax.plot(lim, lim, color='grey', ls='--', lw=1, label='Özdeşlik doğrusu (y = x)')
ax.scatter(a, b, color='#5B4C9A', s=30, zorder=3)
for i in range(n): ax.annotate(ids[i], (a[i], b[i]), fontsize=6.5, xytext=(3, 3), textcoords='offset points', color='grey')
ax.set_xlim(lim); ax.set_ylim(lim); ax.set_aspect('equal'); ax.set_xlabel('Primescan – Referans (mm)'); ax.set_ylabel('Füzyon – Referans (mm)')
ax.legend(frameon=False, fontsize=8, loc='upper left'); ax.set_title(f'Primescan ve Füzyon sapmalarının ilişkisi\nICC = {icc:.2f}; Spearman ρ = {sp.statistic:.2f}', fontsize=9.5)
save(f, 's08_sacilim.png')

f, ax = plt.subplots(figsize=(6.2, 3.8)); m2 = (a+b)/2
ax.scatter(m2, d, color='#5B4C9A', s=28, zorder=3)
for y, ls, t in [(md, '-', f'Ortalama fark = {md:.3f}'), (loa[1], '--', f'Üst uyum sınırı = {loa[1]:.3f}'), (loa[0], '--', f'Alt uyum sınırı = {loa[0]:.3f}')]:
    ax.axhline(y, ls=ls, color='k', lw=.9); ax.text(0.665, y+0.007, t, fontsize=8, ha='right')
ax.axhspan(-0.10, 0.10, color='#2E8B57', alpha=.08, label='±0.10 mm eşdeğerlik sınırı'); ax.axhline(0, color='grey', lw=.5, ls=':')
ax.set_xlim(0.2, 0.67); ax.set_ylim(-0.23, 0.25); ax.legend(frameon=False, fontsize=8, loc='lower left')
ax.set_xlabel('İki yöntemin ortalaması (mm)'); ax.set_ylabel('Primescan − Füzyon (mm)'); ax.set_title('Bland–Altman analizi: Primescan ve Füzyon model')
save(f, 's09_bland_altman.png')

f, ax = plt.subplots(figsize=(6.2, 2.6))
for M, c, yy in [(0.30, '#DDEBDD', 1.47), (0.10, '#A9D3B0', 1.47), (0.05, '#6DB57A', 1.30)]:
    ax.axvspan(-M, M, color=c, zorder=0); ax.text(M - 0.004, yy, f'±{M:.2f} mm', ha='right', fontsize=7.5)
ax.errorbar(md, 1, xerr=[[md-ci90[0]], [ci90[1]-md]], fmt='o', color='k', capsize=6, ms=7, lw=1.5)
ax.text(md, 0.72, f'Fark = {md:.3f} mm\n%90 GA: {ci90[0]:.3f} ile {ci90[1]:.3f}', ha='center', fontsize=8.5)
ax.axvline(0, color='grey', lw=.6, ls=':'); ax.set_xlim(-0.33, 0.33); ax.set_ylim(0.4, 1.55); ax.set_yticks([])
ax.spines['left'].set_visible(False); ax.set_xlabel('Primescan − Füzyon ortalama farkı (mm)')
ax.set_title('Eşdeğerlik testi (TOST): farkın %90 GA’sı ve eşdeğerlik sınırları'); save(f, 's10_tost.png')

# ---------- güç ----------
tp = TTestPower(); ns = np.arange(4, 31)
pw_b = [tp.power(1.2, m, .05/3) for m in ns]; pw_a = [tp.power(1.2, m, .05) for m in ns]
def tost_power(M, m, B=20000):
    s = rng.normal(0, sd, (B, m)); mm = s.mean(1); se = s.std(1, ddof=1)/np.sqrt(m); q = stats.t.ppf(.95, m-1)
    return ((mm-q*se > -M) & (mm+q*se < M)).mean()
tost_c = {M: [tost_power(M, m) for m in ns] for M in [0.05, 0.10]}
# Friedman simülasyon gücü
def fr_power(m, B=4000):
    idx = rng.integers(0, n, (B, m)); return np.mean([stats.friedmanchisquare(*X[i].T).pvalue < .05 for i in idx])
fr_ns = [3, 4, 5, 6, 8, 10, 12, 15]; fr_c = [fr_power(m) for m in fr_ns]
es = np.linspace(0.1, 2.0, 100); sens = [float(np.nan_to_num(tp.power(e, n, .05/3), nan=1.0)) for e in es]
R['guc'] = dict(nc_chi2=stats.ncx2.sf(stats.chi2.ppf(.95, 2), 2, fr.statistic), fried_sim=dict(zip(map(str, fr_ns), fr_c)),
               apriori_n=tp.solve_power(1.2, alpha=.05/3, power=.8), apriori_n_np=tp.solve_power(1.2, alpha=.05/3, power=.8)/0.864,
               pow15_bonf=tp.power(1.2, 15, .05/3), pow15=tp.power(1.2, 15, .05), mde15=tp.solve_power(nobs=15, alpha=.05/3, power=.8),
               tost15={'0.05': tost_power(0.05, 15, 50000), '0.1': tost_power(0.10, 15, 50000)},
               ikili15=[float(np.nan_to_num(tp.power(min(abs(p['dz']), 3), 15, .05/3), nan=1.0)) for p in pw])

f, axs = plt.subplots(1, 2, figsize=(6.8, 3.2))
axs[0].plot(ns, pw_b, color='#2F6DB5', lw=2, label='α = 0.0167 (Bonferroni)'); axs[0].plot(ns, pw_a, color='#2F6DB5', lw=1.2, ls='--', label='α = 0.05')
axs[0].axhline(.8, color='grey', ls=':', lw=1); axs[0].axvline(15, color='#C0632B', lw=1.2)
axs[0].scatter([15], [tp.power(1.2, 15, .05/3)], color='#C0632B', zorder=4); axs[0].text(15.5, .55, f'n = 15\ngüç = %{100*tp.power(1.2,15,.05/3):.0f}', color='#C0632B', fontsize=8.5)
axs[0].set_xlabel('Katılımcı sayısı (n)'); axs[0].set_ylabel('İstatistiksel güç'); axs[0].set_ylim(0, 1.02); axs[0].legend(frameon=False, fontsize=7.5, loc='lower right')
axs[0].set_title('a) A priori: eşleştirilmiş test, d = 1.2', fontsize=9.5)
axs[1].plot(es, sens, color='#5B4C9A', lw=2); axs[1].axhline(.8, color='grey', ls=':', lw=1)
mde = tp.solve_power(nobs=15, alpha=.05/3, power=.8); axs[1].axvline(mde, color='grey', lw=.8, ls='--'); axs[1].text(mde+.04, .3, f'En küçük\nsaptanabilir\nd = {mde:.2f}', fontsize=8)
axs[1].axvline(1.2, color='#2F6DB5', lw=.8); axs[1].text(1.23, .08, 'Planlanan\nd = 1.2', fontsize=8, color='#2F6DB5')
axs[1].set_xlabel('Etki büyüklüğü (Cohen d_z)'); axs[1].set_ylim(0, 1.02); axs[1].set_title('b) n = 15 için duyarlılık eğrisi', fontsize=9.5)
f.tight_layout(); save(f, 's11_guc_apriori.png')

f, axs = plt.subplots(1, 2, figsize=(6.8, 3.0))
axs[0].plot(fr_ns, fr_c, color='#C0632B', lw=2, marker='o'); axs[0].axhline(.8, color='grey', ls=':', lw=1); axs[0].axvline(15, color='k', lw=.8)
axs[0].set_xlabel('Katılımcı sayısı (n)'); axs[0].set_ylabel('İstatistiksel güç'); axs[0].set_ylim(0, 1.05); axs[0].set_xticks(fr_ns)
axs[0].set_title('a) Friedman testi (gözlenen veri,\nbootstrap simülasyonu)', fontsize=9.5)
for M, c in [(0.10, '#2E8B57'), (0.05, '#9BC79F')]: axs[1].plot(ns, tost_c[M], color=c, lw=2, label=f'±{M:.2f} mm')
axs[1].axhline(.8, color='grey', ls=':', lw=1); axs[1].axvline(15, color='k', lw=.8); axs[1].legend(frameon=False, fontsize=8, title='Eşdeğerlik sınırı', title_fontsize=8)
axs[1].set_xlabel('Katılımcı sayısı (n)'); axs[1].set_ylim(0, 1.05); axs[1].set_title('b) Eşdeğerlik testi (TOST)\nPrimescan – Füzyon', fontsize=9.5)
f.tight_layout(); save(f, 's12_guc_posthoc.png')

# Şekil 13: eşdeğerlik sınırının çalışmanın kendi verisi ve gücüne göre belirlenmesi
Ms = np.round(np.arange(0.02, 0.1501, 0.005), 3)
pow_M = [tost_power(M, n, 30000) for M in Ms]
se_d = sd/np.sqrt(n)
p_M = [max(stats.t.sf((md+M)/se_d, n-1), stats.t.cdf((md-M)/se_d, n-1)) for M in Ms]
min_M = max(abs(ci90[0]), abs(ci90[1]))
M80 = float(np.interp(0.80, pow_M, Ms))
pw10 = R['guc']['tost15']['0.1']
f, axs = plt.subplots(1, 2, figsize=(6.8, 3.1))
axs[0].plot(Ms, pow_M, color='#2E8B57', lw=2); axs[0].axhline(.8, color='grey', ls=':', lw=1)
axs[0].axvline(M80, color='grey', ls='--', lw=.8); axs[0].text(M80+.003, .35, f'%80 güç:\n±{M80:.3f} mm', fontsize=8)
axs[0].axvline(0.10, color='#C0632B', lw=1.2); axs[0].text(0.103, .55, f'Seçilen sınır\n±0.10 mm\ngüç = %{100*pw10:.1f}', fontsize=8, color='#C0632B')
axs[0].set_xlabel('Eşdeğerlik sınırı ±Δ (mm)'); axs[0].set_ylabel('İstatistiksel güç (n = 15)'); axs[0].set_ylim(0, 1.03)
axs[0].set_title('a) n = 15 ile eşdeğerlik testinin gücü', fontsize=9.5)
axs[1].plot(Ms, p_M, color='#5B4C9A', lw=2); axs[1].axhline(.05, color='grey', ls=':', lw=1); axs[1].text(0.128, .055, 'p = 0.05', fontsize=8)
axs[1].axvline(min_M, color='grey', ls='--', lw=.8); axs[1].text(min_M+.003, .12, f'En küçük\neşdeğerlik sınırı:\n±{min_M:.3f} mm', fontsize=8)
axs[1].axvline(0.10, color='#C0632B', lw=1.2)
axs[1].set_xlabel('Eşdeğerlik sınırı ±Δ (mm)'); axs[1].set_ylabel('TOST p değeri'); axs[1].set_ylim(0, 0.25)
axs[1].set_title('b) Eşdeğerlik sınırına göre TOST p değeri', fontsize=9.5)
f.tight_layout(); save(f, 's13_esdegerlik_siniri.png')
R['esd'] = dict(min_M=min_M, M80=M80)

R['ham'] = [[ids[i]] + [float(v) for v in X[i]] + [int(x) for x in rk[i]] for i in range(n)]
json.dump(R, open(os.path.join(out, 'sonuclar.json'), 'w'), ensure_ascii=False, indent=1, default=float)
print(json.dumps({k: v for k, v in R.items() if k != 'ham'}, ensure_ascii=False, indent=1, default=float))
