# -*- coding: utf-8 -*-
"""발표자료용 figure 생성 — 실제 결과 데이터 기반, 일관된 학술 스타일."""
import os, csv
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle, FancyArrowPatch

plt.rcParams.update({
    "font.family": "Malgun Gothic",
    "axes.unicode_minus": False,
    "figure.dpi": 200,
    "savefig.dpi": 200,
    "axes.edgecolor": "#888888",
    "axes.linewidth": 0.8,
    "font.size": 11,
})

INK="#1A1A1A"; GREEN="#2A7A2A"; RED="#C0392B"; GOLD="#B8860B"
GRAY="#9A9A9A"; LGRAY="#E6E6E6"; BLUE="#2C5F8A"; ORANGE="#E08A2A"
FIG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fig")
os.makedirs(FIG, exist_ok=True)

def load(path):
    with open(path, encoding="utf-8") as f:
        return list(csv.DictReader(f))

ASR = load(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                        "data", "processed", "asset_series.csv"))
def series(asset, col):
    rows=[r for r in ASR if r["asset_type"]==asset]
    rows.sort(key=lambda r:r["year_month"])
    return np.array([float(r[col]) for r in rows]), [r["year_month"] for r in rows]
def norm(a):
    a=np.asarray(a,float); return (a-a.min())/(a.max()-a.min()+1e-9)

def save(fig, name):
    fig.savefig(os.path.join(FIG,name), bbox_inches="tight", facecolor="white", pad_inches=0.12)
    plt.close(fig); print("saved", name)

VC = {"SIG":GREEN, "MARG":GOLD, "NS":GRAY}

# ---------------------------------------------------------------- 1. 시장 성장
def fig_market():
    fig,ax=plt.subplots(figsize=(4.6,3.2))
    yrs=["2024","2025","2029(E)"]; vals=[49,56,74]
    cols=[INK,INK,ORANGE]
    bars=ax.bar(yrs,vals,color=cols,width=0.6)
    for b,v in zip(bars,vals):
        ax.text(b.get_x()+b.get_width()/2,v+1.2,f"${v}B",ha="center",fontweight="bold",fontsize=12)
    ax.annotate("+14.3% YoY",xy=(1,56),xytext=(0.35,68),fontsize=11,color=RED,fontweight="bold",
                arrowprops=dict(arrowstyle="->",color=RED,lw=1.5))
    ax.set_ylim(0,85); ax.set_ylabel("Market Size (USD, Billions)")
    ax.set_title("U.S. Resale Market Growth",fontweight="bold")
    ax.spines[["top","right"]].set_visible(False)
    save(fig,"fig_market.png")

# ---------------------------------------------------------------- 2. 가격 시계열
def fig_price_series():
    fig,ax=plt.subplots(figsize=(9.2,2.7))
    cmap={"sneakers":(INK,"Sneakers"),"cards":(ORANGE,"Cards"),"lego":(BLUE,"LEGO")}
    for a,(c,lab) in cmap.items():
        p,m=series(a,"mean_price"); yn=norm(p)
        ax.plot(range(len(p)),yn,color=c,lw=2)
        ax.text(48.2,yn[-1],lab,color=c,fontweight="bold",fontsize=10,va="center")
    ticks=list(range(0,48,6)); _,m=series("cards","mean_price")
    ax.set_xticks(ticks); ax.set_xticklabels([m[i] for i in ticks],fontsize=9)
    ax.set_ylabel("Normalized\nPrice (0–1)"); ax.set_xlim(0,53)
    ax.set_title("Representative Resale Price Series, 2022–2025 (48 months)",fontweight="bold")
    ax.spines[["top","right"]].set_visible(False)
    save(fig,"fig_price_series.png")

# ---------------------------------------------------------------- 3. 선행성 예시
def fig_leadlag():
    fig,ax=plt.subplots(figsize=(5.6,3.2))
    p,m=series("cards","mean_price"); c1,_=series("cards","score_ch1")
    x=range(len(p))
    ax.plot(x,norm(c1),color=GREEN,lw=2.2,label="CH1 (search)")
    ax.plot(x,norm(p),color=INK,lw=2.2,label="Price",linestyle="-")
    ax.set_title("Cards: Search Interest (CH1) Leads Price",fontweight="bold",fontsize=12)
    ax.annotate("search peaks\nprecede price",xy=(11,0.55),xytext=(1.5,0.78),fontsize=9.5,color=RED,
                arrowprops=dict(arrowstyle="->",color=RED,lw=1.2))
    ticks=list(range(0,48,8)); ax.set_xticks(ticks); ax.set_xticklabels([m[i] for i in ticks],fontsize=8.5)
    ax.set_ylabel("Normalized (0–1)"); ax.set_xlim(0,47); ax.set_ylim(-0.05,1.22)
    ax.legend(loc="upper center",ncol=2,frameon=False,fontsize=9.5)
    ax.spines[["top","right"]].set_visible(False)
    save(fig,"fig_leadlag.png")

# ---------------------------------------------------------------- 4. Granger F bar
def fig_granger():
    data=[("sneakers","CH3",13.38,"SIG"),("sneakers","CH4",4.73,"SIG"),("sneakers","CH1",3.47,"SIG"),
          ("sneakers","CH2",1.55,"NS"),("sneakers","CH5",0.09,"NS"),
          ("cards","CH1",8.42,"SIG"),("cards","CH3",1.41,"NS"),("cards","CH2",0.73,"NS"),
          ("cards","CH5",0.61,"NS"),("cards","CH4",0.25,"NS"),
          ("lego","CH1",3.92,"MARG"),("lego","CH4",1.83,"NS"),("lego","CH3",0.40,"NS"),
          ("lego","CH2",0.32,"NS"),("lego","CH5",0.17,"NS")]
    data=data[::-1]
    fig,ax=plt.subplots(figsize=(5.8,4.2))
    labels=[f"{a[:2]} · {c}" for a,c,_,_ in data]
    vals=[v for _,_,v,_ in data]; cols=[VC[v] for *_,v in data]
    ax.barh(range(len(data)),vals,color=cols)
    ax.set_yticks(range(len(data))); ax.set_yticklabels(labels,fontsize=9)
    for i,(_,_,v,vd) in enumerate(data):
        ax.text(v+0.15,i,f"{v:.1f}",va="center",fontsize=8.5,color=INK)
    ax.set_xlabel("Granger F-statistic"); ax.set_xlim(0,15)
    ax.set_title("RQ1 · Granger Causality (15 tests)",fontweight="bold")
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=GREEN,label="Significant"),Patch(color=GOLD,label="Marginal"),
                       Patch(color=GRAY,label="Not sig.")],loc="lower right",frameon=False,fontsize=9)
    ax.spines[["top","right"]].set_visible(False)
    save(fig,"fig_granger.png")

# ---------------------------------------------------------------- 5. 삼각검증 매트릭스
def fig_triangulation():
    rows=[("cards","CH1","SIG","SIG","SIG","SIG"),
          ("sneakers","CH3","SIG","NS","SIG","NS"),
          ("sneakers","CH4","SIG","SIG","MARG","NS"),
          ("sneakers","CH1","MARG","NS","NS","NS"),
          ("cards","CH2","NS","SIG","NS","NS"),
          ("cards","CH5","NS","MARG","MARG","NS"),
          ("lego","CH1","MARG","NS","SIG","NS")]
    cols=["Single","DH","VAR","BH"]
    fig,ax=plt.subplots(figsize=(6.8,3.9))
    n=len(rows); m=len(cols)
    for i,r in enumerate(rows):
        for j in range(m):
            v=r[2+j]; y=n-1-i
            ax.add_patch(Rectangle((j,y),1,1,facecolor=VC[v],edgecolor="white",lw=2,alpha=0.92))
            ax.text(j+0.5,y+0.5,v,ha="center",va="center",color="white",fontweight="bold",fontsize=9)
    # cards CH1 강조 테두리
    ax.add_patch(Rectangle((0,n-1),m,1,fill=False,edgecolor=INK,lw=2.5))
    ax.set_xticks(np.arange(m)+0.5); ax.set_xticklabels(cols,fontsize=11,fontweight="bold")
    ax.xaxis.set_ticks_position("top"); ax.xaxis.set_label_position("top")
    ax.set_yticks(np.arange(n)+0.5)
    ax.set_yticklabels([f"{a} · {c}" for a,c,*_ in rows][::-1],fontsize=10)
    ax.set_xlim(0,m); ax.set_ylim(0,n); ax.set_aspect("equal")
    ax.tick_params(length=0)
    for s in ax.spines.values(): s.set_visible(False)
    ax.set_title("RQ1 Robustness — only  cards·CH1  survives all four",
                 fontweight="bold",fontsize=12,pad=22)
    save(fig,"fig_triangulation.png")

# ---------------------------------------------------------------- 6. 채널 구성 / Jaccard
def fig_jaccard():
    assets=["sneakers","cards","lego"]; chans=["CH1","CH2","CH3","CH4","CH5"]
    sig={("sneakers","CH1"),("sneakers","CH3"),("sneakers","CH4"),("cards","CH1")}
    fig,ax=plt.subplots(figsize=(5.6,3.0))
    for i,a in enumerate(assets):
        for j,c in enumerate(chans):
            y=len(assets)-1-i
            on=(a,c) in sig
            ax.add_patch(Rectangle((j,y),1,1,facecolor=(GREEN if on else "#F0F0F0"),
                         edgecolor="white",lw=2))
            if on: ax.scatter(j+0.5,y+0.5,s=120,color="white",marker="o",zorder=3)
    ax.set_xticks(np.arange(5)+0.5); ax.set_xticklabels(chans,fontsize=10)
    ax.set_yticks(np.arange(3)+0.5); ax.set_yticklabels(assets[::-1],fontsize=10)
    ax.set_xlim(0,5); ax.set_ylim(0,3); ax.set_aspect("equal"); ax.tick_params(length=0)
    for s in ax.spines.values(): s.set_visible(False)
    ax.set_title("RQ2 · Significant Channels per Asset",fontweight="bold",fontsize=11.5)
    ax.text(5.25,2.5,"Jaccard",fontsize=9.5,fontweight="bold")
    ax.text(5.25,1.95,"sn–cards = 0.333",fontsize=9,color=INK)
    ax.text(5.25,1.55,"sn–lego  = 0.000",fontsize=9,color=GRAY)
    ax.text(5.25,1.15,"cards–lego = 0.000",fontsize=9,color=GRAY)
    save(fig,"fig_jaccard.png")

# ---------------------------------------------------------------- 7. RMSE A vs B
def fig_rmse():
    assets=["Sneakers","Cards","LEGO"]; A=[452,334,176]; B=[569,329,199]
    x=np.arange(3); w=0.36
    fig,ax=plt.subplots(figsize=(4.8,3.2))
    ax.bar(x-w/2,A,w,label="Model A (all channels)",color=INK)
    ax.bar(x+w/2,B,w,label="Model B (Granger-selected)",color=ORANGE)
    for i in range(3):
        ax.text(x[i]-w/2,A[i]+6,str(A[i]),ha="center",fontsize=8.5)
        ax.text(x[i]+w/2,B[i]+6,str(B[i]),ha="center",fontsize=8.5)
    ax.set_xticks(x); ax.set_xticklabels(assets); ax.set_ylabel("Mean RMSE (lower = better)")
    ax.set_title("RQ3 · Prediction Error: Model A vs B",fontweight="bold")
    ax.legend(frameon=False,fontsize=9); ax.set_ylim(0,640)
    ax.spines[["top","right"]].set_visible(False)
    save(fig,"fig_rmse.png")

# ---------------------------------------------------------------- 8. DM by item
def fig_dm():
    items=[("jordan1",-4.29,"A"),("panda",-1.47,""),("yeezy",-1.09,""),("travis",0.38,""),("nb550",1.96,"Bm"),
           ("charizard1",1.23,""),("charizard2",1.49,""),("umbreon",-0.56,""),("rayquaza",-0.02,""),("pikachu",0.38,""),
           ("falcon",-2.28,"A"),("hogwarts",1.27,""),("titanic",-0.56,""),("porsche",-1.75,"Am"),("bugatti",0.42,"")]
    items=items[::-1]
    fig,ax=plt.subplots(figsize=(5.4,4.3))
    cols=[]
    for _,v,s in items:
        if s in ("A","Am"): cols.append(RED)
        elif s in ("Bm",): cols.append(GREEN)
        else: cols.append(GRAY)
    ax.barh(range(len(items)),[v for _,v,_ in items],color=cols)
    ax.axvline(0,color=INK,lw=0.8)
    ax.set_yticks(range(len(items))); ax.set_yticklabels([n for n,_,_ in items],fontsize=8.5)
    ax.set_xlabel("DM statistic   (< 0: A better   |   > 0: B better)")
    ax.set_title("RQ3 · Diebold–Mariano by Item (11/15 n.s.)",fontweight="bold",fontsize=11)
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=RED,label="A superior"),Patch(color=GREEN,label="B superior"),
                       Patch(color=GRAY,label="No difference")],loc="lower right",frameon=False,fontsize=8.5)
    ax.spines[["top","right"]].set_visible(False)
    save(fig,"fig_dm.png")

# ---------------------------------------------------------------- 9. SHAP
def fig_shap():
    chans=["CH1","CH2","CH3","CH4","CH5"]
    sn=[.125,.0,.011,.080,.124]; ca=[.048,.031,.025,.091,.004]; lg=[.036,.003,.0,.070,.021]
    fig,axes=plt.subplots(1,2,figsize=(9.2,3.2),gridspec_kw={"width_ratios":[1.25,1]})
    x=np.arange(5); w=0.26; ax=axes[0]
    ax.bar(x-w,sn,w,label="Sneakers",color=INK)
    ax.bar(x,ca,w,label="Cards",color=ORANGE)
    ax.bar(x+w,lg,w,label="LEGO",color=BLUE)
    ax.set_xticks(x); ax.set_xticklabels(chans); ax.set_ylabel("mean |SHAP|")
    ax.set_ylim(0,0.155)
    ax.set_title("Channel Contribution (Model A)",fontweight="bold",fontsize=11)
    ax.legend(loc="upper center",ncol=3,frameon=False,fontsize=9)
    ax.annotate("Granger F-1 (CH3)\nyet SHAP-lowest",xy=(2-w,.011),xytext=(2.1,.07),fontsize=8.5,color=RED,
                arrowprops=dict(arrowstyle="->",color=RED,lw=1))
    ax.spines[["top","right"]].set_visible(False)
    ax2=axes[1]
    feats=["price_vs_ma3","price_chg_l1","CH1","CH5","CH4","CH3"]; vals=[2.93,0.97,0.125,0.124,0.080,0.011]
    cc=[INK,INK,ORANGE,ORANGE,ORANGE,ORANGE]
    ax2.barh(range(len(feats))[::-1],vals,color=cc)
    ax2.set_yticks(range(len(feats))[::-1]); ax2.set_yticklabels(feats,fontsize=9)
    ax2.set_xlabel("mean |SHAP|")
    ax2.set_title("Sneakers: Momentum Dominates",fontweight="bold",fontsize=11)
    ax2.text(2.0,4.3,"price momentum\nabsorbs variance",fontsize=8.5,color=RED)
    ax2.spines[["top","right"]].set_visible(False)
    save(fig,"fig_shap.png")

# ---------------------------------------------------------------- 10. IRF
def fig_irf():
    rows=load(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                           "results","irf_results.csv"))
    def get(asset,ch):
        d=[r for r in rows if r["asset"]==asset and r["channel"]==ch]
        d.sort(key=lambda r:int(r["horizon"]))
        h=[int(r["horizon"]) for r in d]
        v=[float(r["irf_price_to_channel"]) for r in d]
        lo=[float(r["ci_lo"]) for r in d]; hi=[float(r["ci_hi"]) for r in d]
        return h,v,lo,hi
    fig,axes=plt.subplots(1,2,figsize=(9.0,3.0),sharex=True)
    for ax,(asset,ch,title,col) in zip(axes,[("cards","score_ch1","Cards · CH1 (positive)",GREEN),
                                             ("sneakers","score_ch3","Sneakers · CH3 (negative)",RED)]):
        h,v,lo,hi=get(asset,ch)
        if not h:
            ax.text(0.5,0.5,"n/a",ha="center"); continue
        ax.axhline(0,color=GRAY,lw=0.8,ls="--")
        ax.fill_between(h,lo,hi,color=col,alpha=0.18)
        ax.plot(h,v,color=col,lw=2.2,marker="o",ms=3)
        ax.set_title(title,fontweight="bold",fontsize=11)
        ax.set_xlabel("Months after shock")
        ax.spines[["top","right"]].set_visible(False)
    axes[0].set_ylabel("Price response")
    save(fig,"fig_irf.png")

# ---------------------------------------------------------------- 11. 2x2 프레임
def fig_2x2():
    fig,ax=plt.subplots(figsize=(5.6,3.6))
    ax.axhline(0.5,color=LGRAY,lw=1); ax.axvline(0.5,color=LGRAY,lw=1)
    pts=[("Sneakers",0.6,0.85,"CH3 News Sentiment",INK),
         ("Cards",0.32,0.25,"CH1 Search Trends",ORANGE),
         ("LEGO",0.12,0.5,"(no channel — supply)",BLUE)]
    for name,x,y,ch,c in pts:
        ax.scatter(x,y,s=320,color=c,zorder=3,edgecolor="white",lw=1.5)
        ax.text(x,y+0.10,name,ha="center",va="bottom",color=c,fontweight="bold",fontsize=11,zorder=4)
        ax.annotate(ch,xy=(x,y),xytext=(x,y-0.13),ha="center",fontsize=9,color="#444444")
    ax.set_xlim(0,1); ax.set_ylim(0,1)
    ax.set_xlabel("Supply Elasticity  →  (higher)")
    ax.set_ylabel("Information Asymmetry  →  (higher)")
    ax.set_title("Asset Channel Differentiation (2×2)",fontweight="bold",fontsize=11.5)
    ax.set_xticks([]); ax.set_yticks([])
    save(fig,"fig_2x2.png")

if __name__=="__main__":
    fig_market(); fig_price_series(); fig_leadlag(); fig_granger(); fig_triangulation()
    fig_jaccard(); fig_rmse(); fig_dm(); fig_shap(); fig_irf(); fig_2x2()
    print("ALL FIGURES DONE ->", FIG)
