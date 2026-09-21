import json, numpy as np, matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MPoly
plt.rcParams.update({"font.family":"Malgun Gothic","axes.unicode_minus":False,
  "figure.facecolor":"white","axes.facecolor":"white"})
C1,C2,C3="#5B4FCF","#B5650A","#0E8C7A"; INK="#2A2A33"
색={1:C1,2:C2,3:C3}
# 코드 → (표기명, 유형, 격차억, dx, dy)
후보 = {
 "11220": ("서초",      3, 106, -0.22,  0.10),
 "32010": ("춘천",      2,  54,  0.20,  0.06),
 "31240": ("동탄",      1,  77, -0.22,  0.00),
 "33020": ("충주",      2,  30,  0.20,  0.02),
 "29010": ("세종",      1,  82, -0.22,  0.08),
 "25030": ("대전 서구", 1, 298,  0.20,  0.10),
 "35030": ("익산",      1, 164, -0.20,  0.00),
 "35012": ("덕진",      1,  51,  0.20,  0.06),
 "35011": ("완산",      2,  66,  0.20, -0.24),
 "38030": ("진주",      2, 130,  0.20, -0.08),
}
g=json.load(open("data/cache/korea_sigungu.json", encoding="utf-8"))
def 링들(geom):
    return [geom["coordinates"][0]] if geom["type"]=="Polygon" else [p[0] for p in geom["coordinates"]]
def 면적(r):
    a=np.array(r); x,y=a[:,0],a[:,1]
    return abs(np.dot(x,np.roll(y,1))-np.dot(y,np.roll(x,1)))/2

fig,ax=plt.subplots(figsize=(4.2,6.6))
중심={}
for f in g["features"]:
    code=f["properties"]["code"]; hit=후보.get(code); rings=링들(f["geometry"])
    for r in rings:
        ax.add_patch(MPoly(np.array(r), closed=True,
            facecolor=색[hit[1]] if hit else "#F2F2F7",
            edgecolor="white" if hit else "#E4E4EC",
            lw=1.2 if hit else .45, zorder=3 if hit else 2))
    if hit: 중심[code]=np.array(max(rings, key=면적)).mean(axis=0)

for code,(nm,t,억,dx,dy) in 후보.items():
    x,y=중심[code]
    ax.scatter([x],[y], s=억*1.5+90, color=색[t], alpha=.30, edgecolor="none", zorder=4)
    ax.scatter([x],[y], s=40, color=색[t], edgecolor="white", lw=1.6, zorder=5)
    if abs(dy) > 0.05:
        ax.plot([x, x+dx*0.55], [y, y+dy*0.9], color=색[t], lw=1.0, zorder=5)
    ax.text(x+dx, y+dy, nm, fontsize=15, fontweight="bold", color=INK, zorder=6,
            ha="left" if dx>0 else "right", va="center")

for t,nm in [(1,"지역 전체 기회"),(2,"업종 간 이동"),(3,"채널 이동")]:
    ax.scatter([], [], s=150, color=색[t], label=nm)
ax.legend(loc="lower left", frameon=False, fontsize=14, handletextpad=.6,
          labelspacing=.7, bbox_to_anchor=(-0.03, 0.01))
ax.set_xlim(125.4, 130.2); ax.set_ylim(33.0, 38.75)
ax.set_aspect(1/np.cos(np.radians(36))); ax.axis("off")
fig.tight_layout(pad=0.2)
fig.savefig("data/그림_결과물/15_지도.png", dpi=200, bbox_inches="tight", facecolor="white")
from PIL import Image
im=Image.open("data/그림_결과물/15_지도.png"); r=im.size[0]/im.size[1]
print("저장", im.size, "비율 %.3f"%r, "→ h=4.25면 w=%.2f"%(4.25*r))
