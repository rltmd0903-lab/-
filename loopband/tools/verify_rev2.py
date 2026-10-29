"""2차 재심본 표 3·4의 보고값만으로 하는 역검산(원자료 불필요).

사용: python3 verify_rev2.py
"""
import math
from scipy import stats
nE,nC=12,14
# name: (Epre,Epost,Cpre,Cpost,Etest,Ctest,G,T,GxT)  test=(kind,stat,p)
V2={
'KFES':((50.17,11.49),(41.75,9.82),(49.57,8.06),(50.14,7.58),('t',5.484,'<.001'),('t',-1.228,'.241'),(1.200,'.284'),(22.778,'.001'),(35.759,'<.001')),
'의자일어서기':((10.58,.90),(13.42,1.51),(8.64,1.01),(8.43,1.16),('Z',-3.153,'.002'),('Z',-1.732,'.083'),(64.775,'<.001'),(63.255,'<.001'),(102.708,'<.001')),
'팔굽히기':((13.75,1.71),(18.25,1.54),(11.00,.88),(10.71,1.07),('t',-17.234,'<.001'),('Z',-2.000,'.046'),(105.946,'<.001'),(194.625,'<.001'),(299.544,'<.001')),
'등뒤손잡기':((-19.83,2.29),(-9.42,1.98),(-20.86,2.03),(-20.93,2.27),('t',-12.837,'<.001'),('Z',-0.577,'.564'),(69.112,'<.001'),(157.974,'<.001'),(189.865,'<.001')),
'앉아윗몸굽히기':((-8.08,2.07),(1.25,.87),(-7.64,1.95),(-7.93,1.94),('t',-13.767,'<.001'),('Z',2.000,'.046'),(48.708,'<.001'),(169.492,'<.001'),(225.878,'<.001')),
'2분걷기':((84.00,6.41),(84.25,6.18),(81.93,4.71),(81.07,4.51),('Z',-1.342,'.180'),('Z',2.226,'.026'),(1.512,'231'),(2.880,'.103'),(7.321,'.012')),
'2.44m':((15.83,.83),(12.50,.90),(13.00,1.24),(13.00,1.18),('Z',-3.276,'.001'),('Z',0.000,'1.000'),(8.557,'.007'),(138.462,'<.001'),(161.538,'<.001')),
}
def fp(p): return '<.001' if p<.001 else ('%.3f'%p).lstrip('0')
print('=== A. 사전 동질성(보고된 M±SD로 역산, 최종 n=12/14) ===')
for k,(Ep,Eq,Cp,Cq,*_) in V2.items():
    s=stats.ttest_ind_from_stats(Ep[0],Ep[1],nE,Cp[0],Cp[1],nC,equal_var=True)
    w=stats.ttest_ind_from_stats(Ep[0],Ep[1],nE,Cp[0],Cp[1],nC,equal_var=False)
    sp=math.sqrt(((nE-1)*Ep[1]**2+(nC-1)*Cp[1]**2)/24); d=(Ep[0]-Cp[0])/sp
    print(f'{k:8s} Student t(24)={s.statistic:7.3f} p={fp(s.pvalue)} | Welch t={w.statistic:7.3f} p={fp(w.pvalue)} | d={d:5.2f}')
print('\n=== B. 보고된 F·t·Z의 p값 재계산 ===')
for k,(Ep,Eq,Cp,Cq,Et,Ct,G,T,I) in V2.items():
    out=[]
    for lab,(F,p) in zip('G T GxT'.split(),(G,T,I)):
        pc=stats.f.sf(F,1,24); out.append(f'{lab} F={F} 보고p={p} 계산p={fp(pc)} ηp²={F/(F+24):.2f}')
    for g,(kind,st,p),n in (('E',Et,nE),('C',Ct,nC)):
        pc = 2*stats.t.sf(abs(st),n-1) if kind=='t' else 2*stats.norm.sf(abs(st))
        alt = 2*stats.norm.sf(abs(st)) if kind=='t' else 2*stats.t.sf(abs(st),n-1)
        out.append(f'{g} {kind}={st} 보고p={p} 계산p={fp(pc)} (다른 검정으로 볼 때 {fp(alt)})')
    print(k); [print('   ',o) for o in out]
print('\n=== C. 시기/상호작용 F 비율 검산 (비율은 표준편차와 무관) ===')
for k,(Ep,Eq,Cp,Cq,Et,Ct,G,T,I) in V2.items():
    dE=Eq[0]-Ep[0]; dC=Cq[0]-Cp[0]
    unw=((dE+dC)/(dE-dC))**2
    dbar=(nE*dE+nC*dC)/26; wtd=26*dbar**2*(1/nE+1/nC)/(dE-dC)**2
    print(f'{k:8s} dE={dE:6.2f} dC={dC:6.2f}  보고 T/GxT={T[0]/I[0]:.3f}  가중평균 기대={wtd:.3f}  비가중 기대={unw:.3f}')
print('\n=== D. 대응 t → 차이값 SD, 사전·사후 상관 r, 효과크기 dz ===')
for k,(Ep,Eq,Cp,Cq,Et,Ct,G,T,I) in V2.items():
    for g,(pre,post),(kind,st,p),n in (('E',(Ep,Eq),Et,nE),('C',(Cp,Cq),Ct,nC)):
        if kind!='t' or st==0: continue
        d=post[0]-pre[0]; sdd=abs(d)*math.sqrt(n)/abs(st)
        r=(pre[1]**2+post[1]**2-sdd**2)/(2*pre[1]*post[1])
        print(f'{k:8s} {g} 평균변화={d:6.2f} SDdiff={sdd:5.2f} 사전·사후 r={r:5.2f} dz={abs(st)/math.sqrt(n):4.2f}')
print('\n=== E. GxT F → 합동 차이값 SD (두 집단 모두 t인 경우 대응 t와 비교) ===')
for k,(Ep,Eq,Cp,Cq,Et,Ct,G,T,I) in V2.items():
    dE=Eq[0]-Ep[0]; dC=Cq[0]-Cp[0]
    sd=math.sqrt((dE-dC)**2/(I[0]*(1/nE+1/nC)))
    msg=f'{k:8s} GxT에서 나온 합동 SDdiff={sd:5.2f}'
    if Et[0]=='t' and Ct[0]=='t':
        sE=abs(dE)*math.sqrt(nE)/abs(Et[1]); sC=abs(dC)*math.sqrt(nC)/abs(Ct[1])
        sp=math.sqrt((11*sE**2+13*sC**2)/24); msg+=f' | 대응 t에서 나온 합동 SDdiff={sp:5.2f} → GxT 기대 F={(dE-dC)**2/(sp**2*(1/nE+1/nC)):.3f}'
    print(msg)
print('\n=== F. Wilcoxon Z 역산: 변화가 0이 아닌 인원(k)과 동순위 구조 ===')
import itertools
def zval(diffs):
    nz=[x for x in diffs if x!=0]; k=len(nz)
    if k==0: return 0.0,0
    ab=sorted(set(abs(x) for x in nz))
    # average ranks
    vals=sorted(abs(x) for x in nz); ranks={}
    i=0
    while i<k:
        j=i
        while j<k and vals[j]==vals[i]: j+=1
        ranks[vals[i]]=(i+1+j)/2; i=j
    Tp=sum(ranks[abs(x)] for x in nz if x>0); Tm=sum(ranks[abs(x)] for x in nz if x<0)
    W=min(Tp,Tm); mu=k*(k+1)/4
    ties={}
    for v in vals: ties[v]=ties.get(v,0)+1
    var=k*(k+1)*(2*k+1)/24-sum(t**3-t for t in ties.values())/48
    return (W-mu)/math.sqrt(var),k
targets=[('의자 E',-3.153,12,13.42-10.58),('의자 C',-1.732,14,8.43-8.64),('팔 C',-2.000,14,10.71-11.00),('등뒤 C',-0.577,14,-20.93+20.86),('앉아 C',2.000,14,-7.93+7.64),('2분 E',-1.342,12,84.25-84.00),('2분 C',2.226,14,81.07-81.93),('2.44 E',-3.276,12,12.50-15.83)]
for name,Z,n,md in targets:
    tot=round(md*n)
    # 탐색 범위: 총변화와 같은 부호의 1~5, 그리고 반대 부호의 1~2
    sign=1 if tot>=0 else -1
    nonzero_vals=[sign*v for v in range(1,6)]+[-sign*v for v in (1,2)]
    sols=set()
    for k in range(0,n+1):
        for combo in itertools.combinations_with_replacement(nonzero_vals,k):
            if sum(combo)!=tot or k>n: continue
            z,kk=zval(list(combo))
            if abs(abs(z)-abs(Z))<0.0006: sols.add((n-k,combo))
        if len(sols)>3: break
    s=sorted(sols)[:3]
    print(f'{name}: 보고 Z={Z}, 평균변화×n≈{md*n:.2f} → 가능한 해(변화 0인 인원, 0이 아닌 변화값): {s}')
