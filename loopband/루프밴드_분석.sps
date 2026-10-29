* 루프밴드 연구: 원자료 불러오기, 변수 정의, 입력 점검, 원고 표 분석.
* build_spss_template.py가 만든 구문. 변수를 바꾸려면 그 스크립트를 고쳐서 다시 만든다.
* 실행 전: 아래 FILE= 경로를 실제 엑셀 파일 위치로 바꾼다.

* ---- 1. 불러오기 (엑셀 양식에 입력한 경우).
GET DATA /TYPE=XLSX
  /FILE='C:\루프밴드\원자료_입력양식_SPSS.xlsx'
  /SHEET=NAME 'data'
  /CELLRANGE=FULL
  /READNAMES=ON
  /ASSUMEDSTRWIDTH=100.
* 빈 .sav에 직접 입력한 경우에는 위 GET DATA 대신 다음 줄의 * 를 지우고 사용.
* GET FILE='C:\루프밴드\원자료_SPSS_빈파일.sav'.
SELECT IF NOT SYSMIS(id).
EXECUTE.
DATASET NAME raw WINDOW=FRONT.

* ---- 2. 변수 정의.
VARIABLE LABELS
  id '대상자 번호(기록지 번호)'
  group '집단'
  floor '이용 층'
  incl '최종 분석 포함'
  excl_rsn '제외 사유'
  excl_txt '제외 사유 상세'
  excl_when '제외 시점'
  attend '출석 횟수(24회 중)'
  band_fin '8주차 최종 밴드 색상'
  age '나이(만, 세)'
  height '신장(cm)'
  weight '체중(kg)'
  date_pre '사전검사 날짜'
  rater_pre '사전 측정자'
  kfes01_pre 'KFES-I 1번 문항 사전'
  kfes02_pre 'KFES-I 2번 문항 사전'
  kfes03_pre 'KFES-I 3번 문항 사전'
  kfes04_pre 'KFES-I 4번 문항 사전'
  kfes05_pre 'KFES-I 5번 문항 사전'
  kfes06_pre 'KFES-I 6번 문항 사전'
  kfes07_pre 'KFES-I 7번 문항 사전'
  kfes08_pre 'KFES-I 8번 문항 사전'
  kfes09_pre 'KFES-I 9번 문항 사전'
  kfes10_pre 'KFES-I 10번 문항 사전'
  kfes11_pre 'KFES-I 11번 문항 사전'
  kfes12_pre 'KFES-I 12번 문항 사전'
  kfes13_pre 'KFES-I 13번 문항 사전'
  kfes14_pre 'KFES-I 14번 문항 사전'
  kfes15_pre 'KFES-I 15번 문항 사전'
  kfes16_pre 'KFES-I 16번 문항 사전'
  chair_pre '의자 일어서기 사전(회/30초)'
  curl_pre '팔 굽히기 사전(회/30초)'
  step_pre '2분 제자리 걷기 사전(회)'
  bs1_pre '등 뒤로 손잡기 사전 1회(cm)'
  bs2_pre '등 뒤로 손잡기 사전 2회(cm)'
  sr1_pre '의자에 앉아 윗몸 앞으로 굽히기 사전 1회(cm)'
  sr2_pre '의자에 앉아 윗몸 앞으로 굽히기 사전 2회(cm)'
  tug1_pre '2.44 m 갔다 오기 사전 1회(초)'
  tug2_pre '2.44 m 갔다 오기 사전 2회(초)'
  date_post '사후검사 날짜'
  rater_post '사후 측정자'
  kfes01_post 'KFES-I 1번 문항 사후'
  kfes02_post 'KFES-I 2번 문항 사후'
  kfes03_post 'KFES-I 3번 문항 사후'
  kfes04_post 'KFES-I 4번 문항 사후'
  kfes05_post 'KFES-I 5번 문항 사후'
  kfes06_post 'KFES-I 6번 문항 사후'
  kfes07_post 'KFES-I 7번 문항 사후'
  kfes08_post 'KFES-I 8번 문항 사후'
  kfes09_post 'KFES-I 9번 문항 사후'
  kfes10_post 'KFES-I 10번 문항 사후'
  kfes11_post 'KFES-I 11번 문항 사후'
  kfes12_post 'KFES-I 12번 문항 사후'
  kfes13_post 'KFES-I 13번 문항 사후'
  kfes14_post 'KFES-I 14번 문항 사후'
  kfes15_post 'KFES-I 15번 문항 사후'
  kfes16_post 'KFES-I 16번 문항 사후'
  chair_post '의자 일어서기 사후(회/30초)'
  curl_post '팔 굽히기 사후(회/30초)'
  step_post '2분 제자리 걷기 사후(회)'
  bs1_post '등 뒤로 손잡기 사후 1회(cm)'
  bs2_post '등 뒤로 손잡기 사후 2회(cm)'
  sr1_post '의자에 앉아 윗몸 앞으로 굽히기 사후 1회(cm)'
  sr2_post '의자에 앉아 윗몸 앞으로 굽히기 사후 2회(cm)'
  tug1_post '2.44 m 갔다 오기 사후 1회(초)'
  tug2_post '2.44 m 갔다 오기 사후 2회(초)'
  note '비고'.
VALUE LABELS
 group 1 '실험집단' 2 '통제집단'
  / floor 2 '2층' 4 '4층'
  / incl 1 '최종 분석 포함' 0 '제외'
  / excl_rsn 1 '출석 미달' 2 '건강상 이유' 3 '사후검사 불참' 4 '참여 철회' 5 '기타'
  / excl_when 1 '사전검사 후 중재 시작 전' 2 '중재 중' 3 '사후검사 전'.
VARIABLE LEVEL group floor incl excl_rsn excl_when (NOMINAL).
FORMATS height (F6.1) weight (F6.1) bs1_pre (F6.1) bs2_pre (F6.1) sr1_pre (F6.1) sr2_pre (F6.1) tug1_pre (F6.2) tug2_pre (F6.2) bs1_post (F6.1) bs2_post (F6.1) sr1_post (F6.1) sr2_post (F6.1) tug1_post (F6.2) tug2_post (F6.2).

* ---- 3. 계산 변수: KFES-I 총점(16문항 모두 있을 때만), 유연성 최고값(MAX), 2.44 m 최고값(MIN).
COMPUTE kfes_pre = SUM.16(kfes01_pre TO kfes16_pre).
COMPUTE kfes_post = SUM.16(kfes01_post TO kfes16_post).
COMPUTE bs_pre = MAX(bs1_pre, bs2_pre).
COMPUTE bs_post = MAX(bs1_post, bs2_post).
COMPUTE sr_pre = MAX(sr1_pre, sr2_pre).
COMPUTE sr_post = MAX(sr1_post, sr2_post).
COMPUTE tug_pre = MIN(tug1_pre, tug2_pre).
COMPUTE tug_post = MIN(tug1_post, tug2_post).
COMPUTE d_kfes = kfes_pre - kfes_post.
COMPUTE same_kfes = (kfes_pre = kfes_post).
COMPUTE d_chair = chair_pre - chair_post.
COMPUTE same_chair = (chair_pre = chair_post).
COMPUTE d_curl = curl_pre - curl_post.
COMPUTE same_curl = (curl_pre = curl_post).
COMPUTE d_bs = bs_pre - bs_post.
COMPUTE same_bs = (bs_pre = bs_post).
COMPUTE d_sr = sr_pre - sr_post.
COMPUTE same_sr = (sr_pre = sr_post).
COMPUTE d_step = step_pre - step_post.
COMPUTE same_step = (step_pre = step_post).
COMPUTE d_tug = tug_pre - tug_post.
COMPUTE same_tug = (tug_pre = tug_post).
VARIABLE LABELS
  kfes_pre '낙상효능감(KFES-I) 사전'
  kfes_post '낙상효능감(KFES-I) 사후'
  d_kfes '낙상효능감(KFES-I) 차이(사전-사후)'
  same_kfes '낙상효능감(KFES-I) 사전=사후 동일(1)'
  chair_pre '의자 일어서기 사전'
  chair_post '의자 일어서기 사후'
  d_chair '의자 일어서기 차이(사전-사후)'
  same_chair '의자 일어서기 사전=사후 동일(1)'
  curl_pre '팔 굽히기 사전'
  curl_post '팔 굽히기 사후'
  d_curl '팔 굽히기 차이(사전-사후)'
  same_curl '팔 굽히기 사전=사후 동일(1)'
  bs_pre '등 뒤로 손잡기 사전'
  bs_post '등 뒤로 손잡기 사후'
  d_bs '등 뒤로 손잡기 차이(사전-사후)'
  same_bs '등 뒤로 손잡기 사전=사후 동일(1)'
  sr_pre '의자에 앉아 윗몸 앞으로 굽히기 사전'
  sr_post '의자에 앉아 윗몸 앞으로 굽히기 사후'
  d_sr '의자에 앉아 윗몸 앞으로 굽히기 차이(사전-사후)'
  same_sr '의자에 앉아 윗몸 앞으로 굽히기 사전=사후 동일(1)'
  step_pre '2분 제자리 걷기 사전'
  step_post '2분 제자리 걷기 사후'
  d_step '2분 제자리 걷기 차이(사전-사후)'
  same_step '2분 제자리 걷기 사전=사후 동일(1)'
  tug_pre '2.44 m 갔다 오기 사전'
  tug_post '2.44 m 갔다 오기 사후'
  d_tug '2.44 m 갔다 오기 차이(사전-사후)'
  same_tug '2.44 m 갔다 오기 사전=사후 동일(1)'.
EXECUTE.

* ---- 4. 입력 점검: 범위, 결측, 사전=사후 동일 인원. 원고 분석 전에 기록지와 대조.
FREQUENCIES VARIABLES=group incl excl_rsn excl_when /ORDER=ANALYSIS.
CROSSTABS /TABLES=excl_rsn excl_when BY group /CELLS=COUNT.
DESCRIPTIVES VARIABLES=age height weight attend kfes_pre chair_pre curl_pre bs_pre sr_pre step_pre tug_pre kfes_post chair_post curl_post bs_post sr_post step_post tug_post /STATISTICS=MEAN STDDEV MIN MAX.
MEANS TABLES=same_kfes same_chair same_curl same_bs same_sr same_step same_tug BY group /CELLS=SUM COUNT.
LIST VARIABLES=id group incl kfes_pre kfes_post chair_pre chair_post curl_pre curl_post bs_pre bs_post.
LIST VARIABLES=id group incl sr_pre sr_post step_pre step_post tug_pre tug_post.

* ---- 5. 최종 분석 대상만 사용(incl=1).
USE ALL.
COMPUTE filter_incl = (incl = 1).
FILTER BY filter_incl.
EXECUTE.

* 5-1. 표 1과 사전 동질성: 독립표본 t(Levene 결과에 따라 등분산 가정 행 또는 Welch 행을 사용).
T-TEST GROUPS=group(1 2) /VARIABLES=age height weight kfes_pre chair_pre curl_pre bs_pre sr_pre step_pre tug_pre /MISSING=ANALYSIS /CRITERIA=CI(.95).

* 5-2. KFES-I 신뢰도(Cronbach α).
RELIABILITY /VARIABLES=kfes01_pre TO kfes16_pre /SCALE('KFES-I 사전') ALL /MODEL=ALPHA.
RELIABILITY /VARIABLES=kfes01_post TO kfes16_post /SCALE('KFES-I 사후') ALL /MODEL=ALPHA.

* 5-3. 집단별 차이값 정규성(Shapiro-Wilk). p≥.05면 대응 t, p<.05면 Wilcoxon 결과를 원고에 사용.
EXAMINE VARIABLES=d_kfes d_chair d_curl d_bs d_sr d_step d_tug BY group
  /PLOT NPPLOT /STATISTICS NONE /NOTOTAL /MISSING=PAIRWISE.

* 5-4. 집단 내 사전·사후 비교(집단별). 대응 t는 사전-사후, Wilcoxon Z는 SPSS 출력 그대로 보고.
SORT CASES BY group.
SPLIT FILE LAYERED BY group.
T-TEST PAIRS=kfes_pre chair_pre curl_pre bs_pre sr_pre step_pre tug_pre WITH kfes_post chair_post curl_post bs_post sr_post step_post tug_post (PAIRED) /CRITERIA=CI(.95) /MISSING=ANALYSIS.
NPAR TESTS /WILCOXON=kfes_pre chair_pre curl_pre bs_pre sr_pre step_pre tug_pre WITH kfes_post chair_post curl_post bs_post sr_post step_post tug_post (PAIRED) /MISSING ANALYSIS.
SPLIT FILE OFF.

* 5-5. 반복측정 이원분산분석(집단×시기, 제Ⅲ유형, ηp²).
* 낙상효능감(KFES-I).
GLM kfes_pre kfes_post BY group
  /WSFACTOR=time 2 Polynomial
  /METHOD=SSTYPE(3)
  /PRINT=DESCRIPTIVE ETASQ
  /CRITERIA=ALPHA(.05)
  /WSDESIGN=time
  /DESIGN=group.
* 의자 일어서기.
GLM chair_pre chair_post BY group
  /WSFACTOR=time 2 Polynomial
  /METHOD=SSTYPE(3)
  /PRINT=DESCRIPTIVE ETASQ
  /CRITERIA=ALPHA(.05)
  /WSDESIGN=time
  /DESIGN=group.
* 팔 굽히기.
GLM curl_pre curl_post BY group
  /WSFACTOR=time 2 Polynomial
  /METHOD=SSTYPE(3)
  /PRINT=DESCRIPTIVE ETASQ
  /CRITERIA=ALPHA(.05)
  /WSDESIGN=time
  /DESIGN=group.
* 등 뒤로 손잡기.
GLM bs_pre bs_post BY group
  /WSFACTOR=time 2 Polynomial
  /METHOD=SSTYPE(3)
  /PRINT=DESCRIPTIVE ETASQ
  /CRITERIA=ALPHA(.05)
  /WSDESIGN=time
  /DESIGN=group.
* 의자에 앉아 윗몸 앞으로 굽히기.
GLM sr_pre sr_post BY group
  /WSFACTOR=time 2 Polynomial
  /METHOD=SSTYPE(3)
  /PRINT=DESCRIPTIVE ETASQ
  /CRITERIA=ALPHA(.05)
  /WSDESIGN=time
  /DESIGN=group.
* 2분 제자리 걷기.
GLM step_pre step_post BY group
  /WSFACTOR=time 2 Polynomial
  /METHOD=SSTYPE(3)
  /PRINT=DESCRIPTIVE ETASQ
  /CRITERIA=ALPHA(.05)
  /WSDESIGN=time
  /DESIGN=group.
* 2.44 m 갔다 오기.
GLM tug_pre tug_post BY group
  /WSFACTOR=time 2 Polynomial
  /METHOD=SSTYPE(3)
  /PRINT=DESCRIPTIVE ETASQ
  /CRITERIA=ALPHA(.05)
  /WSDESIGN=time
  /DESIGN=group.

* 5-6. 민감도 분석: 사전값을 공변량으로 한 공분산분석과 회귀기울기 동질성.
* 기울기 동질성: 아래 첫 번째 분석에서 group*사전값 상호작용 p<.05면 공분산분석 해석 제한.
* 낙상효능감(KFES-I).
UNIANOVA kfes_post BY group WITH kfes_pre
  /METHOD=SSTYPE(3)
  /DESIGN=kfes_pre group kfes_pre*group.
UNIANOVA kfes_post BY group WITH kfes_pre
  /METHOD=SSTYPE(3)
  /EMMEANS=TABLES(group) WITH(kfes_pre=MEAN)
  /PRINT=DESCRIPTIVE ETASQ PARAMETER
  /DESIGN=kfes_pre group.
* 의자 일어서기.
UNIANOVA chair_post BY group WITH chair_pre
  /METHOD=SSTYPE(3)
  /DESIGN=chair_pre group chair_pre*group.
UNIANOVA chair_post BY group WITH chair_pre
  /METHOD=SSTYPE(3)
  /EMMEANS=TABLES(group) WITH(chair_pre=MEAN)
  /PRINT=DESCRIPTIVE ETASQ PARAMETER
  /DESIGN=chair_pre group.
* 팔 굽히기.
UNIANOVA curl_post BY group WITH curl_pre
  /METHOD=SSTYPE(3)
  /DESIGN=curl_pre group curl_pre*group.
UNIANOVA curl_post BY group WITH curl_pre
  /METHOD=SSTYPE(3)
  /EMMEANS=TABLES(group) WITH(curl_pre=MEAN)
  /PRINT=DESCRIPTIVE ETASQ PARAMETER
  /DESIGN=curl_pre group.
* 등 뒤로 손잡기.
UNIANOVA bs_post BY group WITH bs_pre
  /METHOD=SSTYPE(3)
  /DESIGN=bs_pre group bs_pre*group.
UNIANOVA bs_post BY group WITH bs_pre
  /METHOD=SSTYPE(3)
  /EMMEANS=TABLES(group) WITH(bs_pre=MEAN)
  /PRINT=DESCRIPTIVE ETASQ PARAMETER
  /DESIGN=bs_pre group.
* 의자에 앉아 윗몸 앞으로 굽히기.
UNIANOVA sr_post BY group WITH sr_pre
  /METHOD=SSTYPE(3)
  /DESIGN=sr_pre group sr_pre*group.
UNIANOVA sr_post BY group WITH sr_pre
  /METHOD=SSTYPE(3)
  /EMMEANS=TABLES(group) WITH(sr_pre=MEAN)
  /PRINT=DESCRIPTIVE ETASQ PARAMETER
  /DESIGN=sr_pre group.
* 2분 제자리 걷기.
UNIANOVA step_post BY group WITH step_pre
  /METHOD=SSTYPE(3)
  /DESIGN=step_pre group step_pre*group.
UNIANOVA step_post BY group WITH step_pre
  /METHOD=SSTYPE(3)
  /EMMEANS=TABLES(group) WITH(step_pre=MEAN)
  /PRINT=DESCRIPTIVE ETASQ PARAMETER
  /DESIGN=step_pre group.
* 2.44 m 갔다 오기.
UNIANOVA tug_post BY group WITH tug_pre
  /METHOD=SSTYPE(3)
  /DESIGN=tug_pre group tug_pre*group.
UNIANOVA tug_post BY group WITH tug_pre
  /METHOD=SSTYPE(3)
  /EMMEANS=TABLES(group) WITH(tug_pre=MEAN)
  /PRINT=DESCRIPTIVE ETASQ PARAMETER
  /DESIGN=tug_pre group.

* 5-7. 사전값 분포가 겹치는지 확인(공분산분석 조정 평균 해석 전에 확인).
MEANS TABLES=kfes_pre chair_pre curl_pre bs_pre sr_pre step_pre tug_pre BY group /CELLS=COUNT MEAN STDDEV MIN MAX.
FILTER OFF.
USE ALL.
EXECUTE.
