import pandas as pd
#사교육비 만들기
teachmoney='사교육비.xlsx' #출처 KOSIS 교육훈련탭 초중고사교육비조사 총액분포 엑셀
teachdf=pd.read_excel(teachmoney, index_col=0) #사교육비 읽어오기
teachdff=teachdf.iloc[0:2:,::5] #년도별로 사교육비 전체로 뽑기
teachdff['최소']=teachdff.loc['전체'].min() #사교육비 최소를 구하고 평균을구함.
teachdff['평균']=teachdff.loc['전체'].mean()

#*******************************************************************************
#소득분배지표 분석
buymoney='소득분배지표.xlsx' #출처 KOSIS 소득 소비 자산탭 소득분배지표 엑셀
buymoneydf=pd.read_excel(buymoney, index_col=0)
finalmoney=buymoneydf.iloc[13:25:2,1::3] #년도별 5분위 소득.
# finalmoney
meanlist=[]
minlist=[]
for n in finalmoney.index:
    meanlist.append(finalmoney.loc[n].mean())
    minlist.append(finalmoney.loc[n].min())
# meanlist
finalmoney['최소값']=minlist
finalmoney['평균값']=meanlist
# print(finalmoney) #최소값과 평균값 열을 추가한 년도별 5분위 소득.
# finalmoney


#*******************************************************************************
#사교육비 + 소득분배지표 합치기
sadf=teachdff.iloc[1::,::]
list(sadf.loc['전체'])
# finalmoney

yearlist=list(range(2013,2023))
yearlist.append("최소값")
yearlist.append("평균값")
# yearlist
finalmoney.loc[' ']="ㅤ"
finalmoney.loc['연도']=yearlist 
finalmoney.loc['사교육비 금액전체(단위:억원)']=list(sadf.loc['전체'])

#*******************************************************************************
#네이버 트렌드 검색량 추출
serch='검색어.xlsx' #출처 네이버트렌드 검색량 변화 데이터
df=pd.read_excel(serch,header=6,index_col=0)
serchdf=df.iloc[::,0::2]
sumlist=[]
for n in serchdf.index:
    sumlist.append(serchdf.loc[n].sum())
# sumlist
sum2016=int(sum(sumlist[:6])) #2016년 합계 6개월
sum2017=int(sum(sumlist[6:18])) #2017년 합계 1년
sum2018=int(sum(sumlist[18:30])) #2018년 합계
sum2019=int(sum(sumlist[30:42])) #2019년 합계
sum2020=int(sum(sumlist[42:54])) #2020년 합계
sum2021=int(sum(sumlist[54:66])) #2021년 합계
sum2022=int(sum(sumlist[66:78])) #2022년 합계
sum2023=int(sum(sumlist[78:])) #2023년 합계 6개월

minlist=[serchdf['금일'].min(),serchdf['명일'].min(),serchdf['심심'].min(),serchdf['연패'].min()]
maxlist=[serchdf['금일'].max(),serchdf['명일'].max(),serchdf['심심'].max(),serchdf['연패'].max()]
minmaxdf=pd.DataFrame([minlist,maxlist],index=['최소 검색','최대 검색'],columns=['금일','명일','심심','연패'])
print(minmaxdf)

perlist=[]
perlist.append(round(sum2016/sum(sumlist)*100))
perlist.append(round(sum2017/sum(sumlist)*100))
perlist.append(round(sum2018/sum(sumlist)*100))
perlist.append(round(sum2019/sum(sumlist)*100))
perlist.append(round(sum2020/sum(sumlist)*100))
perlist.append(round(sum2021/sum(sumlist)*100))
perlist.append(round(sum2022/sum(sumlist)*100))
perlist.append(round(sum2023/sum(sumlist)*100))

finalsumlist=[sum2016,sum2017,sum2018,sum2019,sum2020,sum2021,sum2022,sum2023]
# finalsumlist
sumperdf=pd.DataFrame([finalsumlist,perlist],columns=list(range(2016,2024)),index=['검색량합계','퍼센트'])
# sumperdf
sumperdf.loc[' ']="ㅤ"
sumperdf.loc['단어']=['금일','명일','심심','연패','금일','명일','심심','연패']
mx=minlist+maxlist
sumperdf.loc['최소/최대']=mx
total=0
for n in finalsumlist:
    total+=n
finalmean=total/len(finalsumlist)
pertotal=0
for n in perlist:
    pertotal+=n
permean=pertotal/len(perlist)
finalmeanlist=[finalmean,permean," "," "," "]
sumperdf['평균값']=finalmeanlist


#*******************************************************************************
#가설 : 문해능력이 사람간의 대화 즉 커뮤니케이션과 연간이 있다.
#커뮤니케이션 관련 데이터들을 분석해보자.


#*******************************************************************************
workdouble='맞벌이.xlsx' #출처 KOSIS 노동탭 년도별맞벌이가구
workdf=pd.read_excel(workdouble,index_col=0)
doublelist=list(workdf.loc['계'][1::5]) #2015년부터 2022년
# doublelist
gender='젠더갈등.xlsx'
genderdf=pd.read_excel(gender,header=3,index_col=0)
genderdf=genderdf.iloc[:-1].T
genderdf.plot(kind="bar")


#*******************************************************************************
gendermeanlist=[]
genderminlist=[]
# for n in genderdf.index:
#     gendermeanlist.append(genderdf.loc[n].mean())
#     genderminlist.append(genderdf.loc[n].min())
for n in range(0,4):
    gendermeanlist.append(genderdf.iloc[n].mean())
    genderminlist.append(genderdf.iloc[n].min())

genderdf.loc[' ']="ㅤ"
genderdf.loc['연도']=list(range(2015,2023))
genderdf.loc['맞벌이가구수(천)']=doublelist
index2013='2015년대비증감률'
indexlist=[]
for n in doublelist:    
    indexlist.append(n-doublelist[0])
indextotal=0
for n in indexlist:
    indextotal+=n
indexmean=indextotal/len(indexlist)
# indexlist
genderdf.loc[index2013]=indexlist
# genderdf.loc['  ']="ㅤ"
#1인가구를해보자
gagu1='1인가구.xlsx'  #출처 KOSIS 인구탭 년도별 1인가구
gagudf=pd.read_excel(gagu1,index_col=0)
gagudf.iloc[1,1::3]
gagulist=list(gagudf.iloc[1,1::3])
gagulist.append(int(gagudf.iloc[1,1::3].mean()))
gagulist #평균값까지 추가.
gaguyearlist=list(range(2015,2022))
gaguyearlist.append("평균값")
# genderdf
# gaguyearlist
# gagulist
genderdf.loc[' ㅤ  ']="ㅤㅤ "
genderdf.loc['1인가구연도']=gaguyearlist
genderdf.loc['1인가구량']=gagulist

genderdfmean=[] #젠더갈등 평균값 열에다가 추가
genderdfmean.append(int(genderdf.iloc[0].mean()))
genderdfmean.append(int(genderdf.iloc[1].mean()))
genderdfmean.append(int(genderdf.iloc[2].mean()))
genderdfmean.append(int(genderdf.iloc[3].mean()))
genderdfmean.append(" ")
genderdfmean.append("평균값")
workdfmean=int(workdf.loc['계'][1::5].mean()) #맞벌이가구 평균값도 추가.
genderdfmean.append(workdfmean)
genderdfmean.append(int(indexmean))
genderdfmean.append(" ")
genderdfmean.append(" ")
genderdfmean.append(" ")
genderdf["평균값"]=genderdfmean




#*******************************************************************************
#가족간 대화 도출
genderdf.loc[' ㅤ ㅤㅤ']="ㅤㅤ"
talkfile='가족간대화.xlsx' #kosis 월평균 가족간대화횟수 지역통계탭 경상남도
#거의매일, 자주, 가끔씩, 필요한경우에만, 거의없음, 모이기도힘듬
talkdf=pd.read_excel(talkfile, index_col=0)
talkdf.iloc[1]
talklist=[]
talklist.append(talkdf.iloc[1][0:4].sum()) #1993
talklist.append(talkdf.iloc[1][6:10].sum()) #1996
talklist.append(talkdf.iloc[1][12:16].sum()) #1999
talklist.append(talkdf.iloc[1][18:22].sum()) #2002
talklist.append(talkdf.iloc[1][24:28].sum()) #2005
talklist.append(talkdf.iloc[1][30:34].sum()) #2008
# talklist
talktotal=0
for n in talklist:
    talktotal+=n
talkmean=talktotal/len(talklist)
talklist.append(" ")
talklist.append(" ")
talklist.append(talkmean)
# talklist
talkyearlist=list(range(1993,2009,3))
talkyearlist.append(" ")
talkyearlist.append(" ")
talkyearlist.append("평균값")
# talkyearlist
genderdf.loc["대화연도"]=talkyearlist
genderdf.loc["가족간대화빈도"]=talklist #거의매일,자주,가끔씩,필요한경우에만 이것만 더함. 거의없음 모이기도힘듬은 제외
genderdf

#### 결론-  데이터분석결과 연도별로 커뮤니케이션이 떨어졌음을 확인 
#### 이러한 커뮤니케이션 능력 저하가 문해능력 저하와의 
# 상관관계가 있을것 같다는 가설이 성립되었습니다.