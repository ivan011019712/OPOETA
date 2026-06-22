#1#####Begin##################################################################################################################
import os
import sys
import xlwings as xw
import pandas as pd
import numpy as np
from datetime import date,timedelta,datetime
import re
import warnings
warnings.filterwarnings("ignore", "This pattern is interpreted as a regular expression")

NORMAL_USE_ADDDATE=1
ETA_ADDDATE=3
td=date.today().strftime('%Y/%m/%d')

# OPOfile=sys.argv[1]
# RawMatlfile=sys.argv[1]
# MBfile=sys.argv[1]

OPOfile=sys.argv[1]#'ASUS_OPOShortage_20260421.xlsx'
RawMatlfile='../OPOs/汇总ASUS FRU 分料  0415.xlsx'
MBfile='../OPOs/ASUS SMT-0416.xlsx'
arrOPO=sys.argv[2]#'arrDf.xlsx'



arrDf=pd.read_excel(arrOPO,sheet_name='FRU',usecols=[0,1,2,3])
arrDf['mat']=arrDf['mat'].replace('',np.nan)
arrDf=arrDf.dropna(subset=['mat'])

arrDf.fillna('',inplace=True) 
arrDf.dropna(subset='mat',inplace=True)


    
#####OPO####################################################################################### 
OPOdf=pd.read_excel(OPOfile,sheet_name='Material',usecols=[1,24,25])##抓特定欄位
OPOdf.fillna('',inplace=True) 
Clddf=pd.read_excel(OPOfile,sheet_name='Calendar3')
Altdf=pd.read_excel(OPOfile,sheet_name='Alternative',usecols=[0,1,2])
Altdf.drop_duplicates(inplace=True)##刪除重複


# if(RawMatlfile==''): 
#     if(MBfile==''):
#         print('both empty')
#     else: ##分板only
#         ####分板list#######################################################################################
#         MBdf=pd.read_excel(MBfile,sheet_name='MB',usecols=[0,1,2,3])
#         ##Rename
#         #MBdf.rename(columns={'料号':'mat'},inplace=True)
#         ##Remove blank mat
#         MBdf['mat']=MBdf['mat'].replace('',np.nan)
#         MBdf=MBdf.dropna(subset=['mat'])
        
#         MBdf.fillna('',inplace=True) 
#         MBdf.dropna(subset='mat',inplace=True)
#         #MBdf['SA']=MBdf['mat'].str[:12]        
#         arrDf=MBdf.copy()

# else:
#     if(MBfile==''): ##分料only
#         #####欠料表#######################################################################################
#         PCdf=pd.read_excel(RawMatlfile,sheet_name='FRU',usecols=[0,1,2,3])
#         ##Rename
#         #PCdf.rename(columns={'料號':'mat','分料数量':'shortage'},inplace=True)
#         ##Remove blank mat
#         PCdf['mat']=PCdf['mat'].replace('',np.nan)
#         PCdf=PCdf.dropna(subset=['mat'])
        
#         PCdf.fillna('',inplace=True)    
#         arrDf=PCDf.copy()

        
#     else:        
#         #####欠料表#######################################################################################
#         PCdf=pd.read_excel(RawMatlfile,sheet_name='FRU',usecols=[0,1,2,3])
#         ##Rename
#         #PCdf.rename(columns={'料號':'mat','分料数量':'shortage'},inplace=True)
#         ##Remove blank mat
#         PCdf['mat']=PCdf['mat'].replace('',np.nan)
#         PCdf=PCdf.dropna(subset=['mat'])        
#         PCdf.fillna('',inplace=True)          
        
        
#         ####分板list#######################################################################################
#         MBdf=pd.read_excel(MBfile,sheet_name='MB',usecols=[0,1,2,3])
#         ##Rename
#         #MBdf.rename(columns={'料号':'mat'},inplace=True)
#         ##Remove blank mat
#         MBdf['mat']=MBdf['mat'].replace('',np.nan)
#         MBdf=MBdf.dropna(subset=['mat'])
        
#         MBdf.fillna('',inplace=True) 
#         MBdf.dropna(subset='mat',inplace=True)
#         #MBdf['SA']=MBdf['mat'].str[:12]        
#         arrDf=pd.concat([PCdf,MBdf],ignore_index=True, sort=False)


arrDf['MaterialETA']=''  
arrDf.rename(columns={'stotal':'STotal','shortage':'Support','remark':'Remark'},inplace=True)
#arrDf.info()


print('Step 1 OK')

#2#######################Function########################################################################
#### Convert String to Date
####(2026/04/14)################## Convert String to Date#################################################
def convertStr_Dt(strD):    
    if re.match(r'^\d{1,2}(/|-)\d{1,2}$',strD):
        strD=strD.replace('-','/')       
        strD=str(date.today().year)+'/'+strD
        date_check = pd.to_datetime(strD, format='%Y/%m/%d', errors='coerce')
        #print(date_check)
    elif re.match(r'^\d{4}(/|-)\d{1,2}(/|-)\d{1,2}$',strD):
        date_check = pd.to_datetime(strD, format='%Y/%m/%d', errors='coerce')       
    else:
        return "Wrong Date Format"
        
    if (pd.isna(date_check)):
        return "Wrong Format"            
    else:
        strD=pd.to_datetime(strD)
        return strD

####(2026/04/14)############# 輸入日期-> 加天數。 WK==1的話取當周最後一天#######################
def getMEDate(strDate,AddDays=1,WK=0):
    #print(strDate)
    cDate=convertStr_Dt(strDate)
    #print(cDate)
    did=Clddf[Clddf['Dt']==cDate]['DID']
    #print(did)
    if did.empty:
        return 'Not found in Calendar'
    if WK==1:
        #print(cDate.weekday())
        days_to_next_monday = 6 - cDate.weekday()        
        next_monday = cDate + timedelta(days=days_to_next_monday)        
        cDate=next_monday#.strftime("%Y/%m/%d"))  
        #print('WK==1',cDate)     
        
    did=Clddf[Clddf['Dt']==cDate]['DID']
    #did=did+AddDays
    #print(did)
    did=did.iloc[0]
    rr=Clddf[(Clddf['DID']>did) & (Clddf['WorkingDay']==1)]['DID'].head(AddDays)
    #print(rr)
    rr.iloc[AddDays-1]
    dt=Clddf[Clddf['DID']==rr.iloc[AddDays-1]]['Dt']   
    #dt=Clddf[Clddf['DID']==did]['Dt']
    if dt.empty:
        return 'Error'
    else:
        return dt.iloc[0].strftime('%Y/%m/%d')

####(2026/04/14)############# 輸入 dataframe ->生出 MaterialETA)######################
def getMEdf(df,addday=1,pn=r"(?:\d{1,2}/\d{1,2}\*)?\d+"):
    lstM,lstD,lstR=[],[],[]
    for row in df.itertuples():
        # 使用 row.欄位名稱 取值
        val1 = row.mat
        val2 = row.Remark
        val3 = row.Remark
        # print(f"{val1}, {val2}, {val3}")
        val2=getMatETA(val2,addday,pn)
        lstM.append(val1)
        lstD.append(val2)
        lstR.append(val3)

  
    #Convert to DataFrame    
    ETAdata2 = [lstM,lstD,lstR]
    return pd.DataFrame({'mat': lstM, 'MaterialETA': lstD, 'Remark :': lstR})

####(2026/04/14)############# 輸入Remark -> 抓住出日期(*Qty)######################
def getMatETA(tb,addday=1,pattern=r"(?:\d{1,2}/\d{1,2}\*)?\d+"): 
    #pattern = r"(?:\d{1,2}/\d{1,2}\*)?\d+"
    result=re.findall(pattern,tb)
    if type(result[0])==tuple:
        result=list(result[0])
    # print(result)
    if(len(result)==1):   #####(yyyy/mm/dd*qty)      
        if('WK' in result[0]): #####(WKmm/yy*qty)
            if('*' in result[0]):                
                strD=getMEDate(result[0][2:result[0].index('*')],ETA_ADDDATE,1) 
                return strD+'*'+result[0][result[0].index('*')+1:100] 
        elif('SWA' in result[0]):
            if('*' in result[0]):
                strD=getMEDate(td,ETA_ADDDATE) 
                return strD+'*'+result[0][result[0].index('*')+1:100]
        else: 
            if('*' not in result[0]): 
                if('/' in result[0]): ####Only(YYYY/MM/DD)
                    #print(result[0][0:result[0].index('*')],result[0][result[0].index('*')+1:100])
                    #print(result[0])
                    strD=getMEDate(result[0],ETA_ADDDATE) 
                    #print(strD)
                    return strD+'*'+"Fill shortage"  
                    

            #print(result[0][0:result[0].index('*')],result[0][result[0].index('*')+1:100])
            strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE) 
            #print(strD)
            return strD+'*'+result[0][result[0].index('*')+1:100] 
    elif(len(result)==2):
        if('*' in result[0]):            
            if('*' in result[1]): ####yyyy/mnm/dd*qty   yyyy/mnm/dd*qty
                strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE)
                tmp=strD+'*'+result[0][result[0].index('*')+1:100]
                del result[0:1]  
                strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE)
                return tmp+','+strD+'*'+result[0][result[0].index('*')+1:100]
            else: ##### yyyy/mm/dd*qty    qty
                if('/' not in result[1]):
                    strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE)
                    return strD+'*'+result[1]
                else:
                    strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE)
                    return strD+'*'+result[0][result[0].index('*')+1:100]
        else:
            if('*' not in result[1]):
                if(('SWA' in result[0]) and ('SWA' in result[1])):
                    strD=getMEDate(td,NORMAL_USE_ADDDATE)
                    return strD+'*'+"Fill shortage"  
                else:
                    strD=getMEDate(result[0]+'/'+result[1],ETA_ADDDATE) 
                    return strD+'*'+"Fill shortage"    
    elif(len(result)==3):
        if('*' not in result[0]):
            if ('SWA' in result[0]):
                #strD=str(date.today().year)+'/'+result[1]+'/'+result[2]            
                strD=getMEDate(td,ETA_ADDDATE)    
                #print(3,(datetime.strptime(strD,"%Y/%m/%d")+ timedelta(days=10)).strftime('%Y/%m/%d')+'*'+result[2])
                #return (datetime.strptime(strD,"%Y/%m/%d")+ timedelta(days=addday)).strftime('%Y/%m/%d')+'*'+result[2]  
                return strD+'*'+result[2][0]
            else:            
                #strD=str(date.today().year)+'/'+result[0]+'/'+result[1]            
                strD=getMEDate(result[0]+'/'+result[1],ETA_ADDDATE)    
                #print(3,(datetime.strptime(strD,"%Y/%m/%d")+ timedelta(days=10)).strftime('%Y/%m/%d')+'*'+result[2])
                #return (datetime.strptime(strD,"%Y/%m/%d")+ timedelta(days=addday)).strftime('%Y/%m/%d')+'*'+result[2]  
                return strD+'*'+result[2]
        else:
            strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE)
            return strD+'*'+result[0][result[0].index('*')+1:100]
        
    # elif(len(result)==4):
    #     if('*' in result[0]):
    #         if('*' in result[2]):
    #             #strD=str(date.today().year)+'/'+result[0][0:result[0].index('*')]
    #             strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE)   
    #             #tmp=(datetime.strptime(strD,"%Y/%m/%d")+ timedelta(days=addday)).strftime('%Y/%m/%d')+'*'+result[1]
    #             tmp=strD+'*'+result[1]
    #             del result[0:2]  
    #             #strD=str(date.today().year)+'/'+result[0][0:result[0].index('*')]
    #             strD=getMEDate(result[0][0:result[0].index('*')],ETA_ADDDATE)   
    #             #print(44,tmp+','+(datetime.strptime(strD,"%Y/%m/%d")+ timedelta(days=10)).strftime('%Y/%m/%d')+'*'+result[1])
    #             return tmp+','+strD+'*'+result[1]

    elif(len(result)==4): #### Multilpe ETA Date
      tmp=''
      if(('*' in result[0]) and ('*' in result[1])  and ('*' in result[2])  and ('*' in result[3])):
          for i in result:              
              strD=getMEDate(i[0:i.index('*')],ETA_ADDDATE)             
              tmp=tmp+strD+'*'+i[i.index('*')+1:10]+','
          return tmp
          
    elif(len(result)==5):  #### Multilpe ETA Date
      tmp=''
      if(('*' in result[0]) and ('*' in result[1])  and ('*' in result[2])  and ('*' in result[3])  and ('*' in result[4])):
          for i in result:              
              strD=getMEDate(i[0:i.index('*')],ETA_ADDDATE)              
              tmp=tmp+strD+'*'+i[i.index('*')+1:10]+','
          return tmp
    else:
        return '-'
########################################################################################################################      

print('Step 2 Functions OK')
# qq=arrDf[arrDf['mat']=='6033B0123901']
# # # print(qq[['MaterialETA']])
# getMEdf(qq,ETA_ADDDATE)
# # # re.findall(SWA_pn,'6037B0305206 SWA3*1+6037B0305506 SWA3*8+6037B0317006 SWA3*1')

#3#####處理分料資料################################################################################################################################
#Altdf[Altdf['Alternative']=='6054B3117501']
# Support>0 ,加3天收料
#arrDf.loc[arrDf['Support']>0,'MaterialETA']=(date.today() + timedelta(days=3)).strftime('%Y/%m/%d')+"*"+arrDf['Support'].astype(str)
####!!!如果沒有用copy()，會同步修改原 dataframe!!!

####Arrange Data :find ex : mm/dd;qty --> mm/dd*qty
#arrDf[arrDf['Remark'].str.contains(r'\d{1,2}/\d{1,2};\d+',case=False)==True]
arrDf['Remark']=arrDf['Remark'].str.replace(r'(\d{1,2}/\d{1,2});(\d+)',r'\1*\2',regex=True)

#0 blank and Support=0
arrDf.loc[(arrDf['Remark']=='') & (arrDf['Support']==0),'MaterialETA']='Unknown'

# 量产无需求
arrDf.loc[arrDf['Remark'].str.contains('量产无需求', na=False, regex=True),'MaterialETA']='量产无需求'#arrDf.loc[arrDf['Remark'].str.contains('量产无需求', na=False, regex=True),'MaterialETA']='-'

# '无法分'
arrDf.loc[arrDf['Remark'].str.contains('无法分', na=False, regex=True),'MaterialETA']='無法分'

# Support>0
arrDf.loc[arrDf['Support']>0,'MaterialETA']=getMEDate(td,NORMAL_USE_ADDDATE)+"*"+arrDf['Support'].astype(str)

# BS --> '-' & Support==0)
arrDf.loc[(arrDf['Remark'].str.contains(r'BS',case=False,regex=True)) & (arrDf['Support']==0),'MaterialETA']='-'


# 产线叫料 & Support==0 --> 加 3 天
arrDf.loc[(arrDf['Remark']=='產線叫料') & (arrDf['Support']==0),'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)
arrDf.loc[(arrDf['Remark']=='产线叫料') & (arrDf['Support']==0),'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)

# 用其他60代用。
# arrDf.loc[(arrDf['Remark'].str.contains(r'^6.{10}\d{1}$',case=False, regex=True)) & (arrDf['Support']==0),'MaterialETA']=getMEDate(td,NORMAL_USE_ADDDATE)+"*"+arrDf['STotal'].astype(str)

# 量产不足分料，待交期
arrDf.loc[(arrDf['Remark'].str.contains(r'量产不足分料，待交期',case=False, regex=True)==True) & (arrDf['Support']==0),'MaterialETA']='待交期'
#arrDf[(arrDf['Remark'].str.contains(r'量产不足分料，待交期',case=False)==True) & (arrDf['Support']==0)]

# LTB
arrDf.loc[(arrDf['Remark'].str.contains(r'LTB',case=False, regex=True)==True) & (arrDf['Support']==0),'MaterialETA']='LTB'

# 后确定分料状况
arrDf.loc[(arrDf['Remark'].str.contains(r'后确定分料状况',case=False, regex=True)==True) & (arrDf['Support']==0),'MaterialETA']='后确定分料状况'


# 厂商
arrDf.loc[(arrDf['Remark'].str.contains(r'厂商',case=False, regex=True)) & (arrDf['Support']==0),'MaterialETA']='廠商'
arrDf.loc[(arrDf['Remark'].str.contains(r'廠商',case=False, regex=True)) & (arrDf['Support']==0),'MaterialETA']='廠商'

# 模具
arrDf.loc[(arrDf['Remark'].str.contains(r'模具',case=False, regex=True)) & (arrDf['Support']==0),'MaterialETA']='模具'


# 確認??
arrDf.loc[(arrDf['Remark'].str.contains(r'确认',case=False, regex=True)) & (arrDf['Support']==0),'MaterialETA']='確認 what?'

# 新增不分
arrDf.loc[(arrDf['Remark'].str.contains(r'新增不分',case=False, regex=True)) & (arrDf['Support']==0),'MaterialETA']='-'

# 多給
arrDf.loc[(arrDf['Remark'].str.contains(r'多[給给]',case=False,regex=True)) & (arrDf['Support']==0),'MaterialETA']='-'

# 已給
arrDf.loc[(arrDf['Remark'].str.contains(r'已[給给]',case=False,regex=True)) & (arrDf['Support']==0),'MaterialETA']='-'

#13待齐料后再安排打板 待齐料后安排打板
arrDf.loc[(arrDf['Remark'].str.contains(r'待齐料后再安排打板',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'
arrDf.loc[(arrDf['Remark'].str.contains(r'待齐料后安排打板',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'
arrDf.loc[(arrDf['Remark'].str.contains(r'待给板',case=False, regex=True) & (arrDf['MaterialETA'].fillna('')=='')) & (arrDf['Support']==0),'MaterialETA']='-'
arrDf.loc[(arrDf['Remark'].str.contains(r'安排打板',case=False, regex=True)) & (arrDf['Support']==0),'MaterialETA']='-'

#不配套
arrDf.loc[(arrDf['Remark'].str.contains(r'不配套',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'

#打板时间待通知
arrDf.loc[(arrDf['Remark'].str.contains(r'打板时间待通知',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'

#打板时间待定
arrDf.loc[(arrDf['Remark'].str.contains(r'打板时间待定',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'

#待Asus分量
arrDf.loc[(arrDf['Remark'].str.contains(r'待Asus分量',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'

#未收到实单
arrDf.loc[(arrDf['Remark'].str.contains(r'未.*实单',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'

#无法接单
arrDf.loc[(arrDf['Remark'].str.contains(r'无法接单',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'

#新机型
arrDf.loc[(arrDf['Remark'].str.contains(r'新机型',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'


# #5.Remark use 替代,加3天收料
# arrDf.loc[(arrDf['Remark'].str.contains(r'^use\s?替代(?:\d+[a-z]?\d+)?$',case=False, regex=True)) & (arrDf['MaterialETA'].fillna('')==''),'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)
# #6.Remark use 60xx,加3天收料
# arrDf.loc[(arrDf['Remark'].str.contains(r'^use\s?6',case=False, regex=True)) & (arrDf['MaterialETA'].fillna('')==''),'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)
# arrDf.loc[(arrDf['Remark'].str.contains(r'Use 替代',case=False, regex=True)) & (arrDf['Support']==0) & (arrDf['MaterialETA'].isnull()),'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)

#自领 
arrDf.loc[(arrDf['Remark'].str.contains(r'自领',case=False, regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)


# USE Sxx/FA/NORMAL/SA
arrDf.loc[arrDf['Remark'].str.contains(r'USE SW',case=False, regex=True)==True,'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)
arrDf.loc[arrDf['Remark'].str.contains(r'USE FA',case=False, regex=True)==True,'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)
arrDf.loc[arrDf['Remark'].str.contains(r'USE NORMAL',case=False, regex=True)==True,'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)
arrDf.loc[arrDf['Remark'].str.contains(r'USE SA',case=False, regex=True)==True,'MaterialETA']=getMEDate(td,ETA_ADDDATE)+"*"+arrDf['STotal'].astype(str)

#arrDf[arrDf['mat']=='6017B2254301']
print('Step 3 OK')
#datetime.strptime('2026/2/3',"%Y/%m/%d").strftime('%Y/%m/%d')ˇ

#4#########清理資料###############################################################################################################################
# 輸入remark --> 得 日期*數量。

pn=r'\d{1,2}/\d{1,2}(?:\*\d+)?'#r"(?:\d{1,2}/\d{1,2}\*)?\d+"

WK_pn=r'WK\d{1,2}/\d{1,2}\*\d+' ##WK*Qty
###ETA###################################################################################
tt=arrDf[(arrDf['Remark'].str.contains(WK_pn,na=False,regex=True)) & (arrDf['MaterialETA'].fillna('')=='')]
ETAdf2=getMEdf(tt,ETA_ADDDATE,WK_pn)
arrDf.loc[(arrDf['Remark'].str.contains(WK_pn, na=False, regex=True)) & (arrDf['MaterialETA'].fillna('')==''),'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################

as_pn=r"(?:\d{1,2}/\d{1,2}\*)?\d+"
###入料分###################################################################################
tt=arrDf[arrDf['Remark'].str.contains(r'入料分',na=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='') ]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,as_pn)
arrDf.loc[arrDf['Remark'].str.contains(r'入料分', na=False, regex=True) & (arrDf['MaterialETA'].fillna('')=='') ,'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################

asf_pn=r"(\d+/\d+)入料後分(\d+)"
###入料後分###################################################################################
tt=arrDf[arrDf['Remark'].str.contains(asf_pn,na=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='') ]#[['mat','Remark']]
#tt=arrDf[arrDf['Remark'].str.extract(f"({r'(入料)(後分)'})",flags=2)[0]]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,asf_pn)
arrDf.loc[arrDf['Remark'].str.contains(asf_pn, na=False, regex=True) & (arrDf['MaterialETA'].fillna('')=='') ,'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
#arrDf.loc[arrDf['Remark'].str.extract(f"({r'(入料)(後分)'})",flags=2)[0],'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################

# 剩餘 WKmm/dd打板 --> '-'
arrDf.loc[(arrDf['Remark'].str.contains(r'WK\d{1,2}/\d{1,2}打板',case=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='')),'MaterialETA']='-'

dt_pn=r'\d{1,2}/\d{1,2}(?:\*\d+)?'
###ETA###################################################################################
tt=arrDf[arrDf['Remark'].str.contains(dt_pn,na=False,regex=True)& (arrDf['MaterialETA'].fillna('')=='')]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,dt_pn)
#print(ETAdf2)
arrDf.loc[arrDf['Remark'].str.contains(dt_pn, na=False, regex=True) & (arrDf['MaterialETA'].fillna('')==''),'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################


dt_pn=r'\d{1,2}/\d{1,2}(?:\*\d+)?'
###ETA###################################################################################
tt=arrDf[arrDf['Remark'].str.contains(dt_pn,na=False,regex=True)& (arrDf['MaterialETA'].fillna('')=='')]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,dt_pn)
#print(ETAdf2)
arrDf.loc[arrDf['Remark'].str.contains(dt_pn, na=False, regex=True) & (arrDf['MaterialETA'].fillna('')==''),'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################



SWA_pn=r'(SWA\d{1}\*\d+)'
####rSWAx*qty##################################################################################
tt=arrDf[(arrDf['Remark'].str.contains(SWA_pn,case=False, regex=True)) & (arrDf['MaterialETA'].fillna('')=='') ]
ETAdf2=getMEdf(tt,ETA_ADDDATE,SWA_pn)
#print(ETAdf2)
arrDf.loc[(arrDf['Remark'].str.contains(SWA_pn, na=False, regex=True)) & (arrDf['MaterialETA'].fillna('')=='') ,'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################



#tt=arrDf[(arrDf['Remark'].str.contains(r'^\d{1,2}/\d{1,2}\*',case=False)==True) & (arrDf['Support']==0)]
##预计/預計####################################################################################
tt=arrDf[arrDf['Remark'].str.contains(r'预计',na=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='') ]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,pn)
arrDf.loc[arrDf['Remark'].str.contains(r'预计', na=False, regex=True) & (arrDf['MaterialETA'].fillna('')=='') ,'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])

tt=arrDf[arrDf['Remark'].str.contains(r'預計',na=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='') ]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,pn)
arrDf.loc[arrDf['Remark'].str.contains(r'預計', na=False, regex=True) & (arrDf['MaterialETA'].fillna('')=='') ,'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################



ETA_pn=r'ETA\s?(\d{1,2}/\d{1,2})'
###ETA###################################################################################
tt=arrDf[arrDf['Remark'].str.contains(r'ETA',na=False,regex=True) & (arrDf['MaterialETA'].fillna('')=='') ]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,ETA_pn)
arrDf.loc[arrDf['Remark'].str.contains(r'ETA', na=False, regex=True) & (arrDf['MaterialETA'].fillna('')=='') ,'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################


# ####r'^\d{1,2}/\d{1,2}\*##################################################################################
# tt=arrDf[(arrDf['Remark'].str.contains(r'^\d{1,2}/\d{1,2}\*',case=False)==True) & (arrDf['Support']==0)]
# ETAdf2=getMEdf(tt,ETA_ADDDATE,pn)
# arrDf.loc[arrDf['Remark'].str.contains(r'^\d{1,2}/\d{1,2}\*', na=False, regex=True),'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
# ######################################################################################



# D_pn=r"(^\d{1,2}/\d{1,2}\*\d+),"
# ###ETA###################################################################################
# tt=arrDf[arrDf['Remark'].str.contains(D_pn,na=False,regex=True)& (arrDf['MaterialETA']=='-')]#[['mat','Remark']]
# ETAdf2=getMEdf(tt,ETA_ADDDATE,D_pn)
# #print(ETAdf2)
# arrDf.loc[arrDf['Remark'].str.contains(D_pn, na=False, regex=True) & (arrDf['MaterialETA']=='-'),'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
# ######################################################################################



###修模中###################################################################################
tt=arrDf[arrDf['Remark'].str.contains(r'修模中',na=False,regex=True)]#[['mat','Remark']]
ETAdf2=getMEdf(tt,ETA_ADDDATE,pn)
arrDf.loc[arrDf['Remark'].str.contains(r'修模中', na=False, regex=True),'MaterialETA'] = arrDf['mat'].map(ETAdf2.set_index('mat')['MaterialETA'])
######################################################################################



####將 "Fill Shortage" 替換成 Shortage 數量
idx=arrDf[arrDf['MaterialETA'].str.contains('Fill shortage', na=False, regex=False)].index
#idx
# mask=mask.set_index('mat')
arrDf.loc[idx,'MaterialETA']=arrDf.loc[idx].apply(lambda x: x['MaterialETA'].replace('Fill shortage', str(x['STotal'])), axis=1)

print('Step 4 OK')



# arrDf[arrDf['mat']=='6017B2254301*']


#5###########整合Alternative##############################################################################

ComDf=arrDf.copy()
ComDf['SA']=ComDf['mat'].str[:12]
# ComDf['Material']=''
#ComDf

temp_ComDf = pd.merge(ComDf,Altdf, left_on='SA', right_on='Alternative',how='left')
temp_ComDf.dropna(subset=['Material'],inplace=True) ####沒抓到的delete
#temp_ComDf
##清 duplicate再update
temp_ComDf2=temp_ComDf.drop_duplicates('SA',keep='first')

ComDf['Material'] =ComDf['SA'].map(temp_ComDf2.set_index('SA')['Material'])
#沒有Alternative的就用原本的        
ComDf.loc[ComDf['Material'].isnull(),'Material']=ComDf['SA']
ComDf['adRK']=''
ComDf['adRK']=ComDf['Remark'].astype(str)
ComDf.loc[ComDf['Material'].str.contains('*',regex=False),'adRK']=ComDf['mat'].astype(str)+':'+ComDf['Remark'].astype(str)
#ComDf['adRK']=ComDf['SA'].astype(str)+':'+ComDf['Remark'].astype(str)
#print(ComDf['MaterialETA'])

#ComDf #####要 drop 到空值，不然會fail.
arrComDf=ComDf.groupby('Material',as_index=False).agg(STotal=('STotal','sum'),Support=('Support','sum'),MaterialETA=('MaterialETA',lambda x:','.join(x.dropna().astype(str).sort_values())),Remark=('adRK',lambda x: '~'.join(x.dropna().astype(str)))) ##MaterialETA照日期排序
arrComDf[(arrComDf['MaterialETA'].str.contains('/',regex=False)==True)]
arrComDf[(arrComDf['MaterialETA'].str.contains(',-',regex=False)==True)].replace(',-','')

mk0 = arrComDf['MaterialETA'].str.contains(',-', regex=False, na=False)
arrComDf.loc[mk0, 'MaterialETA'] = arrComDf.loc[mk0, 'MaterialETA'].str.replace(',-', '', regex=False).str.strip(',')
mk1 = arrComDf['MaterialETA'].str.contains(',-', regex=False, na=False)
arrComDf.loc[mk1, 'MaterialETA'] = arrComDf.loc[mk1, 'MaterialETA'].str.replace('-,', '', regex=False).str.strip(',')

print('Step 5 OK')
# arrComDf.loc[arrComDf[(arrComDf['MaterialETA'].str.contains('-,',regex=False)==True)],'MaterialETA']=arrComDf['MaterialETA'].str.replace('-,','')
# arrComDf#[(arrComDf['MaterialETA'].str.contains('-,',regex=False)==True)]

        # ###整合Alternative
        # temp_MBdf = pd.merge(MBdf,Altdf, left_on='SA', right_on='Alternative',how='left')
        # temp_MBdf.dropna(subset=['Material'],inplace=True)
        # ##清 duplicate再update
        # MBdf['Material'] =MBdf['SA'].map(temp_MBdf.drop_duplicates('SA',keep='first').set_index('SA')['Material'])
        # #print(MBdf[MBdf['Material']=='1310A3764904*'])
        # ##沒有Alternative的就用原本的
        # MBdf.loc[MBdf['Material'].isnull(),'Material']=MBdf['SA']
        # arrMBdf=MBdf.groupby('Material',as_index=False).agg(Shortage=('stotal','sum'),Remark=('remark',lambda x:'~'.join(x.astype(str))))
        # ###補齊欄位以合併
        # arrMBdf['MaterialETA']=''
        # arrMBdf['Support']=0
        # arrDf=arrMBdf.copy()

#6### Insert Back to OPO #############################################################################################Finish#################################
app=xw.App(visible=True,add_book=False)
wb=app.books.open(OPOfile)
ws=wb.sheets['Material']
#Get OPO Material 
id_values = ws.range('B2').expand('down').value
#print('id_values : ',id_values)
#print("Excel 抓到的範例 ss:", id_values[:5])

#############Get combine df##################################################
finalDf=OPOdf.copy()
finalDf[['FMTA','FRK']]='',''
#finalDf.head()

############################
lookup = arrComDf.drop_duplicates('Material').set_index('Material')[['MaterialETA', 'Remark']]
lookup=lookup.rename(columns={'MaterialETA':'FMTA','Remark':'FRK'})

# 2. 關鍵修正：執行更新 FMTA 的動作
# 假設 OPOdf 與 arrComDf 是透過 'Material' 欄位關聯
finalDf = finalDf.set_index('Material')
finalDf.update(lookup) # 這一步才會把 FMTA 的值填進去
finalDf = finalDf.reset_index()
#finalDf
#gi_pn=r'(?:\(\w*\))?(#\w*),?(?:.*)?'
gi_pn=r'#([^,;]+)'
#if(re.findall(gi_pn,finalDf['Remark']))!=''
#finalDf['FRK']=finalDf.loc[finalDf['Remark'].str.contains(gi_pn,case=False,regex=True),'Remark']['Remark'].str.extract(gi_pn)+','+finalDf['FRK'].astype(str)
# 1. 先提取匹配的內容 (假設 gi_pn 是你的正規表示式變數)
extracted = finalDf[finalDf['FRK']!='']['Remark'].str.extract(f'({gi_pn})', flags=2)[0] # flags=2 是忽略大小寫
extracted=extracted.dropna()


# # 2. 定義更新邏輯：只有當提取成功 (notna) 且符合條件時才更新
# # 使用 np.where 或 fillna 處理原有的 FRK 避免 NaN 相加問題
mask = finalDf['Remark'].str.contains(gi_pn, case=False, na=False, regex=True)
# extracted[mask]
# 3. 更新 FRK 欄位：將新提取的值加上舊的值 (先處理空值避免 NaN)
finalDf.loc[mask, 'FRK'] = extracted[mask] + ',' + finalDf.loc[mask, 'FRK'].fillna('').astype(str)

#finalDf

# 4. (可選) 移除結尾多餘的逗號
finalDf['FRK'] = finalDf['FRK'].str.strip(',')
finalDf['FRK'] = finalDf['FRK'].str.lstrip(',')
finalDf['FRK'] = finalDf['FRK'].str.lstrip('-')
finalDf['FRK'] = finalDf['FRK'].str.lstrip('-')
finalDf['FRK'] = finalDf['FRK'].fillna('')

finalDf['FMTA'] = finalDf['FMTA'].str.strip(',')
finalDf['FMTA'] = finalDf['FMTA'].str.lstrip(',')
finalDf['FMTA'] = finalDf['FMTA'].str.lstrip('-')
finalDf['FMTA'] = finalDf['FMTA'].str.lstrip('-')
finalDf['FMTA'] = finalDf['FMTA'].fillna('')

#############################################################################

#####(2026/04/19)###################找出有代用但alternative 單獨出現的
mpn=ComDf[ComDf['Material'].isin(Altdf[Altdf['Alternative'].isin(finalDf['Material'])]['Material'])][['mat','MaterialETA','Remark']]
mpn.rename(columns={'mat':'Material','MaterialETA':'FMTA','Remark':'FRK'},inplace=True)
mpn=mpn.set_index('Material')[['FMTA','FRK']]
finalDf=finalDf.set_index('Material')
finalDf.update(mpn)
finalDf= finalDf.reset_index()
mpn= mpn.reset_index()
# print(finalDf[finalDf['Material'].isin(mpn['Material'])])
# finalDf

mapping = finalDf.set_index('Material')[['FMTA','FRK']].astype(str).T.to_dict('list') ##T是轉置的意思
#print("字典範例 Key:", list(mapping.keys())[:5])

final_column = [mapping.get(ss, ["",""]) for ss in id_values]
# print(final_column[:5])


# 一次填入兩欄 (例如從 J2 開始，會填入 J 和 K 欄)
ws.range('AI2').value = final_column

chng_c=finalDf[(finalDf['Material ETA']!=finalDf['FMTA']) & (finalDf['FMTA']!='-') & (finalDf['FMTA'].notnull()) & (finalDf['FMTA']!='')]['Material']
#chng_c
#type(chng_c)
for idx in chng_c.index:
    ws.range('AI'+str(idx+2)).color= (255, 255, 0)  # 黃色
    ws.range('AJ'+str(idx+2)).color= (255, 255, 0)  # 黃色

print('Finish!!')
