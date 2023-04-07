#匯入套件
import pymysql
import pandas as pd
import numpy as np


patient_num_dis=[]    # 病患人數_疾病分
freq_dis=[]           # 就醫人次_疾病分
patient_num_city=[]   # 病患人數_縣市分
freq_city=[]          # 就醫人次_縣市分

# 匯入資料
for i in range(106,111):
    Rnum_dis = pd.read_excel("%d_急診就診統計.xls" %i, sheet_name='table4',header=7)
    patient_num_dis.append(Rnum_dis)
    Rfreq_dis = pd.read_excel("%d_急診就診統計.xls" %i, sheet_name='table7',header=7)
    freq_dis.append(Rfreq_dis)
    
    Rnum_city = pd.read_excel("%d_急診就診統計.xls" %i, sheet_name='table6',header=7)
    patient_num_city.append(Rnum_city)
    Rfreq_city = pd.read_excel("%d_急診就診統計.xls" %i, sheet_name='table9',header=7)
    freq_city.append(Rfreq_city)

# 刪除全空值的欄位
for i in range(5):
    # dis
    patient_num_dis[i].drop(['Unnamed: 3','Unnamed: 4','Unnamed: 5'], axis=1, inplace=True)
    freq_dis[i].drop(['Unnamed: 3','Unnamed: 4','Unnamed: 5'], axis=1, inplace=True)
    # city
    patient_num_city[i].drop(['Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 4'], axis=1, inplace=True)
    freq_city[i].drop(['Unnamed: 1','Unnamed: 2','Unnamed: 3','Unnamed: 4'], axis=1, inplace=True)

# 自訂函式：更改 疾病分list 欄位名
def change_dis_columns(lstname):
    for i in range(5):
        lstname[i].columns=['疾病分組代號','疾病分組名稱','疾病','疾病分類碼',
                            '總計','男','女',
                            '0~4歲','1_男','1_女','5~9歲','2_男','2_女',
                            '10~14歲','3_男','3_女','15~19歲','4_男','4_女',
                            '20~24歲','5_男','5_女','25~29歲','6_男','6_女',
                            '30~24歲','7_男','7_女','35~39歲','8_男','8_女',
                            '40~44歲','9_男','9_女','45~49歲','10_男','10_女',
                            '50~54歲','11_男','12_女','55~59歲','13_男','13_女',
                            '60~64歲','14_男','14_女','65~69歲','15_男','15_女',
                            '70~74歲','16_男','16_女','75~79歲','16_男','16_女',
                            '80~84歲','17_男','17_女','85歲以上','18_男','18_女']
change_dis_columns(patient_num_dis)
change_dis_columns(freq_dis)

# 自訂函式：更改 縣市分list 欄位名
def change_city_columns(lstname):
    for i in range(5):
        lstname[i].columns=['縣市',
                            '總計','男','女',
                            '0~4歲','1_男','1_女','5~9歲','2_男','2_女',
                            '10~14歲','3_男','3_女','15~19歲','4_男','4_女',
                            '20~24歲','5_男','5_女','25~29歲','6_男','6_女',
                            '30~24歲','7_男','7_女','35~39歲','8_男','8_女',
                            '40~44歲','9_男','9_女','45~49歲','10_男','10_女',
                            '50~54歲','11_男','12_女','55~59歲','13_男','13_女',
                            '60~64歲','14_男','14_女','65~69歲','15_男','15_女',
                            '70~74歲','16_男','16_女','75~79歲','16_男','16_女',
                            '80~84歲','17_男','17_女','85歲以上','18_男','18_女']
change_city_columns(patient_num_city)
change_city_columns(freq_city)

#  自訂函式：年齡區間分組，新增df欄位，合併年齡區間
def gr_age(lstname):
    for i in range(5):
        lstname[i]['0~14歲'] = lstname[i]['0~4歲'] + lstname[i]['5~9歲'] + lstname[i]['10~14歲']
        lstname[i]['15~64歲'] = lstname[i]['15~19歲'] + lstname[i]['20~24歲'] + lstname[i]['25~29歲'] + \
                                lstname[i]['30~24歲'] + lstname[i]['35~39歲'] + lstname[i]['40~44歲'] + \
                                lstname[i]['45~49歲'] + lstname[i]['50~54歲'] + lstname[i]['55~59歲'] + \
                                lstname[i]['60~64歲']
        lstname[i]['65歲以上'] = lstname[i]['65~69歲'] + lstname[i]['70~74歲'] + \
                                 lstname[i]['75~79歲'] + lstname[i]['80~84歲'] + lstname[i]['85歲以上'] 

gr_age(patient_num_dis)
gr_age(freq_dis)
gr_age(patient_num_city)
gr_age(freq_city)


# '疾病'若為空值，加入'疾病分組名稱'
for i in range(5):
    # 病患人數_疾病分
    disease = np.where(patient_num_dis[i]['疾病'].isnull(), patient_num_dis[i]['疾病分組名稱'], patient_num_dis[i]['疾病'])
    patient_num_dis[i]['疾病'] = disease
    # 最後一列  設定疾病 '不詳'
    patient_num_dis[i].iloc[181,2] = patient_num_dis[i].iloc[181,0]
    
    # 就醫人次_疾病分
    disease = np.where(freq_dis[i]['疾病'].isnull(), freq_dis[i]['疾病分組名稱'], freq_dis[i]['疾病'])
    freq_dis[i]['疾病'] = disease
    # 最後一列  設定疾病 '不詳'
    freq_dis[i].iloc[181,2] = freq_dis[i].iloc[181,0]

# '疾病'欄位只保留中文，因為匯入SQL發生錯誤，英文中有SQL保留字
for i in range(5):
    for j in range(1,182):
        patient_num_dis[i].iloc[j,2] = patient_num_dis[i].iloc[j,2].split()[0]
        freq_dis[i].iloc[j,2] = freq_dis[i].iloc[j,2].split()[0]


# 設定連線資料庫資訊
mysql_host='127.0.0.1'         # 主機名稱
mysql_user='root'              # 帳號
mysql_password='sql129042*'    # 密碼
mysql_db='emergency'           # 資料庫
mysql_port=3306                # port

# 連線資料庫
def connect_mysql(): # 開啟資料庫連接
    try:
        global connect, cursor
        connect = pymysql.connect(host = mysql_host, user = mysql_user,
                               password = mysql_password, db = mysql_db,
                               port=mysql_port, charset = 'utf8', # 資料庫編碼
                               use_unicode = True)
        cursor = connect.cursor()  # 使用cursor()方法操作資料庫
    except Exception as e:
        print("資料庫連接失敗：", e)

#呼叫連線資料庫函式
connect_mysql()
try:
    # SQL 創資料表、資料欄位(疾病分)
    for i in range(106,111):
        sqlstr = '''CREATE TABLE IF NOT EXISTS `data{}_dis`
                    (C_ID INT AUTO_INCREMENT PRIMARY KEY,
                     gr_dis TEXT, disease TEXT, 
                     patient_num BIGINT, num_male INT, num_female INT, 
                                      n_age_0_14 INT, n_age_15_64 INT, n_age_65up INT,
                     freq BIGINT, freq_male INT, freq_female INT,
                               f_age_0_14 INT, f_age_15_64 INT, f_age_65up INT)'''
        sqlstr = sqlstr.format(i)
        cursor.execute(sqlstr)
        print("data{:d}_dis資料表創建成功".format(i))
    
    # 將資料寫到資料表中(疾病分)
    for i in range(0,5):
        for j in range(1,182):
            sqlstr = '''INSERT INTO `data{0}_dis` 
                        (disease, patient_num, num_male, num_female,
                                  n_age_0_14, n_age_15_64, n_age_65up,
                                  freq, freq_male, freq_female,
                                  f_age_0_14, f_age_15_64, f_age_65up)
                     VALUES ('{1}',{2}, {3}, {4}, {5}, {6}, {7}, {8}, {9}, {10}, {11}, {12}, {13})'''
            sqlstr = sqlstr.format(i+106, patient_num_dis[i].iloc[j,2], patient_num_dis[i].iloc[j,4], patient_num_dis[i].iloc[j,5], patient_num_dis[i].iloc[j,6],\
                                   patient_num_dis[i].iloc[j,61], patient_num_dis[i].iloc[j,62], patient_num_dis[i].iloc[j,63],\
                                   freq_dis[i].iloc[j,4], freq_dis[i].iloc[j,5], freq_dis[i].iloc[j,6],\
                                   freq_dis[i].iloc[j,61], freq_dis[i].iloc[j,62], freq_dis[i].iloc[j,63])
            cursor.execute(sqlstr)
        print("data{}_dis資料寫入完成".format(i+106))
    
    # 自訂函式：新增疾病分組欄位資料
    def group_disease(a,b,c):
        for i in range(106,111):
            sqlstr="UPDATE data{:d}_dis SET gr_dis='{:s}' WHERE {:d} <= C_ID and C_ID <= {:d}"
            sqlstr=sqlstr.format(i, a,b,c)
            cursor.execute(sqlstr)
        connect.commit()
        print('data_dis分組欄位新增資料完成')
        
    # 依國際疾病分類為各疾病分組
    group_disease('infectious', 2, 7) # 感染及寄生蟲病
    group_disease('neoplasm', 9, 47)  # 腫瘤
    group_disease('blood', 49, 50)    # 血液和造血器官疾病
    group_disease('endorcrine_nutrition_metabolic', 52, 53) # 內分泌、代謝
    group_disease('mental_disorder', 55, 62) # 精神
    group_disease('nervous', 64, 68) # 神經系統
    group_disease('eye', 70, 72) # 眼
    group_disease('ear', 73, 73) # 耳與乳突
    group_disease('circulatory_system', 75, 85) #循環系統
    group_disease('respiratory', 87, 94) # 呼吸系統
    group_disease('digestive_system', 96, 115) # 消化系統
    group_disease('skin', 117, 119) # 皮膚及皮下組織
    group_disease('musculoskeletal', 121, 130) # 肌肉骨骼系統及結締組織
    group_disease('genitourinary', 132, 143) # 生殖泌尿系統
    group_disease('pregnancy',145,152) # 妊娠、生產與產褥期
    group_disease('perinatal', 154, 155) # 周產期
    group_disease('malformation', 156, 156) # 先天性畸形
    group_disease('condition', 158, 161)   # 症狀、徵候與臨床和實驗室的異常
    group_disease('injury', 163, 173) # 傷害、中毒與其他外因造成的特定影響
    group_disease('outreason', 174, 174) # 導致罹病或致死的外因
    group_disease('others',176,180)   # 影響健康狀況及健康服務
    group_disease('unknown',181, 181) # 不詳病因
    
      
    # SQL 創資料表、資料欄位(縣市分)
    for i in range(106,111):
        sqlstr = '''CREATE TABLE IF NOT EXISTS `data{}_city`
                    (C_ID INT AUTO_INCREMENT PRIMARY KEY,
                     city TEXT, 
                     patient_num BIGINT, num_male INT, num_female INT, 
                                      n_age_0_14 INT, n_age_15_64 INT, n_age_65up INT,
                     freq BIGINT, freq_male INT, freq_female INT,
                               f_age_0_14 INT, f_age_15_64 INT, f_age_65up INT)'''
        sqlstr = sqlstr.format(i)
        cursor.execute(sqlstr)
        print("data{:d}_city資料表創建成功".format(i))
    
    # 將資料寫到資料表中(縣市分)
    for i in range(0,5):
        for j in range(1,24):
            sqlstr = '''INSERT INTO `data{}_city` 
                        (city, patient_num, num_male, num_female,
                                  n_age_0_14, n_age_15_64, n_age_65up,
                                  freq, freq_male, freq_female,
                                  f_age_0_14, f_age_15_64, f_age_65up)
                     VALUES ('{}','{}', {}, {}, {}, {}, {}, {}, {}, {}, {}, {}, {})'''
            sqlstr = sqlstr.format(i+106, patient_num_city[i].iloc[j,0], patient_num_city[i].iloc[j,1], patient_num_city[i].iloc[j,2], patient_num_city[i].iloc[j,3],\
                                   patient_num_city[i].iloc[j,58], patient_num_city[i].iloc[j,59], patient_num_city[i].iloc[j,60],\
                                   freq_city[i].iloc[j,1], freq_city[i].iloc[j,2], freq_city[i].iloc[j,3],\
                                   freq_city[i].iloc[j,58], freq_city[i].iloc[j,59], freq_city[i].iloc[j,60])
            cursor.execute(sqlstr)
        print("data{}_city資料寫入完成".format(i+106))
    
    connect.commit()
except Exception as e:
        print("錯誤訊息：", e)
finally:
    connect.close()
    print("資料庫連線結束")

#%%
import warnings
warnings.filterwarnings("ignore")

# 疾病分組
# 從每年疾病分資料表中，依疾病分組，計算出每個疾病的 總計病患人數 & 就醫人次 & 平均就醫次數(就醫人次/患者人數)
connect_mysql()
data_sum=[]
for i in range(106,111):
    sql = pd.read_sql('''SELECT SUM(patient_num), SUM(freq), SUM(freq)/SUM(patient_num)
                             FROM data{}_dis;'''.format(i),con = connect)
    data_sum.append(sql)

# 疾病分組
# 從每年疾病分資料表中，依疾病分組，計算出每個疾病的 總計病患人數 & 就醫人次 & 平均就醫次數(就醫人次/患者人數)
connect_mysql()
data_dis=[]
for i in range(106,111):
    sql = pd.read_sql('''SELECT gr_dis, SUM(patient_num), SUM(freq)
                             FROM data{}_dis GROUP BY gr_dis;'''.format(i),con = connect)
    data_dis.append(sql)
    data_dis[i-106].drop(0,axis=0, inplace=True)
#觀察data_dis，發現病患人數和就醫人次最高的前10個疾病為condition, injury, respiratory, digestive_system, genitourinary, circulatory, endorcrine, musculoskeletal, skin, infections(index=18,19,11,10,14,9,4,13,12,1) , 因此刪除其他列
for i in range(106,111):
    data_dis[i-106].drop([2,3,5,6,7,8,15,16,17,20,21,22],axis=0, inplace=True)

# 性別分組
# 從每年疾病分資料表中，算出不同性別的 總計病患人數 & 就醫人次 & 平均就醫次數(就醫人次/患者人數)
connect_mysql()
data_gender=[]
for i in range(106,111):
    sql = pd.read_sql('''SELECT SUM(num_male), SUM(num_female), SUM(freq_male), SUM(freq_female), 
                         SUM(freq_male)/SUM(num_male), SUM(freq_female)/SUM(num_female)
                         FROM data{}_dis;'''.format(i),con = connect)
    data_gender.append(sql)

# 年齡區間分組
# 從每年資料表中，依年齡區間分組，計算出每個年齡區間的 總計病患人數 & 就醫人次 & 平均就醫次數(就醫人次/患者人數)
data_age=[]
connect_mysql()
for i in range(106,111):
    sql = pd.read_sql('''SELECT SUM(n_age_0_14), SUM(n_age_15_64), SUM(n_age_65up),
                         SUM(f_age_0_14), SUM(f_age_15_64), SUM(f_age_65up),
                         SUM(f_age_0_14)/SUM(n_age_0_14), SUM(f_age_15_64)/SUM(n_age_15_64), SUM(f_age_65up)/SUM(n_age_65up)
                         FROM data{}_dis'''\
                         .format(i),con = connect)
    data_age.append(sql)
    
# 縣市分
# 從每年age資料表中，依年齡區間分組，計算出每個年齡區間的 總計病患人數 & 就醫人次
data_city=[]
connect_mysql()
for i in range(106,111):
    sql = pd.read_sql('''SELECT city, patient_num, freq, freq/patient_num
                         FROM data{}_city'''\
                         .format(i),con = connect)
    data_city.append(sql)

#%%繪圖
# x:年
# y:長條圖 病患人數
# y2:折線圖 就醫人次

#匯入套件
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams['font.family']='Microsoft YaHei'
import math

#設定資料
x = np.arange(0,5)
y = [data_sum[i].iloc[0,0]/(10**7) for i in range(0,5)] #病患人數
y2 = [data_sum[i].iloc[0,1]/(10**7) for i in range(0,5)]#就醫人次

t1 = ['106','107','108','109','110']

#畫布
fig = plt.figure(figsize=(7,6),facecolor='#F0F0F0')

ax1 = fig.add_subplot()
ax2 = ax1.twinx()

#圖表類型
ax1.bar(x,y,width=0.6, color='#a0d8d6',label='病患人數') #長條圖 病患人數
ax2.plot(x,y2,linestyle='-.',marker='o',markersize=7,color='#b0955f', label='就醫人次') #折線圖 就醫人次


#雙y軸刻度範圍
ax1.set_ylim(0,2.5)
ax2.set_ylim(0,3)

# x軸、雙y軸標籤
ax1.set_xlabel('年分',fontsize=15)
ax1.set_ylabel('病患人數',fontsize=13)
ax2.set_ylabel('就醫人次',fontsize=13)

#標題
plt.title('急診病患人數 vs. 就醫人次',fontsize=20)

#圖例
ax1.legend(loc='upper left',fontsize=13)
ax2.legend(fontsize=13)

plt.xticks(x,labels=t1)

#在圖上增加文字
#每年病患人數
plt.text(-0.2, 0.2, str(math.floor(data_sum[0].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(0.8, 0.2, str(math.floor(data_sum[1].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(1.8, 0.2, str(math.floor(data_sum[2].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(2.8, 0.2, str(math.floor(data_sum[3].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(3.8, 0.2, str(math.floor(data_sum[4].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
#每年就醫人次
plt.text(-0.2, 2.4, str(math.floor(data_sum[0].iloc[0,1]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(0.8, 2.4, str(math.floor(data_sum[1].iloc[0,1]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(1.8, 2.5, str(math.floor(data_sum[2].iloc[0,1]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(2.9, 2.2, str(math.floor(data_sum[3].iloc[0,1]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(3.8, 2.1, str(math.floor(data_sum[4].iloc[0,1]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')

#顯示單位
plt.text(-1, 3.1,'單位(百萬)',fontsize=13)
plt.text(4.2, 3.1,'單位(千萬)',fontsize=13)

#顯示圖表
plt.tight_layout()
plt.savefig('num_freq.png',dpi=300)

plt.show()

#%%繪圖
# x:年
# y:長條圖 病患人數
# y3:折線圖 平均就醫人次

#匯入套件
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams['font.family']='Microsoft YaHei'
import math

#設定資料
x = np.arange(0,5)
y = [data_sum[i].iloc[0,0]/(10**7) for i in range(0,5)] #病患人數
y3 = [data_sum[i].iloc[0,2] for i in range(0,5)]#平均就醫人次

t1 = ['106','107','108','109','110']

#畫布
fig = plt.figure(figsize=(7,6),facecolor='#F0F0F0')

ax1 = fig.add_subplot()
ax2 = ax1.twinx()

#圖表類型
ax1.bar(x,y,width=0.6, color='#a0d8d6',label='病患人數') #長條圖 病患人數
ax2.plot(x,y3,linestyle='-.',marker='o',markersize=7,color='b', label='平均就醫次數') #折線圖 平均就醫人次

#雙y軸刻度範圍
ax1.set_ylim(0,2.5)
ax2.set_ylim(0,3)

# x軸、雙y軸標籤
ax1.set_xlabel('年分',fontsize=15)
ax1.set_ylabel('病患人數',fontsize=13)
ax2.set_ylabel('平均就醫次數',fontsize=13)

#標題
plt.title('急診病患人數 vs. 平均就醫次數',fontsize=20)

#圖例
ax1.legend(loc='upper left',fontsize=13)
ax2.legend(fontsize=13)

plt.xticks(x,labels=t1)

#在圖上增加文字
#每年病患人數
plt.text(-0.2, 0.2, str(math.floor(data_sum[0].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(0.8, 0.2, str(math.floor(data_sum[1].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(1.8, 0.2, str(math.floor(data_sum[2].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(2.8, 0.2, str(math.floor(data_sum[3].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')
plt.text(3.8, 0.2, str(math.floor(data_sum[4].iloc[0,0]/(10**7)*100)/100.0),fontsize=15, color='#6a6a6a')

plt.text(-1, 3.1,'單位(百萬)',fontsize=13)

#顯示圖表
plt.tight_layout()
plt.savefig('num_avgfreq.png',dpi=300)

plt.show()

#%% 設定多線條
# x: 年分
# y: 病患人數
# 圖例：年齡分組


import matplotlib.pyplot as plt

plt.rcParams['font.family']='Microsoft YaHei'

#病患人數
df_age = pd.DataFrame([data_age[i].iloc[0,0:3]/(10**6) for i in range(5)])    #病患人數
#df_age = pd.DataFrame([data_age[i].iloc[0,3:6] for i in range(5)])  #就醫人次
#df_age = pd.DataFrame([data_age[i].iloc[0,10:15] for i in range(5)]) #平均就醫人次

df_age.columns = ['0~14','15~64','65以上']

x = np.arange(5)
y1 = df_age['0~14']
y2 = df_age['15~64']
y5 = df_age['65以上']

#畫布
plt.figure(figsize=(7,6),facecolor='#F0F0F0')

plt.plot(x,y1, color='#16b016',label='0~14歲')
plt.plot(x,y2, color='#eaa45d',label='15~64歲')
plt.plot(x,y5, color='r',label='65歲以上')

# X/Y標籤
plt.xlabel('年分',fontsize=15)
plt.ylabel('病患人數',fontsize=15)

#標題
plt.title('急診病患人數(年齡別)',fontsize=20)


plt.ylim(1,12)#num
#plt.ylim(1,3)#avg_freq

plt.xticks(x,labels=t1)

plt.legend()

plt.text(-0.5, 12.3,'單位(百萬)')

plt.savefig('age_num.png',dpi=300)

plt.show()

#%% 設定多線條 ok
# x: 年分
# y: 就醫人次
# 圖例：年齡分組

import matplotlib.pyplot as plt

plt.rcParams['font.family']='Microsoft YaHei'

#就醫人次
df_age = pd.DataFrame([data_age[i].iloc[0,3:6]/(10**6) for i in range(5)])  #就醫人次

df_age.columns = ['0~14','15~64','65以上']

x = np.arange(5)
y1 = df_age['0~14']
y2 = df_age['15~64']
y5 = df_age['65以上']

t1 = ['106','107','108','109','110']


#畫布
plt.figure(figsize=(7,6),facecolor='#F0F0F0')

plt.plot(x,y1, color='#16b016',label='0~14歲')
plt.plot(x,y2, color='#eaa45d',label='15~59歲')
plt.plot(x,y5, color='r',label='65歲以上')


plt.ylabel('就醫人次',fontsize=15)
plt.title('急診就醫人次(年齡別)',fontsize=20)
plt.xlabel('年分',fontsize=15)

#plt.ylim(1,8)#num
plt.ylim(1,17)#freq
#plt.ylim(1,3)#avg_freq

plt.xticks(x,labels=t1)

plt.legend(loc='upper right')

plt.text(-0.5, 17.5,'單位(百萬)')

plt.savefig(r'D:\emergency\age_freq.png',dpi=300)

plt.show()

#%%繪圖
# x:各疾病
# y:病患人數
# 圖例：年分
# 5個x軸長條圖

#匯入套件
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams['font.family']='Microsoft YaHei'

#設定資料
df_dis = pd.DataFrame([data_dis[i]['SUM(patient_num)']/(10**6) for i in range(5)]).T

df_dis.columns=['106freq','107freq','108freq','109freq','110freq']

x = [i for i in range(1,11)]
y1 = df_dis['106freq']
y2 = df_dis['107freq']
y3 = df_dis['108freq']
y4 = df_dis['109freq']
y5 = df_dis['110freq']

t1 = ['感染','內分泌\n系統','循環\n系統','呼吸\n系統','消化\n系統','皮膚','肌肉骨骼\n系統','生殖\n泌尿\n系統','症狀\n異常\n(他處\n未歸類)','傷害\n中毒']

#畫布
plt.figure(figsize=(6, 4.5),facecolor='#F0F0F0')

#設定x軸位移
width = 1

x1=[i-(width/6)*2 for i in range(1,11)]
x2=[i-(width/6)*1 for i in range(1,11)]
x3=[i+(width/6)*0  for i in range(1,11)]
x4=[i+(width/6)*1  for i in range(1,11)]
x5=[i+(width/6)*2  for i in range(1,11)]

#圖表類型
plt.bar(x1,y1,width/6, color='pink',label='106')
plt.bar(x2,y2,width/6, color='#6495ED',label='107')
plt.bar(x3,y3,width/6, color='#73E68C',label='108')
plt.bar(x4,y4,width/6, color='#AE57A4',label='109')
plt.bar(x5,y5,width/6, color='#FFB366',label='110')

#標題
plt.title('急診病患人數(疾病別)',fontsize=20)

plt.xticks(x,labels=t1)

#X/Y 標籤
plt.xlabel('疾病分類',fontsize=15)
plt.ylabel('病患人數',fontsize=15)

#圖例
plt.legend(loc='upper left',fontsize=10)

#在圖上顯示單位
plt.text(-1, 2.4,'單位(百萬)')

#顯示圖表
plt.tight_layout()

plt.grid()
plt.savefig(r'D:\emergency\patient_num_dis.png',dpi=300)

plt.show()

#%%繪圖
# x:各疾病
# y:就醫人次
# 圖例：年分
# 5個x軸長條圖

#匯入套件
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams['font.family']='Microsoft YaHei'

#設定資料
df_dis = pd.DataFrame([data_dis[i]['SUM(freq)']/(10**6) for i in range(5)]).T

df_dis.columns=['106freq','107freq','108freq','109freq','110freq']

x = [i for i in range(1,11)]
y1 = df_dis['106freq']
y2 = df_dis['107freq']
y3 = df_dis['108freq']
y4 = df_dis['109freq']
y5 = df_dis['110freq']

t1 = ['感染','內分泌\n系統','循環\n系統','呼吸\n系統','消化\n系統','皮膚','肌肉骨骼\n系統','生殖\n泌尿\n系統','症狀\n異常\n(他處\n未歸類)','傷害\n中毒']

#畫布
plt.figure(figsize=(6, 4.5),facecolor='#F0F0F0')

#設定x軸位移
width = 1

x1=[ i-(width/6)*2 for i in range(1,11)]
x2=[ i-(width/6)*1 for i in range(1,11)]
x3=[ i+(width/6)*0  for i in range(1,11)]
x4=[ i+(width/6)*1  for i in range(1,11)]
x5=[ i+(width/6)*2  for i in range(1,11)]

#圖表類型
plt.bar(x1,y1,width/6, color='pink',label='106')
plt.bar(x2,y2,width/6, color='#6495ED',label='107')
plt.bar(x3,y3,width/6, color='#73E68C',label='108')
plt.bar(x4,y4,width/6, color='#AE57A4',label='109')
plt.bar(x5,y5,width/6, color='#FFB366',label='110')


#標題
plt.title('急診就醫人次(疾病別)',fontsize=20)

plt.xticks(x,labels=t1)

#X/Y 標籤
plt.xlabel('疾病分類',fontsize=15)
plt.ylabel('就醫人次',fontsize=15)

#圖例
plt.legend(loc='upper left',fontsize=10)


plt.text(-1, 3.3,'單位(百萬)')

#顯示圖表
plt.tight_layout()

plt.grid()
plt.savefig(r'D:\emergency\freq_dis.png',dpi=300)

plt.show()


#%%
#讀取各縣市人口數資料
population=[]
import pandas as pd
for i in range(106,111):
    county_population = pd.read_excel(r'D:\emergency\raw_data\鄉鎮戶數及人口數_%d年12月.xls' %i,header=2,nrows=25,usecols=[0,2])
    population.append(county_population)
    #刪除 "臺灣省"和"福建省" 資料
    population[i-106].drop([7,22],axis=0, inplace=True)
    
connect_mysql()
try:
    # SQL 創資料表、資料欄位(縣市分)
    sqlstr = '''CREATE TABLE IF NOT EXISTS `city_population`
                (C_ID INT AUTO_INCREMENT PRIMARY KEY,
                 city TEXT, population106 INT, population107 INT, population108 INT,
                 population109 INT, population110 INT)'''
    sqlstr = sqlstr.format(i)
    cursor.execute(sqlstr)
    print("city_population資料表創建成功")
    
    # 將資料寫到資料表中(縣市分)
    
    for j in range(1,23):
        sqlstr = '''INSERT INTO `city_population` 
                    (city, population106, population107, 
                     population108, population109, population110)
                    VALUES ('{}',{}, {}, {}, {}, {})'''
        sqlstr = sqlstr.format(population[0].iloc[j,0], population[0].iloc[j,1],
                               population[1].iloc[j,1], population[2].iloc[j,1],
                               population[3].iloc[j,1], population[4].iloc[j,1])
        cursor.execute(sqlstr)
    connect.commit()
    print("city_population資料寫入完成")
except:
    print('資料庫連線失敗')

connect_mysql()
data_county_population = pd.read_sql('''SELECT city, population106, population107, 
                                     population108, population109, population110 
                                     FROM city_population''', con=connect)

#刪除"不詳" 因為各縣市人口數無 "不祥"資料
for i in range(5):
    data_city[i].drop(22,axis=0,inplace=True)


#%%繪圖
# x:各縣市
# y:就醫人次 病患人數
# 圖例：年分
# 5個x軸長條圖

#匯入套件
import matplotlib.pyplot as plt
import numpy as np
plt.rcParams['font.family']='Microsoft YaHei'

#設定資料
df_city = pd.DataFrame([data_city[i]['patient_num'] for i in range(5)]).T
df_city.columns=['106num','107num','108num','109num','110num']

x = np.arange(22)
y1 = df_city['106num']/data_county_population['population106']*100
y2 = df_city['107num']/data_county_population['population107']*100
y3 = df_city['108num']/data_county_population['population108']*100
y4 = df_city['109num']/data_county_population['population109']*100
y5 = df_city['110num']/data_county_population['population110']*100

t1 = ['新北市','台北市','桃園市','台中市','台南市','高雄市','宜蘭縣','新竹縣','苗栗縣','彰化縣',
      '南投縣','雲林縣','嘉義縣','屏東縣','台東縣','花蓮縣','澎湖縣','基隆市','新竹市','嘉義市',
      '金門縣','連江縣']

#畫布
plt.figure(figsize=(16,6),facecolor='#F0F0F0')

#設定x軸位移
width = 1

x1=[i-(width/6)*2 for i in range(22)]
x2=[i-(width/6)*1 for i in range(22)]
x3=[i+(width/6)*0  for i in range(22)]
x4=[i+(width/6)*1  for i in range(22)]
x5=[i+(width/6)*2  for i in range(22)]

#圖表類型
plt.bar(x1,y1,width/6, color='pink',label='106')
plt.bar(x2,y2,width/6, color='#6495ED',label='107')
plt.bar(x3,y3,width/6, color='#73E68C',label='108')
plt.bar(x4,y4,width/6, color='#AE57A4',label='109')
plt.bar(x5,y5,width/6, color='#FFB366',label='110')

#X/Y 標籤
plt.xlabel('縣市',fontsize=15)
plt.ylabel('各縣市急診病患人數 / 各縣市人口數',fontsize=15)

plt.xticks(x,labels=t1)

plt.text(-2, 30,'%')

#圖例
plt.legend(loc='upper left',fontsize=10)

#顯示圖表
plt.tight_layout()
plt.grid()
plt.savefig('avg_patient_num_city.png',dpi=300)

plt.show()