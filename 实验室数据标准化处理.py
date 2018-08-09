
# coding: utf-8

# In[1]:


import pandas as pd
import math,xlrd
from scipy.stats import ttest_rel #配对t检验
from scipy.stats import ranksums#,wilcoxon #wilcox非参检验
from scipy.stats import wilcoxon
print(wilcoxon([1,2,3,6,7,6,1,4,3,2,2,1,5],[2,3,5,4,4,3,2,4,6,7,7,3,2]))
print(wilcoxon)


# In[2]:


#产品对照表

def excel_dict(excelfile,sheet):
    pro = []
    workbook = xlrd.open_workbook(excelfile)
    sheet_names= workbook.sheet_names()
    sheet = workbook.sheet_by_name(sheet)
    nrows = sheet.nrows
    
    prodict = {}
    for i in range(1,nrows):
        pro.append(sheet.cell_value(i,0))
        prodict[sheet.cell_value(i,0)] = sheet.cell_value(i,1)
    #print(prodict)
    return prodict,pro
#pro_code(u'保湿模板所需数据.xlsx')
procode =  excel_dict(u'保湿模板所需数据1.xlsx','产品对照')
prodict =procode[0]

pro = procode[1].pop(0)
PRODUCT_CODE = procode[1]
print(PRODUCT_CODE)
realpro = []
for i in PRODUCT_CODE:
    realpro.append(prodict[i].replace(' ','\n')) 
#print(realpro)


# In[3]:


def excel_list(excelfile,sheet,n):
    workbook = xlrd.open_workbook(excelfile)
    sheet_names= workbook.sheet_names()
    sheet = workbook.sheet_by_name(sheet)
    ncols = sheet.ncols
    for i in range(ncols):
        ldata = sheet.col_values(n) 
    ldata.pop(0)
    if '' in ldata:
        ldata.remove('')
    print(ldata)
    return ldata
excel_list(u'保湿模板所需数据1.xlsx','分析用list',3)


# In[4]:


from scipy import stats
ff = [4,6,3,14,4,99]
##检验是否正态
def norm_test(data):
    t,p =  stats.shapiro(data)
    #print(t,p)
    if p>=0.05:
        return True
    else:
        return False
print(norm_test(ff))


# In[5]:


data = pd.read_excel('E:\desktop\实验室数据标准处理\C182206-原始数据-最终.xlsx', sheetname=3)
#print(data)
an = excel_dict(u'保湿模板所需数据1.xlsx','第一个分析所用数据')
anlist = an[1]#['PRODUCT_CODE','CORNEOMETER_BL_AVERAGE','CORNEOMETER_T1H_AVERAGE','CORNEOMETER_T4H_AVERAGE','CORNEOMETER_T8H_AVERAGE']
andict =an[0] #{'PRODUCT_CODE':'产品名称','CORNEOMETER_BL_AVERAGE':'基础值(Baseline)','CORNEOMETER_T1H_AVERAGE':'1小时后(T1h)','CORNEOMETER_T4H_AVERAGE':'4小时后(T4h)','CORNEOMETER_T8H_AVERAGE':'8小时后(T8h)'}
print(anlist)
def get_serror():
#读取第一个待求列，计算均值、标准误
    data_group_b = data.groupby(['PRODUCT_CODE'])[anlist[0]]
    data_b = data[anlist[0]][data.PRODUCT_CODE=='B']
    data_mean =data_group_b .mean()
    data_std = data_group_b.std()
    data_size = data_group_b.size()
    data_error = data_std/data_size.map(lambda x: x**0.5)
    data_result = data_mean.map(lambda x: str('%.2f' % x)+'±')+data_error.map(lambda x: str('%.2f' % x))
    #print(data_mean)
    #print(data_result)
    #读取剩下的待求列，计算均值、标准误
    for i in range(1,len(anlist)):
        data_group_p = data.groupby(['PRODUCT_CODE'])[anlist[i]]
       # print(type(data_group_p))
        data_p = data[anlist[i]][data.PRODUCT_CODE=='B']
        t,p=ttest_rel(list(data_b),list(data_p))
        #print(p)
        data_mean1 =data_group_p .mean()
        data_std = data_group_p.std()
        data_size = data_group_p.size()
        data_error1 = data_std/data_size.map(lambda x: x**0.5)
        data_resultn = data_mean1.map(lambda x: str('%.2f' % x)+'±')+data_error1.map(lambda x: str('%.2f' % x))
        data_result = pd.concat([data_result,data_resultn],axis = 1)
        data_mean = pd.concat([data_mean,data_mean1],axis = 1)
        data_error = pd.concat([data_error,data_error1],axis = 1)
    data_result.rename(columns=andict, inplace=True)
    return data_result,data_mean,data_error
print(get_serror()[0])
writer = pd.ExcelWriter('output.xlsx')
get_serror()[0].to_excel(writer,'不同产品时间均值标准误')


# In[6]:


#PRODUCT_CODE = ['A','B','C','D','E']
columns = ['PRODUCT_CODE','对比时间点','N','t值','P值','显著性']
timelist = excel_list(u'保湿模板所需数据1.xlsx','分析用list',0)#['1小时 vs. 基础值','4小时 vs. 基础值','8小时 vs. 基础值']

def t_test():
    rank = []
    tlist= []
    plist = []
    nums = []
    times = []
    sigs = []
    for i in PRODUCT_CODE:
        data_b = data[anlist[0]][data.PRODUCT_CODE==i]
        #print(data_b)
        num = str(int(len(data_b)))
        for j in range(1,len(anlist)):
            times.append(timelist[j-1])
            data_p = data[anlist[j]][data.PRODUCT_CODE==i]
            #print(data_p)
            if norm_test(data_b) and norm_test(data_p):
                print('yes')
                t,p=ttest_rel(list(data_b),list(data_p))
            else:
                print('no')
                t,p=wilcoxon(list(data_b),list(data_p),zero_method='wilcox', correction=False)#
            p = '%.3f' %p
            t = '%.2f' %t
            #print(i,p)
            if float(p) >=0.05:
                sigs.append('n.s')
            else:
                sigs.append('***')
            rank.append(i)
            tlist.append(t)
            plist.append(p)
            nums.append(num)
    c={"PRODUCT_CODE" : rank, "t值" : tlist, "P值" : plist, "N" : nums,'对比时间点':times,'显著性':sigs}
    datacc=pd.DataFrame(c)#将字典转换成为数据框 
    print(datacc) 
    return datacc

t_test().to_excel(writer,'不同产品时间均值差异p值',columns=columns)            


# In[7]:


import matplotlib.pyplot as plt
from PIL import Image
from pylab import *  
mpl.rcParams['font.sans-serif'] = ['SimHei'] #指定默认字体  
  
mpl.rcParams['axes.unicode_minus'] = False #解决保存图像是负号'-'显示为方块的问题   
pro_list = ['产品A','产品B','产品C','产品D','空白对照']
time_list = excel_list(u'保湿模板所需数据1.xlsx','分析用list',1)#['基础值','使用后1小时','使用后4小时','使用后8小时']
#print(anlist)
data_mean = get_serror()[1]
data_error =  get_serror()[2]
#print(data_mean)
#print(data_mean[anlist[1]])
x =array(list(range(len(data_mean[anlist[0]]))))
total_width, n = 0.8, 4
width = total_width / n
x = x - (total_width - width) / 3
#print(x)
fig = plt.gcf()
fig.set_size_inches(18.5, 7)
error_kw = {'elinewidth' :2,'capsize' : 5,'C':'g'}#, 'fmt':"."
plt.bar(x, data_mean[anlist[0]], width=width,fc = '#5B9BD5', yerr=data_error[anlist[0]],ecolor = '#010000',error_kw=error_kw, label=time_list[0])#,
# plt.bar(x, data_mean[anlist[0]], width=width, yerr=data_error[anlist[0]],alpha=0.8,color="w",edgecolor="k",hatch=".",error_kw=error_kw, label=time_list[0])#,
#plot_sig(x,x+1,data_mean[anlist[0]],data_mean[anlist[0]]+1,'***')
#plt.errorbar(x ,  data_mean[anlist[1]])#,C = 'r'
colorlist = ['#5B9BD5','#ED7D31','#A5A5A5','#FFC000']
piclist = ['.','....','/','|','\\']
for j in range(1,len(time_list)):
   
    for i in range(len(x)):
        x[i] = x[i] + width
    #print(data_mean[anlist[1+j]])
    plt.bar(x, data_mean[anlist[j]], width=width, yerr=data_error[anlist[j]],fc = colorlist[j],ecolor = '#010000',error_kw=error_kw, label=time_list[j])
#     plt.bar(x, data_mean[anlist[j]], width=width, yerr=data_error[anlist[j]],alpha=0.8,color="w",edgecolor="k",hatch=piclist[j],error_kw=error_kw, label=time_list[j])
    #plot_sig(x,x+1,data_mean[anlist[j]],data_mean[anlist[j]]+1,'***')
    #plt.errorbar(x ,data_mean[anlist[1+j]], yerr=data_error[anlist[1+j]], fmt=".",ecolor = '#010000',elinewidth = 2,capsize = 5,C='g')#
plt.xticks(x-0.3, realpro, rotation=0,fontsize=15)
plt.yticks(fontsize=20)
plt.ylabel('皮肤水分值-使用前后差异值',fontsize=15)
plt.title("皮肤角质层水分含量",y=0.9,fontsize=20)
plt.legend(fontsize=15)
# plt.gca().xaxis.set_major_locator(plt.NullLocator())
# plt.gca().yaxis.set_major_locator(plt.NullLocator())

plt.subplots_adjust(top = 1, bottom = 0.1, right = 0.99, left = 0.05, hspace = 0, wspace = 0)
plt.margins(0,0.1)

plt.savefig('皮肤水分值-使用前后差异值.jpg',dpi=500)
plt.show()


# In[8]:


#时间点之间的差异值
an = excel_dict(u'保湿模板所需数据1.xlsx','第二个分析所用数据')
newminus = an[1]#['time1_baseline','time4_baseline','time8_baseline']
newminusdic = an[0]#{'time1_baseline':'1小时后(T1h)-基础值','time4_baseline':'4小时后(T4h)-基础值','time8_baseline':'8小时后(T8h)-基础值'}
#print(newminus,newminusdic)
#print(anlist)
def get_time_serror():
#读取第一个待求列，计算均值、标准误
    #print(data)
    for j in range(len(newminus)):
        data[newminus[j]] = data[anlist[1+j]]-data[anlist[0]]
    #print(data)
   # print(len(data))
    data_group_b = data.groupby(['PRODUCT_CODE'])[newminus[0]]
    #data_b = data[newminus[0]][data.PRODUCT_CODE=='A']
    data_mean =data_group_b .mean()
    data_std = data_group_b.std()
    data_size = data_group_b.size()
    data_error = data_std/data_size.map(lambda x: x**0.5)
   # data_result = data_mean.map(lambda x: str('%.2f' % x)+'±')+data_error.map(lambda x: str('%.2f' % x))
    #print(data_result)
    data_result = data_mean
    data_result.rename(columns={'time1_baseline':'wuuuuu'}, inplace=True)
    data_mean.rename(columns={'time1_baseline':'wuuuuu'}, inplace=True)
    data_error.rename(columns={'time1_baseline':'wuuuuu'}, inplace=True)
    #print(data_result)
    #print(data_mean)
    #读取剩下的待求列，计算均值、标准误
    for i in range(len(newminus)):
        data_group_p = data.groupby(['PRODUCT_CODE'])[newminus[i]]
       # print(type(data_group_p))
        #data_p = data[newminus[i]][data.PRODUCT_CODE=='A']
        
        data_mean1 =data_group_p .mean()
        data_std = data_group_p.std()
        data_size = data_group_p.size()
        data_error1 = data_std/data_size.map(lambda x: x**0.5)
        data_resultn = data_mean1.map(lambda x: str('%.2f' % x)+'±')+data_error1.map(lambda x: str('%.2f' % x))
        data_result = pd.concat([data_result,data_resultn],axis = 1)
        data_mean = pd.concat([data_mean,data_mean1],axis = 1)
        data_error = pd.concat([data_error,data_error1],axis = 1)
    data_result.rename(columns=andict, inplace=True)
    #print(data_result)
    data_result.rename(columns=newminusdic, inplace=True)
    del data_result[0],data_mean[0],data_error[0]
    return data_result,data_mean,data_error
print(get_time_serror()[0])
get_time_serror()[0].to_excel(writer,'不同产品时间差均值标准误')     


# In[9]:


#newminus = ['time1_baseline','time4_baseline','time8_baseline']
#PRODUCT_CODE = ['A','B','C','D','E']
#columns = ['PRODUCT_CODE','对比时间点','N','P值','显著性']
def t_time_test():
    rank = []
    product = []
    tlist = []
    plist = []
    nums = []
    sigs = []
    for i in newminus:
        data_b = data[i][data.PRODUCT_CODE==PRODUCT_CODE[-1]]
        #print(data_b)
        num = str(int(len(data_b)))
        #print(num)
        for j in range(len(PRODUCT_CODE)-1):
            
            data_p = data[i][data.PRODUCT_CODE==PRODUCT_CODE[j]]
            #print(data_p)
            if norm_test(data_b) and norm_test(data_p):
                print('yes')
                t,p=ttest_rel(list(data_b),list(data_p))
            else:
                print('no')
                t,p=wilcoxon(list(data_b),list(data_p))
            #t,p=ttest_rel(list(data_b),list(data_p))
            p = '%.3f' %p
            t = '%.2f' %t
            #print(PRODUCT_CODE[j],i,p)
            if float(p) >=0.05:
                sigs.append('n.s')
            else:
                sigs.append('***')
            rank.append(newminusdic[i])
            product.append(PRODUCT_CODE[j])
            tlist.append(t)
            plist.append(p)
            nums.append(num)
    c={"对比时间点" : rank,"PRODUCT_CODE" : product, "t值" : tlist, "P值" : plist,'N':nums,'显著性':sigs}
    datacc=pd.DataFrame(c)#将字典转换成为数据框 
    datacc = datacc.sort_values(by="PRODUCT_CODE")
    print(datacc) 
    return datacc

t_time_test().to_excel(writer,'不同产品时间差均值差异p值',columns=columns)     


# In[10]:


#pro_list = ['产品A','产品B','产品C','产品D','空白对照']
#excel_list(u'保湿模板所需数据.xlsx','分析用list',1)
time_list = ['使用后1小时-基础值','使用后4小时-基础值','使用后8小时-基础值']
data_mean = get_time_serror()[1]
data_error =  get_time_serror()[2]
print(data_mean)
print(data_error)
x =array(list(range(len(data_mean.loc["B"]))))
total_width, n = 0.8, 5
width = total_width / n
x = x - (total_width - width) / 3
#print(x)
fig = plt.gcf()
fig.set_size_inches(18.5, 7)
error_kw = {'elinewidth' :2,'capsize' : 5,'C':'g'}#, 'fmt':"."
#print(data_mean.loc["A"],data_error.loc["A"])

plt.bar(x, data_mean.loc["B"], width=width,fc = '#5B9BD5', yerr=data_error.loc["B"],ecolor = '#010000',error_kw=error_kw, label=realpro[0])#,
# plt.bar(x, data_mean.loc["A"], width=width, yerr=data_error.loc["A"],alpha=0.8,color="w",edgecolor="k",hatch=".",error_kw=error_kw, label=realpro[0])#,
for i in range(len(x)):
    x[i] = x[i] + width
plt.bar(x, data_mean.loc["NT"], width=width,fc = '#ED7D31', yerr=data_error.loc["NT"],ecolor = '#010000',error_kw=error_kw, label=realpro[1])#,
# plt.bar(x, data_mean.loc["NT"], width=width, yerr=data_error.loc["NT"],alpha=0.8,color="w",edgecolor="k",hatch="///",error_kw=error_kw, label=realpro[1])#,

#plt.errorbar(x ,  data_mean[anlist[1]])#,C = 'r'
colorlist = ['#5B9BD5','#ED7D31','#A5A5A5','#FFC000','#4472C4']#PRODUCT_CODE
colorlist = colorlist[0:2]
print(colorlist)
# for j in range(1,len(pro_list)):
   
#     for i in range(len(x)):
#         x[i] = x[i] + width
#     #print(data_mean[anlist[1+j]])
#     print(x)
#     plt.bar(x, data_mean.loc[PRODUCT_CODE[j]], width=width, yerr=data_error.loc[PRODUCT_CODE[j]],fc = colorlist[j],ecolor = '#010000',error_kw=error_kw, label=realpro[j])
#     #plt.errorbar(x ,data_mean[anlist[1+j]], yerr=data_error[anlist[1+j]], fmt=".",ecolor = '#010000',elinewidth = 2,capsize = 5,C='g')#
plt.xticks(x-0.05, time_list, rotation=0,fontsize=15)
plt.yticks(fontsize=20)
plt.ylabel('皮肤水分值-使用后减去使用前的差值',fontsize=15)
plt.title("皮肤角质层水分含量",y=0.9,fontsize=20)
plt.legend(fontsize=15)
plt.subplots_adjust(top = 1, bottom = 0.1, right = 0.99, left = 0.05, hspace = 0, wspace = 0)
plt.margins(0,0.1)
plt.savefig('皮肤水分值-使用后减去使用前的差值.jpg',dpi=500)
plt.show()


# In[11]:


writer.save()


# In[135]:


import numpy as np                 #使用import导入模块numpy，并简写成np
import matplotlib.pyplot as plt    #使用import导入模块matplotlib.pyplot，并简写成plt
plt.figure(figsize=(8,4))          #设置绘图对象的宽度和高度
from PIL import Image
from pylab import *  
mpl.rcParams['font.sans-serif'] = ['SimHei'] #指定默认字体  
mpl.rcParams['axes.unicode_minus'] = False #解决保存图像是负号'-'显示为方块的问题   

# x = np.ones((2))
# y = np.arange(0.8,1.1,0.2)
# plt.plot(x,y,label="$y$",color="black",linewidth=1)

# x = np.arange(1,3.1,2)
# y = 1+0*x
# plt.plot(x,y,label="$y$",color="black",linewidth=1)

# x0 = 2
# y0=1
# plt.annotate(r'$***$', xy=(x0, y0), xycoords='data', xytext=(-15, +1),
#              textcoords='offset points', fontsize=16,color="red")
# x = np.ones((2))*3
# y = np.arange(0.8,1.1,0.2)

# x = np.ones((2))*1
# y = np.arange(0.8,1.1,0.2)
# plt.plot(x,y,label="$y$",color="black",linewidth=1)

plt.bar(x, data_mean.loc["A"], width=width,fc = '#5B9BD5', yerr=data_error.loc["A"],ecolor = '#010000',error_kw=error_kw, label=realpro[0])#,
import math
def plot_sig(xstart,xend,ystart,yend,sig):
    for i in range(len(xstart)):
        x = np.ones((2))*xstart[i]
        y = np.arange(ystart[i],yend[i],yend[i]-ystart[i]-0.1)
        plt.plot(x,y,label="$y$",color="black",linewidth=1)

        x = np.arange(xstart[i],xend[i]+0.1,xend[i]-xstart[i])
        y = yend[i]+0*x
        plt.plot(x,y,label="$y$",color="black",linewidth=1)

        x0 = (xstart[i]+xend[i])/2
        y0=yend[i]
        plt.annotate(r'%s'%sig, xy=(x0, y0), xycoords='data', xytext=(-15, +1),
                     textcoords='offset points', fontsize=16,color="red")
        x = np.ones((2))*xend[i]
        y = np.arange(ystart[i],yend[i],yend[i]-ystart[i]-0.1)
        plt.plot(x,y,label="$y$",color="black",linewidth=1)
        #plt.ylim(0,math.ceil(max(yend)+4))             #使用plt.ylim设置y坐标轴范围
    #     plt.xlim(math.floor(xstart)-1,math.ceil(xend)+1)
        #plt.xlabel("随便画画")         #用plt.xlabel设置x坐标轴名称
        '''设置图例位置'''
        #plt.grid(True)
    plt.show()
plot_sig([0.42,1.42],[1.42,2.42],[30,20],[30.8,20.8],'***')
# plot_sig(1.42,2.42,20,20.8,'***')

