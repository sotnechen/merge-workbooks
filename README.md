# merge-workbooks
"""将多个sheet合并成一个sheet。这里的场景是各个sheet字段都是相同的类型的的情景，也不需要再额外进行数据处理"""
merge workbooks as many as possible
import pandas as pd #导入pandas库
zx=pd.ExcelFile(r'c:\users\stonelee\Desktop\channel商品评论.xlsx')#获取工作薄里面的属性
data=zx.parse(zx.sheet_names)#调用属性中的所有sheet名称并将数据传入变量data
data=pd.concat(data)#合并变量中的所有表组成新的DataFrame
data.to_excel(r'C:\Users\stonelee\Desktop\data.xlsx',index=False)#输出excel文件到桌面，不展示索引

"""将一个sheet拆分成多个sheet。继续使用上面合并的数据，现在根据日期这个维度进行一表拆成多表"""
import pandas as pd
data=pd.read_excel(r'C:\Users\stonelee\Desktop\data.xlsx')#导入数据
data['日期']=pd.to_datetime(data['时间']).dt.date#获取表中的日期部分并行增加一列日期
data_excel=[]#建一个用于储存多个sheet的空集
sheetname=[]#建一个用于储存多个sheet名称的空集
for x in data.groupby('日期')#根据日期字段进行分组
    data_excel.append(x[1])#将拆分的sheet存储到data_excel里面
    sheetname.append(x[0])#将拆分的sheet名称储存到sheetname里面
writer=pd.ExcelWriter(r'C:\Users\stonelee\Desktop\data1.xlsx')#定义一个最终文件存储的对象，防止覆盖
for i in range(len(sheetname)):#创建一个循环将多个sheet输出
    data_excel[i].iloc[:,0:9].to_excel(writer,sheet_name=str(sheetname[i]),index=False)
    #循环将多个sheet表中的数据及对应的sheet表名称输出至桌面，并且不展示索引
    
'''将一个工作簿拆分成多个'''
思路和一个工作sheet拆分工作sheet一样，只是随后输出的时候输出多个excel而不是多个sheet
import pandas as pd
data=pd.read_excel(r'c:\Users\stonelee\Desktop\data.xlsx'
data['日期']=pd.to_datetime(data['时间']).dt.date
data_excel=[]
sheetname=[]
for x in data.groupby('日期')
    data_excel.append(x[1])   
    sheetname.append(x[0])
for i in range(len(sheetname)):#区别在于循环创建多个路径，路径中加入变量工作表名称
    data_excel[i].iloc[:,0:9].to_excel（(r"C:\Users\stonelee\Desktop\data\\"+str(sheetname[i])+".xlsx")
 #桌面新建了一个data文件夹，将拆分的工作簿输出到这里
 
'''将多个工作簿合并一个。思路：先获取所有工作簿将数据导入，然后进行数据合并，最后输出成一个。区别在于如何获取所有工作簿的路径'''
import pandas as pd
import os #用到一个新库
op=r'c:\Users\stonelee\Desktop\data\\'定义一下数据存放的文件夹路径
name_list=os.listdir(op) #用os库获取该文件夹下的文件名称
data=[]
for x in range(len(name_list)):
    df=pd.read_excel(op+name_list[x])
    data.append(df) #将每个excel写入到data变量中
data=pd.concat(data) #合并data变量，转化成DataFrame
data.to_excelr'c:\Users\stonelee\Desktop\data3.xlsx',index=False #输出合并后的excel


