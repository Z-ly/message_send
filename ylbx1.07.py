 # -*- coding: utf-8 -*-
 #author ：zhanglinyi
 #version：1.07


from tkinter import *
from tkinter import messagebox
import xlrd
import os,sys
from PIL import ImageTk
import time
top = Tk()
top.title('医疗报销短信生成 1.07')
exc = StringVar()
exc.set('xxxx年xx月公费医疗报销.xlsx')
nameplace = IntVar()
nameplace.set(3)
moneyplace = IntVar()
moneyplace.set(6)
namehang = IntVar()
namehang.set(4)
biao = IntVar()
biao.set(1)
stu = '同学'
tec = '老师'

def message():
    title = exc.get()
    print(title)
    name_place = nameplace.get()
    money_place = moneyplace.get()
    name_hang = namehang.get()
    biao_ = biao.get()
    
    if biao_ == 1:
        sf = stu
    elif biao_ == 2:
        sf = tec
    if (title in os.listdir(os.getcwd())):
        data = xlrd.open_workbook(title)
        table = data.sheets()[biao_ - 1]
        if('姓名' not in table.col_values(name_place-1)[name_hang-2]):
            messagebox.showinfo(title='warnning', message='名字列错了 ' )
        if('金额' not in table.col_values(money_place-1)[name_hang-2]):
            messagebox.showinfo(title='warnning', message='金额列错了 ' )
        a = table.col_values(name_place-1)
        b = table.col_values(money_place-1)


        #date = table.cell(2,5).value
        date =(str)(xlrd.xldate.xldate_as_datetime(table.cell(1,5).value,0).strftime("%Y-%m-%d"))


        #c = table.col_values(money_place+1)
        try :
            c = table.col_values(money_place+1)
        except Exception:
            ycfh = True                                                         #有异常返回
        else:
            ycfh = False
        
        f = open('短信.txt','w')
        #file = docx.Document()
        '''for i in range(name_hang,len(a)-1):
            #k = str(b[i])
            if (b[i]) == 0:
                #print('cuowu')
                #c = table.col_values(money_place+1)

                f.write('【学生事务中心】'+a[i]+ sf +'您好，您本月通过京工飞鸿代办医药费报销业务异常返回，返回原因:'+c[i]+'。请至中心教学楼129补全材料。学生事务中心京工飞鸿将竭诚为您服务！'"\n")
                continue
            else:    
                d = '%.2f'%float(b[i])
                f.write('【学生事务中心】'+a[i]+ sf +'您好，您本月通过京工飞鸿代办医药费报销业务已完成，报销金额为 '+d+' 元，3月底前发放报销款项，请耐心等待。如对报销金额有疑问请本人于7日内携证件到校医院财务室领取。学生事务中心京工飞鸿将竭诚为您服务！代办日期：'+date+"\n")
               #print(d)'''
        if ycfh:
            for i in range(name_hang-1,len(a)-1):
                d = '%.2f'%float(b[i])
                f.write('【学生事务中心】'+a[i]+ sf +'您好，您本月通过京工飞鸿代办医药费报销业务已完成，报销金额为 '+d+' 元，3月底前发放报销款项，请耐心等待。如对报销金额有疑问本人于7日内携证件到校医院财务室领取。学生事务中心京工飞鸿将竭诚为您服务！代办日期：'+date+"\n")
            
        else:
            for i in range(name_hang-1,len(a)-1):
                d = '%.2f'%float(b[i])
                if c[i]=='':
                   
                    f.write('【学生事务中心】'+a[i]+ sf +'您好，您本月通过京工飞鸿代办医药费报销业务已完成，报销金额为 '+d+' 元，3月底前发放报销款项，请耐心等待。如对报销金额有疑问请本人于7日内携证件到校医院财务室领取。学生事务中心京工飞鸿将竭诚为您服务！代办日期：'+date+"\n")
                elif float(b[i]) == 0:
                     f.write('【学生事务中心】'+a[i]+ sf +'您好，您本月通过京工飞鸿代办医药费报销业务异常返回，返回原因:'+c[i]+'。请至中心教学楼129补全材料。学生事务中心京工飞鸿将竭诚为您服务！代办日期：'+date+"\n")
                else:
                     f.write('【学生事务中心】'+a[i]+ sf +'您好，您本月通过京工飞鸿代办医药费报销' + d + ' 元，报销业务部分异常返回，返回原因:'+c[i]+'。请至中心教学楼129补全材料。学生事务中心京工飞鸿将竭诚为您服务！代办日期：'+date+"\n")
                





        f.close()
        messagebox.showinfo(title='exciting', message='mission completed' )
    else:
        messagebox.showinfo(title='warnning', message='眼睛睁大点，没找到文件' )
        
    









top.geometry('500x550')

top.resizable(width=False, height=False)

#__bg__
canvas = Canvas(top,width = 500, height = 600, bg = 'WHITE')
#canvas.pack(expand = YES, fill = BOTH) 
canvas.pack()
image = ImageTk.PhotoImage(file = r'C:\Users\jinggongfeihong1\Desktop\短信生成\xiugai')
canvas.create_image(0, -10, image = image, anchor = NW)



#__视图__

l = Label(top, text = '短信生成-bug版').place(x = 180,y = 10)


l1 = Label(top, text = 'excel文件名').place(x = 50,y = 50) 
l2 = Label(top, text = '申请人列数').place(x = 50,y = 100)
l3 = Label(top, text = '报销金额列数').place(x = 50,y = 150)
l4 = Label(top, text = 'F为第6列').place(x = 180,y = 150)
l5 = Label(top, text = '申请人从第').place(x = 180,y = 100)
l6 = Label(top, text = '行开始').place(x = 265,y = 100)
l7 = Label(top, text = '表单号      （一般学生在第一张表，老师在第二张表）').place(x = 50,y = 200) 



exc.set('xxxx年xx月公费医疗报销.xlsx')
w1 = Entry(top, textvariable = exc , bd = 2, width = 30).place(x = 150,y = 50)
w2 = Entry(top, textvariable = nameplace , bd = 2, width = 3).place(x = 150,y = 100)
w3 = Entry(top, textvariable = moneyplace , bd = 2, width = 3).place(x = 150,y = 150)
w4 = Entry(top, textvariable = namehang, width = 1, bd = 2).place(x = 248,y = 100)
w5 = Entry(top, textvariable = biao, width = 1, bd = 2).place(x = 100,y = 200)


btn = Button(top,text = '生成短信', command = message).place(x= 150,y = 230)

def shuoming():
     messagebox.showinfo(title='warnning', message='将ecxel拖入.exe文件根目录，按表格内容修改读取的表单、行数以及列数，部分报销的金额照常填写，未完成报销的金额请填0，返单原因（包括部分报销）请一定写在excel表格报销金额后面第二列相应位置。' )
    
    

btn1 = Button(top,text = '使用说明', command = shuoming).place(x= 250,y = 230)




top.mainloop()
