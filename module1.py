def txtexcel():
    import os
    os.chdir(r'C:\Users\86136\source\repos')
    #转到对应目录

    #csv文件另存成xls文件，感谢前辈@张行之，
    #这段代码的原文网址在https://blog.csdn.net/qq_33689414/article/details/78307031
    import csv
    import xlwt
    import win32com.client as win32

    csvFile = open("data2.csv",'w',newline='',encoding='utf-8')
    writer = csv.writer(csvFile)
    csvRow=[]

    f = open("Vt of temp (2).txt",'r',encoding='GB2312')
    for line in f:
        csvRow=line.split()
        writer.writerow(csvRow)
    f.close()
    csvFile.close()
    with open('data2.csv', 'r', encoding='Shift-JIS') as f:
        # 由于这个csv文件中的字库是Shift-JIS，这里用了个encoding='Shift-JIS'的参数来指定使用Shift-JIS字库。
        read = csv.reader(f)
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('data')
            # 创建一个名叫data的sheet（原作者的备注）
        l = 0
        for line in read:
            print(line)
            r = 0
            for i in line:
                print(i)
                sheet.write(l, r, i)
                    # 一个一个将单元格数据写入（原作者的备注）
                r = r + 1
            l = l + 1

        workbook.save('data2.xls') 
            # 保存Excel文件


    #csv文件另存成xls文件，感谢前辈@张行之，
    
    #这段代码的原文网址在https://blog.csdn.net/qq_33689414/article/details/78307031
    csvFile = open("data1.csv",'w',newline='',encoding='utf-8')
    writer = csv.writer(csvFile)
    csvRow=[]

    f = open("Vt of temp.txt",'r',encoding='GB2312')
    for line in f:
        csvRow=line.split()
        writer.writerow(csvRow)
    f.close()
    #csvFile.close()
    #def csv_to_xlsx1():
    with open('data1.csv', 'r', encoding='Shift-JIS') as f:
    # 由于这个csv文件中的字库是Shift-JIS，这里用了个encoding='Shift-JIS'的参数来指定使用Shift-JIS字库。
        read = csv.reader(f)
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet('data')
        # 创建一个名叫data的sheet（原作者的备注）
        l = 0
        for line in read:
            print(line)
            r = 0
            for i in line:
                print(i)
                sheet.write(l, r, i)
                # 一个一个将单元格数据写入（原作者的备注）
                r = r + 1
            l = l + 1
        workbook.save('data1.xls')
    #        # 保存Excel文件
    #if __name__ == '__main__':
    ##xls文件转换成xlsx文件
    #    csv_to_xlsx1()

    


    ##xls文件另存成xlsx文件
