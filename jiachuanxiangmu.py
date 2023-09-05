
#删除不匹配的文件
workbookops=openpyxl.load_workbook(r'C:\Users\Administrator\Desktop\tablet\ops2.xlsx')
sheetops=workbookops.active
opslen=sheetops.max_row+1
opscol=sheetops.max_column
opslen
opscol
i=1
x=[]
while i<opslen:
    kk=sheetops.cell(i,(opscol-2)).value
    if kk==None:
        x.append(i)
    i+=1 
x
xlen=len(x)
i=1
while i<xlen+1:
    sheetops.delete_rows(x[xlen-i])
    i+=1
    i
workbookops.save(r'C:\Users\Administrator\Desktop\tablet\final.xlsx')
i
print('finsh')