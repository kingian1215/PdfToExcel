import pdfplumber
import xlwt

workbook = xlwt.Workbook()
sheet = workbook.add_sheet('Sheet1')
i = 0
pdf = pdfplumber.open("PDFtoEXCEL.pdf")
print('開始讀取數據')
for page in pdf.pages:
    print(page)
    for table in page.extract_tables():
        print(table)
        for row in table:
            for j in range(len(row)):
                sheet.write(i, j, row[j])
            i += 1
        print("=============")
pdf.close()
workbook.save('result.xls')

print('保存成功！')