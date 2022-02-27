from PIL import Image
import os
import pytesseract 
import cv2 as cv
import fitz
import xlwt

workbook = xlwt.Workbook()
# sheet = workbook.add_sheet('Sheet1')
i = 1

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# def pdf_image(pdfPath,imgPath,zoom_x,zoom_y,rotation_angle):
#   打開PDF文件
    # pdf = fitz.open(pdfPath)
    # 逐頁讀取PDF
    # for pg in range(0, pdf.pageCount):
        # page = pdf[pg]
        # 設置縮放和旋轉系數
        # trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotation_angle)
        # pm = page.getPixmap(matrix=trans, alpha=False)
        # 開始寫圖像
        # pm.writePNG(imgPath+str(pg)+".png")
        # pm.writePNG(imgPath)
    # pdf.close()
    # 
# pdf_path ='PDFtoEXCEL.pdf'
# img_path ='PDFtoEXCEL'
# pdf_image(pdf_path,img_path,5,5,0)

def pdf_image(pdfPath, imgPath, zoom_x, zoom_y, rotation_angle, i):
    pdf = fitz.open(pdfPath)

    for pag in range(0, 2):
    # for pag in range(0, pdf.pageCount):
        page = pdf[pag]
            # 設置縮放和旋轉系數
        trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotation_angle)
        pm = page.getPixmap(matrix=trans, alpha=False)
        # 開始寫圖像
        pm.writePNG(imgPath+str(pag)+ ".png")
        #pm.writePNG(imgPath)

        # imgPath2 = imgPath+str(pag)+ ".png"
        print(imgPath2)
        # img=cv.imread(imgPath)
        img = cv.imread(imgPath2)
        # text = pytesseract.image_to_string(Image.fromarray(img),lang="chi_tra+eng")
        text = pytesseract.image_to_string(Image.fromarray(img), lang="chi_tra")
        listtxt = text.splitlines(keepends = True)
        pp = str(i)
        sheetpp = 'Sheet'+pp
        sheet = workbook.add_sheet('Sheet'+pp, cell_overwrite_ok=True)

        for j in range(0, len(listtxt)) :
            sheet.write(j, 0, listtxt[j])

        i += 1
        os.remove(imgPath2)

    workbook.save(imgPath+ '.xls')
    pdf.close()

    # 依賴opencv
# img=cv.imread(img_path)
# text=pytesseract.image_to_string(Image.fromarray(img),lang='chi_tra')
# 不依賴opencv寫法
# text=pytesseract.image_to_string(Image.open(img_path))
# print(text)

pdf_path ='PDFtoEXCEL.pdf'
img_path ='PDFtoEXCEL_ch'
pdf_image(pdf_path,img_path, 5, 5, 0, i)