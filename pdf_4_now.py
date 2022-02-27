import cv2
import numpy as np
from PIL import Image
import os
import pytesseract 
import fitz
import xlwt

workbook = xlwt.Workbook()
i = 1

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


def showImage(winname, img):

    cv2.imshow(winname, img)

    cv2.waitKey(0)

    cv2.destroyAllWindows()


#删除表格的横线和竖线，但是保留表格内容

def removelines(imageurl):

    image = cv2.imread(imageurl)
    # image = image0.resize((256, 256))

    result = image.copy()

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

    # 移除水平线

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (500,5))

    remove_horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)

    cnts = cv2.findContours(remove_horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    cnts = cnts[0] if len(cnts) == 2 else cnts[1]

    for c in cnts:

        #通过画白线来清除水平线

        cv2.drawContours(result, [c], -1, (255,255,255), 3)


    # 移除垂直线

    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5,100))

    remove_vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)

    cnts = cv2.findContours(remove_vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    cnts = cnts[0] if len(cnts) == 2 else cnts[1]

    for c in cnts:

        #通过画白线来清除垂直线

        cv2.drawContours(result, [c], -1, (255,255,255), 3)


    # showImage('result', result)

#提取表格

def extracttable(imageurl):

    img = cv2.imread(imageurl)
    # img = img0.resize((256, 256))
    # img.show()

    # gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)


    if len(img.shape) != 2:

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    else:

        gray = img


    gray = cv2.bitwise_not(gray)
    # 
    # gray = cv2.medianBlur(gray)
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 5, 0)

    

    # 创建图像副本

    horizontal = np.copy(bw)

    vertical = np.copy(bw)


    cols = horizontal.shape[1]
    # print("cols=", cols)

    horizontal_size = int(cols /7)
    # print("horizontal_size =", horizontal_size)

    # 定义结构元素

    horizontalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (horizontal_size, 1))


    # 图像形态学操作 腐蚀、膨胀

    horizontal = cv2.erode(horizontal, horizontalStructure)

    horizontal = cv2.dilate(horizontal, horizontalStructure)


    rows = vertical.shape[0]

    verticalsize = int(rows /7)


    # 定义结构元素

    verticalStructure = cv2.getStructuringElement(cv2.MORPH_RECT, (1, verticalsize))


    # 图像形态学操作 腐蚀、膨胀

    vertical = cv2.erode(vertical, verticalStructure)

    vertical = cv2.dilate(vertical, verticalStructure)


    mask = horizontal + vertical

    result = cv2.bitwise_not(mask)

    # showImage("", result)




def pdf_image(pdfPath, imgPath, zoom_x, zoom_y, rotation_angle, i):
    pdf = fitz.open(pdfPath)

    for pag in range(0, 10 ):
    # for pag in range(0, pdf.pageCount):
        page = pdf[pag]
            # 設置縮放和旋轉系數
        trans = fitz.Matrix(zoom_x, zoom_y).preRotate(rotation_angle)
        # Matrix(x, y縮放倍數) .preRotate(執行一個旋轉)
        pm = page.getPixmap(matrix=trans, alpha=False)
        # 開始寫圖像
        pm.writePNG(imgPath+str(pag)+ ".png")
        #pm.writePNG(imgPath)

        imgPath2 = imgPath+str(pag)+ ".png"
        # imgPath2 = 'PDFtoEXCEL_ch0_1.png'
        print(imgPath2)
        img = Image.open(imgPath2)
        # print(img.size)
        img.thumbnail((1024, 2048))
        # print(img.size)
        img.save(imgPath2)

        # img=cv.imread(imgPath)
        
        removelines(imgPath2)

        extracttable(imgPath2)

        img = cv2.imread(imgPath2)
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


pdf_path ='PDFtoEXCEL.pdf'
img_path ='PDFtoEXCEL'
pdf_image(pdf_path, img_path, 2, 2, 0, i)

# removelines('testimgnew/lines.png')

# extracttable('testimgnew/lines.png')