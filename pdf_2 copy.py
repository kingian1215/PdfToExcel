from PIL import Image
import numpy as np
import os
import pytesseract 
import cv2 as cv
import fitz
import xlwt

workbook = xlwt.Workbook()
i = 1

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


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

        imgPath2 = imgPath+str(pag)+ ".png"
        print(imgPath2)
        # img=cv.imread(imgPath)
        img = cv.imread(imgPath2)

        # 將圖片轉成灰階，以利邊緣偵測
        gray = cv.cvtColor(img,cv.COLOR_BGR2GRAY)
        # 用openCV的Canny函式去做邊緣偵測
        edges = cv.Canny(gray,100,150,apertureSize = 3)
        # 用openCV的HoughLines函式找出所有構成邊緣線段的4個點座標(x1,y1,x2,y2)
        minLineLength = 10000
        maxLineGap = 100
        lines = cv.HoughLinesP(edges,0.01,np.pi/180,350,minLineLength,maxLineGap,2500)
        # for i in lines:
        #  #將線段的四個點座標分別assign到x1,y1,x2,y2
            # x1,y1,x2,y2 = i[0]
        #  #利用openCV的line函式，對 img 化上顏色為綠色、粗度為1pixel的線段
            # cv.line(img,(x1,y1),(x2,y2),(0,255,0),1)

        verti_lst = []
        horiz_lst = []
        for k in lines:
        #  將線段的四個點座標分別assign到x1,y1,x2,y2
            x1,y1,x2,y2 = k[0]
        
        #  計算斜率
            slope = (y2-y1)/(x2-x1)
        
        #  判斷斜率篩選出垂直線與水平線
            if slope == -1*np.inf:
                cv.line(img,(x1,y1),(x2,y2),(0,255,0),1)
        
        #  將垂直線的 x1座標紀錄在verti_lst裡
                verti_lst.append(x1)
        
            if slope == 0:
                cv.line(img,(x1,y1),(x2,y2),(0,255,0),1)
        #  將水平線的 y1座標紀錄在horiz_lst裡
                horiz_lst.append(y1)
        print(reselect_y_coordinates(sorted(horiz_lst)))
        print(reselect_x_coordinates(sorted(verti_lst)))
         # #將結果存下來，檔名為test.png
        # cv.imwrite('test'+ str(pag)+ ".png", img)
        # cv.waitKey(0)

        # img = cv.imread(image_name)
        for i1 in range(1, len(horiz_lst)-1):
            for j1 in range(1, len(verti_lst)-1):
                crop_img =/ img[horiz_lst[i1]:horiz_lst[i1+1],verti_lst[j1]:verti_lst[j1+1]]
                cv.imwrite(str(i1)+ "-"+ str(j1)+ ".png", crop_img)

        # text = pytesseract.image_to_string(Image.fromarray(img),lang="chi_tra+eng")
        text = pytesseract.image_to_string(Image.fromarray(img), lang="chi_tra")
        listtxt = text.splitlines(keepends = True)
        pp = str(i)
        # sheetpp = 'Sheet'+ pp
        # sheet = workbook.add_sheet('Sheet'+ pp, cell_overwrite_ok=True)

        # for j in range(0, len(listtxt)) :
            # sheet.write(j, 0, listtxt[j])

        # i += 1
        os.remove(imgPath2)

    # workbook.save(imgPath+ '.xls')
    pdf.close()

def reselect_y_coordinates(coordinates_lst):
    coordinates_lst = sorted(coordinates_lst)
    bag = []
    result = []
    for i in coordinates_lst:
        if len(bag) == 0:
            bag.append(i)
        else:
            if i-bag[-1] ==1:               #//判斷連號
                bag.append(i)
            else:
                if len(bag) >=5:            #//若連號的數量>=5
                    result.append(bag[-1])  #//把最後一號記錄下來
                bag.clear()
                bag.append(i)
    return result

def reselect_x_coordinates(coordinates_lst):
    coordinates_lst = sorted(coordinates_lst)
    bag = []
    result = []
    for i in coordinates_lst:
        if len(bag) == 0:
            bag.append(i)
        else:
            if i-bag[-1] ==1:
                bag.append(i)
            else:
                if len(bag) >=3:
                    result.append(bag[-1])
                bag.clear()
                bag.append(i)
    return result

# 依賴opencv
# img=cv.imread(img_path)
# text=pytesseract.image_to_string(Image.fromarray(img),lang='chi_tra')
# 不依賴opencv寫法
# text=pytesseract.image_to_string(Image.open(img_path))
# print(text)

pdf_path ='PDFtoEXCEL.pdf'
img_path ='PDFtoEXCEL_ch'
pdf_image(pdf_path,img_path, 5, 5, 0, i)
