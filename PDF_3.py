from skimage import io, data, morphology, segmentation
import numpy as np
from PIL import Image
import pytesseract
import os

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# pytesser3.tesseract_exe_name="C:/Program Files (x86)/Tesseract-OCR/tesseract.exe"
#識別圖片中的文字
def img_to_txt(file):
    if os.path.exists(file):
        image = Image.open(file)
        # 英文
        # vcode = pytesseract.image_to_string(image,"eng")
        # txt = pytesser3.image_to_string(image, "chi_sim")
        txt = pytesseract.image_to_string(image, "chi_tra")
        print("{}識別內容:{}".format(file, txt))

for i in range(16):
    img_file = r"img/pic_{}.jpg".format(i)
    img_to_txt(img_file)
    img_to_txt(r"img/abc.png")
    
    img = io.imread("test.jpg", True) # as_grey=true 灰度圖片
    
    #二值化
    bi_th=0.81
    img[img <= bi_th]=0
    img[img > bi_th]=1
io.imsave('gray.jpg', img)
    
# 膨脹腐蝕操作
def dil2ero(img, selem):
    img = morphology.dilation(img, selem) # 膨脹
    imgres = morphology.erosion(img, selem) # 腐蝕
    return imgres

# 求圖像中的橫線和豎線
rows, cols = img.shape
scale = 10 #這個值越大,檢測到的直線越多,圖片中字體越大，值應設置越小。需要靈活調整該參數。

col_selem = morphology.rectangle(cols//scale, 1)
img_cols = dil2ero(img, col_selem)

row_selem = morphology.rectangle(1, rows//scale)
img_rows = dil2ero(img, row_selem)

# 線圖
img_line = img_cols* img_rows
io.imsave('table.jpg', img_line)

# 點圖
img_dot = img_cols + img_rows
img_dot[img_dot>0] = 1

io.imsave('table_dot.jpg', img_dot)

# 收縮點團爲單像素點（3×3）
def isolate(imgdot):
    idx = np.argwhere(imgdot<1) # img值小於1的索引數組    
    rows, cols = imgdot.shape

    for i in range(idx.shape[0]):
        c_row = idx[i, 0]
        c_col = idx[i, 1]
        if c_col+1<cols and c_row+1<rows:
            imgdot[c_row, c_col+1] = 1
            imgdot[c_row+1, c_col] = 1
            imgdot[c_row+1, c_col+1] = 1

        if c_col+2<cols and c_row+2<rows:
            imgdot[c_row+1, c_col+2] = 1
            imgdot[c_row+2, c_col] = 1
            imgdot[c_row, c_col+2] = 1
            imgdot[c_row+2, c_col+1] = 1
            imgdot[c_row+2, c_col+2] = 1
    return imgdot

img_dot = isolate(img_dot)

io.imsave('table_dot_del.jpg', img_dot)

# io.imshow(img)
# io.imsave('1234.jpg', img)
print(type(img)) #顯示類型
print(img_dot.shape) #顯示尺寸
print(img_dot.shape[0]) #圖片高度
print(img_dot.shape[1]) #圖片寬度

# print(dot_idxs.size)
for m in range(img_dot.shape[0]):
    if m > 100: break
        print('col{}'.format(m),end=" ")

    for n in range(img_dot.shape[1]):
        if img_dot[m][n]==0 :
            print(img_dot[m][n], end= " ")
        print("",end="\n")
    dot_idxs=np.argwhere(img_dot<1) # img_dot值等於0的索引數組
    table_cols = [] #記錄每行有幾個單元格
    table_rows = 1
    table_row_index = dot_idxs[0][0] #第一行點圖y值座標
    colu_dot = 0 # 單元格數量

for n,indx in enumerate(dot_idxs):
    if indx[0] == table_row_index : # 如果點圖y值座標相同則爲同一行
        colu_dot = colu_dot + 1
        print(indx,end=" ")
    else :
        table_cols.append(colu_dot-1) # 將行單元格數量加入table_cols列表
        table_row_index = dot_idxs[n+1][0] #記錄下一行位置索引
        colu_dot = 1 # 清0後，換行記錄單元格數量
        print("", end="\n")
    print(indx, end=" ")

print("")
for n,v in enumerate(table_cols):
    print('row{}:{}'.format(n,v))
    row_num = len(table_cols)
    max_row = max(table_cols)
    min_row = min(table_cols)
    cell_nums = sum(table_cols)
print('這個表格一共{}行'.format(row_num))
print('這個表格最多行單元格數：{}'.format(max_row))
print('這個表格最少行單元格數：{}'.format(min_row))
print('這個表格一共{}個單元格'.format(cell_nums))

go = input('表格分析是否正確？')

if go == 'no' :
    exit(0)

# 開始識別單元格
a_dots =[] #記錄已經處理過的A點座標
cells = [] # 記錄單元格的A點和D點座標 [[a11,a12,d11,d12][a21,a22,d21,d22]]]
cell_num = 0
cell_a = []
cell_b = []
cell_c = []
cell_d = []
if min_row == min_row : # 標準表格 n*n
    for n,v in enumerate(table_cols):
        print('row{}:'.format(n),end= ' ')
        for i in range(v):
            row_dot_num = v + 1
            cell_num = n * row_dot_num
            cell_a_index = cell_num + i
            cell_b_index = cell_a_index + 1
            cell_c_index = cell_a_index + row_dot_num
            cell_d_index = cell_c_index +1
            print('cell[{}]:a{} b{} c{} d{}'.format(cell_num+i, dot_idxs[cell_a_index], 
                dot_idxs[cell_b_index], dot_idxs[cell_c_index], dot_idxs[cell_d_index]), end= ' ')
        x1=dot_idxs[cell_a_index][1]+1 # 加1去除列邊框線
        x2=dot_idxs[cell_b_index][1]
        y1=dot_idxs[cell_a_index][0]+1 # 加1去除行邊框線
        y2=dot_idxs[cell_c_index][0]
        #print(x1,x2,y1,y2)
        roi=img[y1:y2,x1:x2] # 50~100 行，50~100 列（不包括第 100 行和第 100 列）
        io.imsave('img/pic_{}.jpg'.format(cell_num+ i), roi)
        print("", end="\n")