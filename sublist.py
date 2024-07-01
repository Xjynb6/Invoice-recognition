import os
import cv2
import xlwt
import numpy as np
import xlrd
from PIL import Image
import pandas as pd
from paddleocr import PaddleOCR, draw_ocr
import re
import difflib

#识别部分
def hide(img,save):
    # 读取图片
    img = cv2.imread(img)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 使用Canny边缘检测算法检测直线
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    # 使用Probabilistic Hough Transform检测直线
    lines = cv2.HoughLinesP(edges, 1, np.pi/180, threshold=100, minLineLength=100, maxLineGap=1)
    # 将直线涂白
    for line in lines:
     x1, y1, x2, y2 = line[0]
     cv2.line(img, (x1, y1), (x2, y2), (255, 255, 255), 2)
     cv2.imwrite(save, img)

def judge(path):  # 判断是否有文件夹
    if not os.path.exists(path):
        os.mkdir(path)


def jpg(img_path, save_path):
    ocr = PaddleOCR(use_angle_cls=True, lang='ch')  # 用角度分析模型，语言为中文
    result = ocr.ocr(img_path, cls=True)
    image = Image.open(img_path).convert('RGB')  # 输出为rgb
    boxes = [line[0] for line in result[0]]  # 坐标
    txts = [line[1][0] for line in result[0]]  # 文字
    scores = [line[1][1] for line in result[0]]  # 可信度
    im_show = draw_ocr(image, boxes, txts, scores, font_path='./fonts/simfang.ttf')
    im_show = Image.fromarray(im_show)
    im_show.save(save_path)
    return boxes, txts


def Init(path):  # 删除上次留下的pdf图片
    files = os.listdir(path)
    for file in files:
        a = os.path.splitext(file)[1]
        if a == '.pdf' or a=='.xlsx' or a=='.png':
            A = os.path.join(path, file)
            os.chmod(A, 0o777)#赋予删除权限
            os.remove(A)

def shibie(totalPage,width,height,wjj,img_path,save_path):
    boxes, txts = jpg(img_path, save_path)
    # 创建表格
    #识别全部内容
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('识别结果_2')
    worksheet.write_merge(0, 0, 0, 3, '坐标')
    worksheet.write(0, 4, '文字')
    for e in range(len(txts)):
        worksheet.write(e + 2, 4, str(txts[e]))
    for i in range(len(boxes)):
        worksheet.write(i + 2, 0, str(boxes[i][0]))
        worksheet.write(i + 2, 1, str(boxes[i][1]))
        worksheet.write(i + 2, 2, str(boxes[i][2]))
        worksheet.write(i + 2, 3, str(boxes[i][3]))
    for i in range(5):
        worksheet.col(i).width = 256 * 20
    workbook.save(wjj + '/main/识别结果_2.xls')

    #提取截取的坐标
    data_excel = xlrd.open_workbook(wjj + '/main/识别结果_2.xls')
    sheet_object = data_excel.sheet_by_index(0)  # 获取第一张工作表的内容
    n_rows = sheet_object.nrows  # 获取行数
    height_0 = 0
    for i in range(n_rows):
        content = sheet_object.cell_value(i, 4)
        if (content == '序号'):
            a = sheet_object.cell_value(i, 2)  # [102.0,355.0]
            a = a.split(',')
            b = a[0].split('[')
            height_0 = b[1]
        if (content == '规格型号'):
            a = sheet_object.cell_value(i, 2)
            a = a.split(',')
            b = a[1].split(']')
            height_1 = b[0]  # 104
        if totalPage==2:
            if (content == '小计'):
                a = sheet_object.cell_value(i, 0)
                a = a.split(',')
                b = a[1].split(']')
                height_2 = b[0]  # 1220
        else:
            if (content == '合'):
                a = sheet_object.cell_value(i, 0)
                a = a.split(',')
                b = a[1].split(']')
                height_2 = b[0]  # 1220
    if totalPage==2:
        img = wjj + '/png/images_2.png'
    else:
        img=wjj + '/png/images_1.png'
    save = wjj + '/png/2.png'
    hide(img, save)

    img = cv2.imread(save)
    # pts1 需要截取图片的四点坐标 pts2 结果图坐标 matrix 存储pts2的像素
    width_q, height_q = int(float(width)*1.5),int((float(height_2)-float(height_1))*2)
    pts1 = np.float32([[height_0, height_1], [width, height_1], [height_0, height_2], [width, height_2]])
    pts2 = np.float32([[0, 0], [width_q, 0], [0, height_q], [width_q, height_q]])
    matrix = cv2.getPerspectiveTransform(pts1, pts2)
    imgoutput = cv2.warpPerspective(img, matrix, (width_q, height_q))
    cv2.imwrite(wjj + '/Output/Output1.png', imgoutput)

    #每张图片之间用！分割
    img_path = wjj + '/Output/Output1.png'
    save_path = wjj + '/Output/货物明细.jpg'
    boxes, txts = jpg(img_path, save_path)
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('y+名字')
    worksheet.write(0, 0, '坐标1')
    worksheet.write(0, 1, '坐标2')
    worksheet.write(0, 2, '名字')
    worksheet.write(0, 3, 'x坐标')
    a = boxes[0][0][1]
    b = 0
    for i in range(len(boxes)):
        if abs(boxes[i][0][1] - a) > 15:
            worksheet.write(b + 1, 1, '!')
            worksheet.write(b + 1, 0, '!')
            worksheet.write(b + 1, 2, '!')
            b = b + 1
            worksheet.write(b + 1, 2, str(txts[i]))
            worksheet.write(b + 1, 0, str(boxes[i][0][1]))
            worksheet.write(b + 1, 1, str(boxes[i][2][1]))
            worksheet.write(b + 1, 3, str(boxes[i][0][0]))
        else:
            worksheet.write(b + 1, 2, str(txts[i]))
            worksheet.write(b + 1, 0, str(boxes[i][0][1]))
            worksheet.write(b + 1, 1, str(boxes[i][2][1]))
            worksheet.write(b + 1, 3, str(boxes[i][0][0]))
        a = boxes[i][0][1]
        b = b + 1
    worksheet.write(b + 1, 0, '!')
    workbook.save(wjj + '/main/y+名字.xls')
    #提取货物名字列
    df = pd.read_excel(wjj + '/main/y+名字.xls')
    list1 = df['坐标1'].tolist()
    list1_1 = df['坐标2'].tolist()
    list2 = df['名字'].tolist()
    list2_1 = df['x坐标'].tolist()
    i = 0
    list3 = []
    list3_1 = []
    list4 = []  # 名字
    list4_1 = []  # x坐标
    start = 0
    while i < len(list1):
        if list1[i] == '!':
            if i - 1 == start:
                list3.extend([list1[start]])
                list3_1.extend([list1_1[start]])
                list4.extend([list2[start]])
                list4_1.extend([list2_1[start]])
            else:
                list3.extend([list1[start:i]])
                list3_1.extend([list1_1[start:i]])
                list4.extend([list2[start:i]])
                list4_1.extend([list2_1[start:i]])
            i = i + 1
            start = i
            continue
        i = i + 1
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('横版名字')
    worksheet.write(0, 0, '名字')
    worksheet.write(0, 1, '坐标1')
    worksheet.write(0, 2, '坐标2')
    row = 1
    for i in range(len(list4)):
        if type(list4[i]) == list:
            d1 = dict(zip(list4_1[i], list4[i]))
            d2 = dict(zip(list4_1[i], list3[i]))
            d3 = dict(zip(list4_1[i], list3_1[i]))
            list4_1[i] = [float(x) for x in list4_1[i]]
            list4_1[i].sort()
            worksheet.write(row, 0, d1[list4_1[i][0]])
            worksheet.write(row, 1, d2[list4_1[i][0]])
            worksheet.write(row, 2, d3[list4_1[i][0]])
        else:
            worksheet.write(row, 0, list4[i])
            worksheet.write(row, 1, list3[i])
            worksheet.write(row, 2, list3_1[i])
        row = row + 1
    workbook.save(wjj + '/main/横版名字.xls')

    # 只取第一列
    df = pd.read_excel(wjj + "/main/横版名字.xls")
    list3 = df['名字'].tolist()
    list4 = df['坐标1'].tolist()
    list4_1 = df['坐标2'].tolist()
    workbook = xlwt.Workbook(encoding='ufo-8')
    worksheet = workbook.add_sheet('货物名字')
    worksheet.write(0, 0, '名字')
    worksheet.write(0, 1, '坐标1')
    worksheet.write(0, 2, '坐标2')
    for i in range(len(list3)):
        list4_1[i] = str(list4_1[i])
    t = 0
    for i in range(len(list3)):
        list4[i] = str(list4[i])
        if '*' in list3[i]:
            list3[i] = '!' + list3[i]
            list4[i] = '!' + list4[i]
            flag = all('*' in a for a in list3) #判断list3是否所有值都包含‘*’
            if flag == True:   #如果是，直接加！
                list4_1[i] = '!' + list4_1[i]
            else:
                for j in range(t + 1, len(list3)):
                    if j == len(list3) - 1:
                        list4_1[j] = '!' + list4_1[j]
                        break
                    if '*' in list3[j]:
                        list4_1[j] = '!' + list4_1[j - 1]
                        break
        t = t + 1
    list4 = [value for value in list4 if "!" in value]  #保留带有‘！’的值
    list4 = [value.replace("!", "") for value in list4]  #将值中的‘！’删去
    list4_1 = [value for value in list4_1 if "!" in value]
    list4_1 = [value.replace("!", "") for value in list4_1]
    list3 = ''.join(list3)
    list3 = list3.split('!')
    list3.remove('')

    # 使用re模块将中文特殊字符替换为英文字符
    list3 = [re.sub(r'\(', '（', re.sub(r'\)', '）', word)) for word in list3]
    for j in range(len(list3)):
        worksheet.write(j + 1, 0, str(list3[j]))
        worksheet.write(j + 1, 1, str(list4[j]))
        worksheet.write(j + 1, 2, str(list4_1[j]))
    workbook.save(wjj + "/main/货物名字.xls")

    # 切割并识别图片
    j = 0
    i = 0
    img_path = wjj + '/Output/Output1.png'
    img = Image.open(img_path)
    width, height = img.size
    # 创表格，并识别图片
    workbook = xlwt.Workbook(encoding='ufo-8')
    worksheet = workbook.add_sheet('二次识别')
    worksheet.write(0, 0, 'y坐标')
    worksheet.write(0, 1, 'x坐标')
    worksheet.write(0, 2, '名字')
    a = 1
    while i < len(list3):
        if i ==len(list3)-1:#如果是最后一个（总数是奇数）
            wei1 = list4[i]
            wei2 = list4_1[i]
        else:
            similarity = difflib.SequenceMatcher(None, list3[i], list3[i + 1]).quick_ratio()
            if similarity>0.9:
                wei1 = list4[i]
                wei2 = list4_1[i + 1]
                i = i + 1
            else:
                wei1 = list4[i]
                wei2 = list4_1[i]
        j = j + 1
        path_2 = '/Output/' + str(j) + '.png'
        img = cv2.imread(wjj + '/Output/Output1.png')
        width_q, height_q = int(float(width) * 1.5), int((float(wei2) - float(wei1)) * 2)
        pts1 = np.float32([[0, wei1], [width, wei1], [0, wei2], [width, wei2]])
        pts2 = np.float32([[0, 0], [width_q, 0], [0, height_q], [width_q, height_q]])
        matrix = cv2.getPerspectiveTransform(pts1, pts2)
        imgoutput = cv2.warpPerspective(img, matrix, (width_q, height_q))
        cv2.imwrite(wjj + path_2, imgoutput)

        # 将图片顶上增加白色
        image = Image.open(wjj + path_2)
        width_1, height_1 = image.size
        # 创建一个新的空白图片，大小为原始图片加上扩展的像素
        new_image = Image.new("RGB", (width_1 + 200, height_1 + 200), (255, 255, 255))
        # 将原始图片添加到新图片上
        new_image.paste(image, (50, 50))
        new_image.save(wjj + path_2)

        save_path = wjj + '/Output/' + 'Outputs'+str(j)+ '.png'
        boxes, txts = jpg(wjj + path_2, save_path)
        q = 0
        while q < len(boxes):
            worksheet.write(a, 0, str(boxes[q][0][1]))
            worksheet.write(a, 1, str(boxes[q][0][0]))
            worksheet.write(a, 2, str(txts[q]))
            q = q + 1
            a = a + 1
        worksheet.write(a, 0, '#')
        worksheet.write(a, 1, '#')
        worksheet.write(a, 2, '#')
        a = a + 1
        i = i + 1
    workbook.save(wjj + '/main/二次识别.xls')

    # 加感叹号
    df = pd.read_excel(wjj + '/main/二次识别.xls')
    list1_x = df['x坐标'].tolist()
    list1_y = df['y坐标'].tolist()
    list2 = df['名字'].tolist()
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('xy+名字_2')
    worksheet.write(0, 0, 'y坐标')
    worksheet.write(0, 1, 'x坐标')
    worksheet.write(0, 2, '名字')
    a = float(list1_y[0])
    b = 1
    for i in range(len(list1_y)):
        if list2[i] == '#':
            if i != len(list1_y) - 1:
                a = float(list1_y[i + 1])
                worksheet.write(b, 0, '!')
                worksheet.write(b, 1, '!')
                worksheet.write(b, 2, '!')
                b = b + 1
            continue
        else:
            if abs(float(list1_y[i]) - a) > 40:
                worksheet.write(b, 0, '!')
                worksheet.write(b, 1, '!')
                worksheet.write(b, 2, '!')
                b = b + 1
                worksheet.write(b, 0, list1_y[i])
                worksheet.write(b, 1, list1_x[i])
                worksheet.write(b, 2, list2[i])
            else:
                worksheet.write(b, 0, list1_y[i])
                worksheet.write(b, 1, list1_x[i])
                worksheet.write(b, 2, list2[i])
            a = float(list1_y[i])
            b = b + 1
    worksheet.write(b, 0, '!')
    worksheet.write(b, 1, '!')
    worksheet.write(b, 2, '!')
    workbook.save(wjj + '/main/xy+名字_2.xls')

    df = pd.read_excel(wjj + '/main/xy+名字_2.xls')
    list1 = df['x坐标'].tolist()
    list2 = df['名字'].tolist()
    i = 0
    list3 = []  # X坐标
    list4 = []  # 名字
    start = 0
    while i < len(list1):
        if list1[i] == '!':
            if i - 1 == start:
                list3.extend([list1[start]])
                list4.extend([list2[start]])
            else:
                list3.extend([list1[start:i]])
                list4.extend([list2[start:i]])
            i = i + 1
            start = i
            continue
        i = i + 1
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('横板')
    A = ['货物名称', '规格型号', '单位', '数量', '单价', '金额+税率', '税额']
    for i in range(len(A)):  # 写表头
        worksheet.write(0, i, A[i])
    row = 1
    list5=[]
    hou=0
    # print(list3)
    # print(list4)
    for i in range(len(list3)): #按行循环
        if type(list3[i]) == list:#如果一行有多个值
            d = dict(zip(list3[i], list4[i]))  # 字典 键为x坐标，值为识别结果
            list3[i] = [float(x) for x in list3[i]]  # 因为识别后顺序出错，进行排序
            list3[i].sort()
            a = 0
            com = list3[0][0]  # 只比较第一行
            # print(list3)
            # print(d)
            # print('123')
            for j in range(len(list3[i])):#按列循环
                if i != 0:#只要不是第一行就执行
                    b = 0
                    for q in range(len(list3[0])):  # 判断这个的x坐标与第一张这一列的x坐标是否一样，不一样的写入下一列
                        c = abs(list3[i][j] - list3[0][b])  # x坐标的差
                        if c < 100:
                            if (b!=0):
                                if list5!=[]:
                                    if b>=list5[0]:
                                        worksheet.write(row, b+hou, d[str(list3[i][j])])
                                        break
                                worksheet.write(row, b , d[str(list3[i][j])])
                                break
                            else:
                                worksheet.write(row, b, d[str(list3[i][j])])
                            break
                        else:
                            b = b + 1
                else:  # 输入第一行
                    if (j > 0) and (j < (len(list3[0]) - 3)):
                        if j == 1:
                            if list3[0][j] - com > 850:
                                list5.append(a)
                                hou=hou+1
                                a = a + 1
                        else:
                            if list3[0][j] - com > 400:
                                list5.append(a)
                                a = a + 1
                                hou=hou+1
                        com = list3[0][j]
                    worksheet.write(row, a, d[str(list3[i][j])])
                    a = a + 1
        else:
            # 只有一个值
            worksheet.write(row, 0, list4[i])
        row = row + 1
    workbook.save(wjj + '/main/横版.xls')

    #将货物名称合并到一起
    df = pd.read_excel(wjj + '/main/横版.xls')
    list1 = df['货物名称'].tolist()
    list2 = []
    for column in df.columns:
        list2.append(df[column].tolist())
    location = []
    workbook = xlwt.Workbook(encoding='ufo-8')
    worksheet = workbook.add_sheet('货物明细')

    for i in range(len(list1)):
        if '*' in list1[i]:
            list1[i] = '!' + list1[i]
            location.append(i)
    list1 = ''.join(list1)
    list1 = list1.split('!')
    list1.remove('')
    # 使用re模块将中文特殊字符替换为英文字符
    list1 = [re.sub(r'\(', '（', re.sub(r'\)', '）', word)) for word in list1]
    list2.pop(0)
    for j in range(len(list1)):
        worksheet.write(j + 1, 0, list1[j])
        wei = location[j]
        for i in range(len(list2)):
            if str(list2[i][wei]) == 'nan':
                list2[i][wei] = ' '
            worksheet.write(j + 1, i+1 , list2[i][wei])
    A = ['货物名称', '规格型号', '单位', '数量', '单价', '金额+税率', '税额']
    for i in range(len(A)):
        worksheet.write(0, i, A[i])
    workbook.save(wjj + '/main/货物明细.xls')

    df = pd.read_excel(wjj + '/main/货物明细.xls')
    list1 = df['货物名称'].tolist()
    list2 = []
    list3 = []
    num=0
    for i in df.columns:
        if i == '金额+税率':
            list2.append(df[i].tolist())
            for j in range(len(list2[num])):
                if '免税' in str(list2[num][j]):
                    list2[num][j] = list2[num][j].replace('免税', '')
                    list3.append('免税')
                elif '%' in str(list2[num][j]):
                    list3.append(list2[num][j][-3:])
                    list2[num][j] = list2[num][j].replace(list3[-1], '')
            if list3!=[]:
                list2.extend([list3])
        else:
            list2.append(df[i].tolist())
        num = num + 1
    workbook = xlwt.Workbook(encoding='ufo-8')
    worksheet = workbook.add_sheet('货物明细')
    # 写内容
    for j in range(len(list1)):
        for i in range(len(list2)):
            worksheet.write(j + 1, i, list2[i][j])
    A = ['货物名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额',]
    for i in range(len(A)):
        worksheet.write(0, i, A[i])
    workbook.save(wjj + '/main/货物明细.xls')
    #核算价格，制作类表格
    df = pd.read_excel(wjj + '/main/货物明细.xls')
    list1 = df['货物名称'].tolist()
    list2 = []
    list3 = []
    num = 0
    for i in df.columns:
        if i == '金额+税率':
            list2.append(df[i].tolist())
            for j in range(len(list2[num])):
                if '免税' in str(list2[num][j]):
                    list2[num][j] = list2[num][j].replace('免税', '')
                    list3.append('免税')
                elif '%' in str(list2[num][j]):
                    list3.append(list2[num][j][-3:])
                    list2[num][j] = list2[num][j].replace(list3[-1], '')
            if list3 != []:
                list2.extend([list3])
        else:
            list2.append(df[i].tolist())
        num = num + 1
    workbook = xlwt.Workbook(encoding='ufo-8')
    worksheet = workbook.add_sheet('货物明细')
    # 写内容
    print(list1)
    for j in range(len(list1)):
        for i in range(len(list2)):
            worksheet.write(j + 1, i, list2[i][j])
    A = ['货物名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额', '价格','总价']
    for i in range(len(A)):
        worksheet.write(0, i, A[i])
    i=0
    al=0
    price=[]
    prelied=[]
    while i < len(list1):
        if list2[6][i] == "免税":
            if i == len(list1) - 1:#判断是否是最后一个
                b=list2[5][i]
                price.append(b)
            else:
                a = difflib.SequenceMatcher(None, list1[i], list1[i + 1]).quick_ratio()  # 判断相似度
                if a > 0.8:  # 判断货物是否相同
                        b = float(list2[5][i]) + float(list2[5][i + 1])
                        i = i + 1
                else:
                        b = float(list2[5][i])
        else:
            if i == len(list1) - 1:
                b = float(list2[5][i]) * (float(list2[6][i][0:-1]) / 100 +1)
            else:
                a = difflib.SequenceMatcher(None, list1[i], list1[i + 1]).quick_ratio() #判断相似度
                if a > 0.8:
                        b = float(list2[5][i])*(float(list2[6][i][0:-1])/100+1)+ float(list2[5][i + 1])*(float(list2[6][i+1][0:-1])/100+1)
                        i = i + 1
                else:
                        b = float(list2[5][i])*(float(list2[6][i][0:-1])/100+1)
        #提取名字的首部
        head_1=list1[i]
        lo = head_1.find('*')
        if lo == 0:
            head_1 = head_1.lstrip('*')
            lo = head_1.find('*')
        head_2 = head_1[0:lo]
        prelied.append(head_2)#**之间的内容
        price.append(b)#各个的价格
        worksheet.write(i+1, 8, b)
        al=al+b
        i = i + 1
    worksheet.write(1,9,al)
    workbook.save(wjj + '/result/货物明细.xls')
    print(prelied)
    workbook = xlwt.Workbook(encoding='ufo-8')
    worksheet = workbook.add_sheet('类')
    i=0    #将名字一样的价格加在一起
    while i < len(prelied) - 1: #内容
        j = i+1
        while j<=len(prelied) - 1:#价格
            a = difflib.SequenceMatcher(None, prelied[i], prelied[j]).quick_ratio()  # 判断相似度
            if a>0.65:
                prelied.pop(j)
                price[i]=price[i]+price[j]
                price.pop(j)
                j=j-1
            j=j+1
        i=i+1

    worksheet.write(0,0,'类')
    worksheet.write(0, 1, '价格')
    for i in range(len(prelied)):
        worksheet.write(1,0,prelied[i])
        worksheet.write(1, 1, price[i])
    workbook.save(wjj + '/result/分类.xls')

