import fitz
import pypinyin
import shutil

from pypinyin import pinyin
from pathlib import Path
from sublist import *
from mysq import *


def Main(path,T):
    if os.path.isfile(path):#如果是文件
        main(path,T)
        print('123')
    else:#是文件夹
        files = os.listdir(path) #输出文件夹内文件的文件名
        for file in files: #遍历文件
            if os.path.splitext(file)[1]== ".pdf":   #判断文件是否是pdf文件
                path_1=path+'/'+file
                main(path_1,T)



def main(pdfpath, T):
    # os.chdir(pdfpath)
    path_1 = r"D:\python_Project\invoice_value\result"
    judge(path_1)
    Init(path_1)
    shutil.copy(pdfpath, path_1)
    file = os.listdir(path_1)  # 遍历所有文件
    for i in file:
        oldname = os.path.join(path_1, i)
        if os.path.isfile(oldname):  # 判断是否是文件
            filename1 = os.path.splitext(i)[0]  # 书籍发票-1
            filename2 = pinyin(filename1, style=pypinyin.FIRST_LETTER)  # [['s'], ['j'], ['f'], ['p'], ['-1']]
            l1 = []
            for j in filename2:
                l1.extend(j)
            l2 = ''.join(l1)  # sjfp-1
            name = l2 + os.path.splitext(i)[1]  # sjfp-1.pdf
            newname = os.path.join(path_1, name)
            os.rename(oldname, newname)
    wjj = os.path.splitext(newname)[0]  # D:\2\xgk\result\sjfp-1
    global path
    path=wjj

    judge(wjj)

    os.chdir(wjj)  # 跳转到路径
    A = ['main', 'Output', 'result', 'png']
    for i in A:
        judge(i)
    # pdf图片转为png图片
    pdfDoc = fitz.open(pdfpath)  # 获取document对象
    totalPage = pdfDoc.page_count  # 发票总页数
    for pg in range(totalPage):
        page = pdfDoc[pg]
        rotate = int(0)
        zoom_x = 2
        zoom_y = 2
        mat = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)  # 创建变换矩阵mat
        pix = page.get_pixmap(matrix=mat)  # 获取图像
        pix.save(wjj + f'/png/images_{pg + 1}.png')

    img_path = wjj + '/png/images_1.png'
    save_path = wjj + '/result/result.jpg'
    boxes, txts = jpg(img_path, save_path)

    # # 密码区
    # img_path = wjj + '/png/images_1.png'
    # img = cv2.imread(img_path)  # 读取图片
    # width, height = 465, 125  # 图片宽高
    # pts1 = np.float32([[730, 170], [1200, 170], [730, 300], [1200, 300]])  # 大图上的坐标
    # pts2 = np.float32([[0, 0], [width, 0], [0, height], [width, height]])  # 切割后图片的坐标
    # matrix = cv2.getPerspectiveTransform(pts1, pts2)  # 输出透视矩阵
    # imgoutput = cv2.warpPerspective(img, matrix, (width, height))  # 输出进行透视变换后的图片矩阵
    # save_path = wjj + '/output/mimaqu.png'  # 输出路径
    # cv2.imwrite(save_path, imgoutput)  # 保存png图片
    # sc = wjj + '/Output/密码区.jpg'
    # boxes0, txts0 = jpg(save_path, sc)
    # txts0 = ['密码区：'] + txts0
    # txts0 = ''.join(txts0)

    # 创建表格
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('识别结果')
    worksheet.write_merge(0, 0, 0, 3, '坐标')
    worksheet.write(0, 4, '文字')
    # worksheet.write(1, 4, txts0)
    for e in range(len(txts)):
        worksheet.write(e + 2, 4, str(txts[e]))
    for i in range(len(boxes)):
        worksheet.write(i + 2, 0, str(boxes[i][0]))
        worksheet.write(i + 2, 1, str(boxes[i][1]))
        worksheet.write(i + 2, 2, str(boxes[i][2]))
        worksheet.write(i + 2, 3, str(boxes[i][3]))
    for i in range(5):
        worksheet.col(i).width = 256 * 20
    workbook.save(wjj + '/main/识别结果.xls')

    T1 = [value for value in T.values()]  # 输出T的值的列表
    l3 = ['称：', '纳税人', '开户行', '地址']
    Q = l3 + T1
    workbook = xlwt.Workbook(encoding='utf-8')
    worksheet = workbook.add_sheet('提取结果')
    data_excel = xlrd.open_workbook(wjj + '/main/识别结果.xls')
    sheet_object = data_excel.sheet_by_index(0)  # 获取第一张工作表的内容
    n_rows = sheet_object.nrows  # 获取行数
    y = 0
    for i in range(n_rows):
        content = sheet_object.cell_value(i, 4)
        for j in range(len(Q)):
            if content.find(Q[j]) != -1:  # 找到
                a = content.split('：')
                if a[1] == '':
                    a[1] = 'none'
                worksheet.write(1, y, a[1])
                y = y + 1

    A = {1: '密码区', 2: '发票代码', 3: '发票号码', 4: '开票日期', 5: '机器编号', 6: '校验码', 7: '购买方名称',
         8: '购买方纳税人识别号', 9: '购买方地址、电话',
         10: '购买方开户行及账号', 11: '销售方名称', 12: '销售方纳税人识别号', 13: '销售方地址、电话',
         14: '销售方开户行及账号', 15: '收款人', 16: '复核', 17: '开票人'}

    key = list(T.keys())
    key.sort()
    for i in key:
        if T[i] == '':
            del T[i]
    key = list(T.keys())
    key.sort()  # 排序
    l2 = []
    for i in key:
        l2.append(A[i])
    for i in range(len(l2)):
        worksheet.write(0, i, l2[i])
    workbook.save(wjj + '/result/保留结果.xls')


    if totalPage==2:
        img_path = wjj+'\png\images_2.png'
        save_path = wjj + '/result/result_2.jpg'
        img = Image.open(img_path)
        # 获取图像的宽度和高度
        width, height = img.size
        shibie(totalPage,width,height,wjj,img_path,save_path)


    else:
        img_path = wjj+'\png\images_1.png'
        save_path = wjj + '/result/result_1.jpg'
        img = Image.open(img_path)
        width, height = img.size
        shibie(totalPage, width, height, wjj, img_path, save_path)

      #给值旁边加一列坐标，比对x坐标大小，替换值

    # 合并结果
    file_path = wjj + '/总结果.xls'
    workbook = pd.ExcelWriter(file_path)
    folder_path = Path(wjj + '/result/')  # 打开路径下所有文件
    file_list = folder_path.glob('*.xls*')
    for file in file_list:
        stem_name = file.stem
        data = pd.read_excel(file, sheet_name=0)
        data.to_excel(workbook, sheet_name=stem_name, index=False)
    workbook.save()
    print("文件创建完成")
    #传输到数据库
    path_file=(wjj+'/result')
    file = os.listdir(path_file)
    for i in file:
        if i == '保留结果.xls':
            path_1=wjj+'/result/保留结果.xls'
            my_1(path_1)
            print('数据输入成功')
        if i == '分类.xls':
            path_2=wjj+'/result/分类.xls'
            my_2(path_2)
            print('数据输入成功')
        if i == '货物明细.xls':
            path_3=wjj+'/result/货物明细.xls'
            my_3(path_3)
            print('数据输入成功')



def file():
    import os
    global path
    start_directory = path + '/' + 'png'
    os.startfile(start_directory)

def file_1(path):
    import os

    start_directory = path
    os.startfile(start_directory)

def dir():
    import os
    global path

    start_directorys = path
    os.startfile(start_directorys)




if __name__ == "__main__":
    pdfpath = r"D:\python_Project\invoice_value\sjfp\dzfp.pdf" # 目标发票位置
    T = {1: '密码区', 2: '发票代码', 3: '发票号码', 4: '开票日期', 5: '机器编号', 6: '校验码', 7: '购买方名称',
         8: '购买方纳税人识别号', 9: '购买方地址、电话', 10: '购买方开户行及账号', 11: '销售方名称',
         12: '销售方纳税人识别号',
         13: '销售方地址、电话', 14: '销售方开户行及账号'}
    Main(pdfpath, T)

