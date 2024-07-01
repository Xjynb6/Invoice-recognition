import pymysql
import pandas as pd
import os

def my_1(path):
    # 实验室端口
    conn=pymysql.connect(
        host='124.222.118.135',   #主机名
        port=3306,          #端口号
        user='root',        #用户名
        password='team2111.',#密码
        autocommit=True        # 自动提交更改
    )

    # conn=pymysql.connect(
    #     host='localhost',   #主机名
    #     port=3306,          #端口号
    #     user='root',        #用户名
    #     password='xjy123456',#密码
    #     autocommit=True        # 自动提交更改
    # )
    # 创建游标对象
    cursor = conn.cursor()
    #选择数据库
    conn.select_db('xjy')
    dirname = os.path.dirname(path)
    folder_name = os.path.basename(dirname)

    # 创建表
    sql = f"create table if not exists baoliu ( password varchar(200) comment '密码区',\
    code varchar(15)comment '发票代码',\
    number int comment '发票号码',\
    date varchar(15) comment '开票日期',\
    machine_code varchar(40) comment '机器编号',\
    ceck_code varchar(40) comment '校验码',\
    buyer_name varchar(100) comment '购买方名称',\
    buyer_number varchar(40) comment '购买方纳税人识别号',\
    buyer_address_phone varchar(100) comment '购买方地址、电话',\
    buyer_bank_number varchar(100) comment '购买方开户行及账号',\
    seller_name varchar(100) comment '销售方名称',\
    seller_number varchar(40) comment '销售方纳税人识别号',\
    seller_address_phone varchar(100) comment '销售方地址、电话',\
    seller_bank_number varchar(100) comment '销售方开户行及账号',\
    payee varchar(10) comment '收款人',\
    re_check varchar(10) comment '复核',\
    drawer varchar(10) comment '开票人')\
    comment '保留结果' default charset=utf8 "
    cursor.execute(sql)
    sql = f"SELECT  COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'xjy' AND TABLE_NAME = 'baoliu'"
    cursor.execute(sql)
    result1 = cursor.fetchall()#注释名
    sql=f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'xjy' AND TABLE_NAME = 'baoliu'"
    cursor.execute(sql)
    result2 = cursor.fetchall()#字段名
    list3=[]
    list4=[]
    # print(path)
    df = pd.read_excel(path)
    list1=df.columns#表头
    list2 = df.iloc[0]#第一行
    for i in range(len(result1)):
        for j in range(len(list1)):
            if result1[i][0] == list1[j]:  #如果表头和注释一样，往数据库添加数据
                list3.append(result2[i][0])
                list4.append(list2[j])
                break
    a=''
    for i in range(len(list3)):
        if i !=len(list3)-1:
            a=a+list3[i]+','
        else:
            a=a+list3[i]
    sql="insert into baoliu ({}) values {}".format(a,tuple(list4))

    cursor.execute(sql)
    cursor.close()
    conn.close()
def my_2(path):
    # 实验室端口
    conn = pymysql.connect(
        host='124.222.118.135',  # 主机名
        port=3306,  # 端口号
        user='root',  # 用户名
        password='team2111.',  # 密码
        autocommit=True  # 自动提交更改
    )
    # conn = pymysql.connect(
    #     host='localhost',  # 主机名
    #     port=3306,  # 端口号
    #     user='root',  # 用户名
    #     password='xjy123456',  # 密码
    #     autocommit=True  # 自动提交更改
    # )
    # 创建游标对象
    cursor = conn.cursor()
    # 选择数据库
    conn.select_db('xjy')
    dirname = os.path.dirname(path)
    folder_name = os.path.basename(dirname)
    sql = f"create table if not exists fenlei ( class varchar(100) comment '类',\
    price varchar(40) comment '价格')comment '分类' default charset=utf8 "
    cursor.execute(sql)
    sql = f"SELECT  COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'xjy' AND TABLE_NAME = 'fenlei'"
    cursor.execute(sql)
    result1 = cursor.fetchall()  # 注释名
    sql = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'xjy' AND TABLE_NAME = 'fenlei'"
    cursor.execute(sql)
    result2 = cursor.fetchall()  # 字段名
    list3 = []
    list4 = []
    df = pd.read_excel(path)
    list1 = df.columns  # 表头
    list2 = df.iloc[0]  # 第一行
    for i in range(len(result1)):
        for j in range(len(list1)):
            if result1[i][0] == list1[j]:  # 如果表头和注释一样，往数据库添加数据
                list3.append(result2[i][0])
                list4.append(list2[j])
                break
    a = ''
    for i in range(len(list3)):
        if i != len(list3) - 1:
            a = a + list3[i] + ','
        else:
            a = a + list3[i]
    sql = "insert into fenlei ({}) values {}".format(a, tuple(list4))
    cursor.execute(sql)
def my_3(path):
    # 实验室端口
    conn = pymysql.connect(
        host='124.222.118.135',  # 主机名
        port=3306,  # 端口号
        user='root',  # 用户名
        password='team2111.',  # 密码
        autocommit=True  # 自动提交更改
    )
    # conn = pymysql.connect(
    #     host='localhost',  # 主机名
    #     port=3306,  # 端口号
    #     user='root',  # 用户名
    #     password='xjy123456',  # 密码
    #     autocommit=True  # 自动提交更改
    # )
    # 创建游标对象
    cursor = conn.cursor()
    # 选择数据库
    conn.select_db('xjy')
    dirname = os.path.dirname(path)
    folder_name = os.path.basename(dirname)
    sql = f"create table if not exists huowu (name varchar(200) comment '货物名称',\
    specifications varchar(100) comment '规格型号',\
    unit varchar(20) comment '单位',\
    quantity varchar(20) comment '数量',\
    unit_price varchar(10) comment '单价',\
    money varchar(20) comment '金额',\
    tax_rate varchar(10) comment '税率',\
    tax varchar(10) comment '税额',\
    price varchar(20) comment '价格',\
    total_price varchar(20) comment '总价')\
    comment '货物明细' default charset=utf8"
    cursor.execute(sql)
    sql = f"SELECT  COLUMN_COMMENT FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'xjy' AND TABLE_NAME = 'huowu'"
    cursor.execute(sql)
    result1 = cursor.fetchall()  # 注释名
    sql = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = 'xjy' AND TABLE_NAME = 'huowu'"
    cursor.execute(sql)
    result2 = cursor.fetchall()  # 字段名

    df = pd.read_excel(path)
    list1 = df.columns  # 表头
    # list2 = df.iloc[0]  # 第一行
    for index, list2 in df.iterrows():
        list3 = []
        list4 = []
        for i in range(len(list2)):
            if str(list2[i])=='nan':
                list2[i]=' '
        for i in range(len(result1)):
            for j in range(len(list1)):
                if result1[i][0] == list1[j]:  # 如果表头和注释一样，往数据库添加数据
                    list3.append(result2[i][0])
                    list4.append(list2[j])
                    break
        a = ''
        for i in range(len(list3)):
            if i != len(list3) - 1:
                a = a + list3[i] + ','
            else:
                a = a + list3[i]
        sql = "insert into huowu ({}) values {}".format(a, tuple(list4))
        cursor.execute(sql)


# if __name__=='__main__':
#     path=r"D:\2\xgk\result\sjfp-1/总结果.xls"
    # my(path)
    # sheet_names = pd.read_excel(path, sheet_name=None).keys()
    # print(sheet_names)
