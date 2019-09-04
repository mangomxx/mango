# -*- coding: utf-8 -*-
import matplotlib.pyplot as plt
import cx_Oracle
import pandas as pd
from pylab import *  # 支持中文
import xlsxwriter
global file_path

file_path = r'C:\Users\18810\Desktop\Python\\'


"""
产生报告思路：
1.加载数据集  loadDataSet(col, sql)
2.写出本地图片 result_pic(result)
3.将图片写入本地已存在的excel表格中，按照图片名称另存为新的表格
4.将每个人的分析数据插入到图片页中
5.将大区 区域 top的数据分别addsheet到上述表格中
6.将明细合并

"""

# 定义取出数据函数
def loadDataSet(col, sql):
    """
    加载数据集
    :param fileName: 文件名称
    :return:list类型
    """
    # dataMat = []
    conn = cx_Oracle.connect("BUS_DA/BUS_DA_20181205@172.16.12.121/biwork")
    # 用自己的实际数据库用户名、密码、主机ip地址 替换即可
    curs = conn.cursor()
    curs.execute(sql)
    result = curs.fetchall()
    dataMat = pd.DataFrame(result, columns=col)
    return dataMat


def result_pic(result):
    """
    雷达图的绘制
    :param result: 分类数据
    :return: 雷达图
    """
    # 解析出类别标签和种类
    data_labels = ['新增有效房源量分数', '新增有效客户量分数', '被带看次数分数', '带看组数分数', '活跃房源占比分数', '二看率分数', '总业绩分数']
    data_name = list(result.iloc[:, 1])  # 筛选出每个表的名字,即每个人的姓名
    data_id = list(result.iloc[:, 0])  # 筛选出每个表的id 即每个人的id
    data_use = np.array(result.iloc[:, 2:])  # 筛选出表的数据
    # print(data_use)
    # 数据个数
    data_len = len(data_labels)
    for i in range(0, len(data_id)):
        data = list(map(int, data_use[i]))  # .array(data_use[i])
        angles = np.linspace(0, 2 * np.pi, data_len, endpoint=False)
        data = np.concatenate((data, [data[0]]))  # 闭合
        angles = np.concatenate((angles, [angles[0]]))  # 闭合

        fig = plt.figure()
        ax = fig.add_subplot(111, polar=True)  # polar参数！！
        ax.plot(angles, data, 'bo-', linewidth=2)  # 画线
        ax.fill(angles, data, facecolor='r', alpha=0.25)  # 填充
        ax.set_thetagrids(angles * 180 / np.pi, data_labels, fontproperties="SimHei")
        ax.set_title('个人业务能力雷达图', va='bottom', fontproperties="SimHei")
        # ax.set_title(str(data_name[i]), va='bottom', fontproperties="SimHei")
        ax.set_rlim(10, 100)  # 设置雷达图的范围
        ax.grid(True)
        plt.savefig(file_path + str(data_name[i]) + ".png", dpi=120)
        # plt.show()

# 定义成绩是否优秀
def levSet(data):
    if data<50: lev='较差'
    elif data>=50 and data<80: lev='中等'
    else:lev='优秀'
    return lev



def Pic_to_Excel(result):
    data_name = list(result.iloc[:, 1])  # 筛选出每个表的名字,即每个人的姓名
    data_id = list(result.iloc[:, 0])  # 筛选出每个表的id 即每个人的id
    data_use = np.array(result.iloc[:, 2:])  # 筛选出表的数据

    for i in range(0, len(data_id)):

        data = list(map(int, data_use[i]))
        # print(data[0:1])
        workbook = xlsxwriter.Workbook(file_path + str(data_name[i]) +'.xlsx')
        worksheet = workbook.add_worksheet('个人业务能力分析')
        worksheet.insert_image('A2', file_path + str(data_name[i]) + ".png") # Insert an image.
        worksheet.set_column('N:N', 14)  # 调节列宽
        worksheet.write('M2', '业务流程：新增——转化——带看——成交')  # Insert an text.
        worksheet.write('M4', '新增指标：新增有效房源量，新增有效客户量')  # Insert an text.
        worksheet.write('M6', '转化指标：活跃房源占比，二看率')  # Insert an text.
        worksheet.write('M8', '带看指标：带看组数，被带看次数')  # Insert an text.
        worksheet.write('M10', '成交指标：总业绩')  # Insert an text.
        worksheet.write('M12', '雷达图解析：各指标最高100分，最低20分，取值周期：发报告日起往前推30天')  # Insert an text.
        worksheet.write('M14', '业务能力分析：')  # Insert an text.
        worksheet.write('N15', '总分：')  # Insert an text.
        worksheet.write('N16', '新增指标分数：')  # Insert an text.
        worksheet.write('N17', '带看指标分数：')  # Insert an text.
        worksheet.write('N18', '转化指标分数：')  # Insert an text.
        worksheet.write('N19', '总业绩分数')  # Insert an text.
        worksheet.write('O15', '%.2f' %np.mean(data))  # 计算总分数
        worksheet.write('O16', str('%.0f' %np.mean(data[0:1]))+'分')  # 计算分块分数
        worksheet.write('O17', str('%.0f' %np.mean(data[2:3]))+'分')  # 计算分块分数
        worksheet.write('O18', str('%.0f' %np.mean(data[4:5]))+'分')  # 计算分块分数
        worksheet.write('O19', str('%.0f' %data[6])+'分')  # 计算分块分数
        worksheet.write('P15', levSet(np.mean(data)))  # 判断等级
        worksheet.write('P16', levSet(np.mean(data[0:1])) )  # 判断等级
        worksheet.write('P17', levSet(np.mean(data[2:3])) )  # 判断等级
        worksheet.write('P18', levSet(np.mean(data[4:5])) )  # 判断等级
        worksheet.write('P19', levSet(data[6]) )  # 判断等级
        workbook.close()

        # 50-80中等
        # >80优秀
        # <50较差



if __name__ == '__main__':
    col1 = ['id', 'name','新增有效房源量分数', '新增有效客户量分数', '被带看次数分数',
            '带看组数分数', '活跃房源占比分数', '二看率分数', '总业绩分数']
    sql1 = '''SELECT id,大区||'-'||区域||'-'||分店||'-'||经纪人,
                新增有效房源量分数,新增有效客户量分数,被带看次数分数,带看组数分数,
                活跃房源占比分数,二看率分数,总业绩分数
                FROM BUS_DA.TEST_LEIDA
                where 大区 in ('西大区')'''  # sql语句

    result1 = loadDataSet(col1,sql1)
    Pic_to_Excel(result1)
    print('finish')
    result_pic(result1)






