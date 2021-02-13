import xlrd
import pymysql
#打开数据所在的工作簿，以及选择存有数据的工作表
book = xlrd.open_workbook(r"C:\Users\Kevin\i\3.xlsx")
sheet = book.sheet_by_name("Sheet1")
#建立一个MySQL连接
conn = pymysql.connect(
        "127.0.0.1", "root", "123456", "testdb", charset='utf8'
        )
# 获得游标
cur = conn.cursor()
# 创建插入SQL语句
query = 'insert into student_tbl(name,sex,minzu,danwei_zhiwu,phone_number,home_number) values (%s, %s, %s, %s, %s, %s)'
# 创建一个for循环迭代读取xls文件每行数据的, 从第二行开始是要跳过标题行
for r in range(1, sheet.nrows):
      name      = sheet.cell(r,1).value
      sex       = sheet.cell(r,2).value
      minzu          = sheet.cell(r,3).value
      danwei_zhiwu     = sheet.cell(r,4).value
      phone_number       = sheet.cell(r,5).value
      home_number = sheet.cell(r,6).value
      values = (name,sex,minzu,danwei_zhiwu,phone_number,home_number)
      # 执行sql语句
      cur.execute(query, values)
cur.close()
conn.commit()
conn.close()
columns = str(sheet.ncols)
rows = str(sheet.nrows)
print ("导入 " +columns + " 列 " + rows + " 行数据到MySQL数据库!")