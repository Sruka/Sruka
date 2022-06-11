import xlrd

xlsx = xlrd.open_workbook('excel.xlsx')

# 通过sheet名查找：xlsx.sheet_by_name("sheet1")
# 通过索引查找：xlsx.sheet_by_index(3)
table = xlsx.sheet_by_index(0)
nrows = table.nrows
print("共有", nrows, "列")
# 校区
print("请输入校区名（开城/壹品）")
a = input()

# 年级（填阿拉伯数字）
print("请输入学生年级（阿拉伯数字）")
b = input()
# 科目
print("请输入学生小班科目1（没有请直接跳过）")
c = input() + "班"
print("请输入学生小班科目2（没有请直接跳过）")
d = input() + "班"
print("请输入学生小班科目3（没有请直接跳过）")
e = input() + "班"
print("请输入学生小班科目4（没有请直接跳过）")
f = input() + "班"
print("请输入学生姓名（无1V1课程可跳过）")
name = input()  # 一对一学生姓名

if c == "班":
    c = "无课程"
if d == "班":
    d = "无课程"
if e == "班":
    e = "无课程"
if f == "班":
    f = "无课程"
if name == "":
    name = "name"
# 以下为周六课程
i = 1
for i in range(nrows):
    value = table.cell_value(i, 1)  # 从第1行第2列开始到第n行第2列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周六":
            print("周六", table.cell_value(0, 1), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 2)  # 从第1行第3列开始到第n行第3列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周六":
            print("周六", table.cell_value(0, 2), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 3)  # 从第1行第4列开始到第n行第4列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周六":
            print("周六", table.cell_value(0, 3), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 4)  # 从第1行第5列开始到第n行第5列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周六":
            print("周六", table.cell_value(0, 4), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 5)  # 从第1行第6列开始到第n行第6列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周六":
            print("周六", table.cell_value(0, 5), value)  # 此处输出 周X 时间 什么课
            i = i + 1

# 以下为周日课程
i = 1
for i in range(nrows):
    value = table.cell_value(i, 1)  # 从第1行第2列开始到第n行第2列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周日":
            print("周日", table.cell_value(0, 1), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 2)  # 从第1行第3列开始到第n行第3列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周日":
            print("周日", table.cell_value(0, 2), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 3)  # 从第1行第4列开始到第n行第4列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周日":
            print("周日", table.cell_value(0, 3), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 4)  # 从第1行第5列开始到第n行第5列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周日":
            print("周日", table.cell_value(0, 4), value)  # 此处输出 周X 时间 什么课
            i = i + 1
i = 1
for i in range(nrows):
    value = table.cell_value(i, 5)  # 从第1行第6列开始到第n行第6列
    if (a in value and b in value and (c in value or d in value or e in value or f in value)) or (name in value):
        if table.cell_value(i, 0) == "周日":
            print("周日", table.cell_value(0, 5), value)  # 此处输出 周X 时间 什么课
            i = i + 1

print("按下回车键退出")
z = input()
