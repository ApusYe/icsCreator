# coding: utf-8
#!/usr/bin/python

import sys
import json
import xlrd

__author__ = 'ChanJH'
# modified for wider application by ytf


# 指定信息在 xls 表格内的列数
_colOfClassName = 0
_colOfStartWeek = 1
_colOfEndWeek = 2
_colOfWeekday = 3
_colOfClassStartTime = 4
_colOfClassEndTime = 5
_colOfClassroom = 6

def main():
	# 读取 excel 文件
	data = xlrd.open_workbook('classInfo.xlsx')
	table = data.sheets()[0]
	# print table.cell(1,0).value
	# 基础信息
	numOfRow = table.nrows  #获取行数,即课程数
	numOfCol = table.ncols  #获取列数,即信息量
	headStr = '{\n"classInfo":[\n'
	tailStr = ']\n}'
	classInfoStr = ''
	classInfoArray = []
	# 信息列表
	# lengthOfList = numOfRow-1
	classNameList = []
	startWeekList = []
	endWeekList = []
	weekdayList = []
	classStartTimeList = []
	classEndTimeList = []
	classroomList = []

	# 确定配置内容
	info = "\nPart 1 : Excel解析。\n请检查 Excel 列信息配置。\n\n"
	info += "ClassName: " + str(_colOfClassName) + "列\n"
	info += "StartWeek: " + str(_colOfStartWeek) + "列\n"
	info += "EndWeek: " + str(_colOfEndWeek) + "列\n"
	info += "Weekday: " + str(_colOfWeekday) + "列\n"
	info += "ClassStartTime: " + str(_colOfClassStartTime) + "列\n"
	info += "ClassEndTime: " + str(_colOfClassEndTime) + "列\n"
	info += "Classroom: " + str(_colOfClassroom) + "列\n"
	print (info)
	# info += "输入 0 继续，输入 1 退出："
	option = raw_input("输入 0 继续，输入其他内容退出：")
	if option == "1":
		sys.exit()
	

	# 开始操作
	# 将信息加载到列表
	i = 1
	while i < numOfRow :
		index = i-1

		classNameList.append(((table.cell(i, _colOfClassName).value)))
		startWeekList.append(str(int((table.cell(i, _colOfStartWeek).value))))
		endWeekList.append(str(int((table.cell(i, _colOfEndWeek).value))))
		weekdayList.append(str(int((table.cell(i, _colOfWeekday).value))))
		classStartTimeList.append(str(int((table.cell(i, _colOfClassStartTime).value))))
		classEndTimeList.append(str(int((table.cell(i, _colOfClassEndTime).value))))
		classroomList.append(str(((table.cell(i, _colOfClassroom).value))))
		
		i += 1
	i = 0
	itemHeadStr = '{\n'
	itemTailStr = '\n}'

	classInfoStr += headStr
	for className in classNameList:
		itemClassInfoStr = ""
		itemClassInfoStr  = itemHeadStr + '"className":"' + className + '",'
		itemClassInfoStr += '"week":{\n"startWeek":' + startWeekList[i] + ',\n'
		itemClassInfoStr += '"endWeek":' + endWeekList[i] + '\n},\n'
		itemClassInfoStr += '"weekday":' + weekdayList[i] + ',\n'
		itemClassInfoStr += '"classStartTime":' + classStartTimeList[i] + ',\n'
		itemClassInfoStr += '"classEndTime":' + classEndTimeList[i] + ',\n'
		itemClassInfoStr += '"classroom":"' + classroomList[i] + '"\n'
		itemClassInfoStr += itemTailStr
		classInfoStr += itemClassInfoStr
		if i!=len(classNameList)-1 :
			classInfoStr += ","
		i += 1
	classInfoStr += tailStr
	# print classInfoStr
	with open('conf_classInfo.json','w') as f:

		f.write(classInfoStr)
		f.close()
	print("\nALL DONE !")

reload(sys);
sys.setdefaultencoding('utf-8');
main()
