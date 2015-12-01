 #coding:utf-8

import sys
import time
import xlrd

reload(sys)
sys.setdefaultencoding('utf-8')


def getArrayStr(strValue, isNumArray):
	strArray = "{"
	if(strValue != ""):
		arrSplit = strValue.split(",")
		for i in range(len(arrSplit)):
			if isNumArray == True:
				strArray = strArray + str(arrSplit[i]) + ", "
			else:
				strArray = strArray + "\"" + str(arrSplit[i]) + "\", "

	strArray = strArray + "}"

	return strArray




def writeComments():
	luaFile.write("--[[\n" + 
			  "\tDate:\t" + time.strftime('%Y-%m-%d %H:%M:%S',time.localtime(time.time())) + "\n" + 
			  "\tAuthor:\thuqing@kingsoft.com" + "\n" + 
			  "\tExcel:\t" + strXlsName + "\n" + 
			  "\tSheet:\t" + strSheetName + "\n" + 
			  "]]\n\n")




def writeBegin():
	luaFile.write("module(\"" + strSheetName + "\")\n\n")
	luaFile.write(strSheetName + " = \n{\n")




def writeLine(nIndex, arrKey, arrType, arrValue):
	strLine = ""
	strKey = arrKey[0]
	strType = arrType[0]

	strKeyQuot = ""

	if(strType == "Number"):
		strKeyQuot = ""
	else:
		strKeyQuot = "\""

	for i in range(len(arrKey)):
		strKey = arrKey[i]
		strType = arrType[i]

		if(i == 0):
			_strKey = str(arrValue[i])
			if strKeyQuot == "":
				_strKey = float(_strKey)
				_strKey = int(_strKey)

			strLine = "[" + strKeyQuot + str(_strKey) + strKeyQuot + "] = {"
			continue

		if(strType == "String"):
			strLine = strLine + "[\"" + str(arrKey[i]) + "\"] = \"" + str(arrValue[i]) + "\", "
		elif(strType == "Number"):
			strLine = strLine + "[\"" + str(arrKey[i]) + "\"] = " + str(arrValue[i]) + ", "
		elif(strType == "NumberArray"):
			strArray = getArrayStr(arrValue[i], True)
			strLine = strLine + "[\"" + str(arrKey[i]) + "\"] =" + strArray + ", "
		elif(strType == "StringArray"):
			strArray = getArrayStr(arrValue[i], False)
			strLine = strLine + "[\"" + str(arrKey[i]) + "\"] =" + strArray + ", "
		else:
			strLine = strLine + "[\"" + str(arrKey[i]) + "\"] = \"" + str(arrValue[i]) + "\", "

	luaFile.write("\t" + strLine + "}, \n")
		




def writeEnd():
	luaFile.write("}")





strXlsName = sys.argv[1]
strSheetName = sys.argv[2]

xls = xlrd.open_workbook(strXlsName)
sheet = xls.sheet_by_name(strSheetName)

luaFile = open(strSheetName + ".lua", 'w')

print "convert " + strXlsName + "[" + strSheetName + "]" + " to " + strSheetName + ".lua"

writeComments()
writeBegin()


arrKey = []
arrType = []
arrDes = []

strLine = ""

for row in range(sheet.nrows):

	arrValue = []

	for col in range(sheet.ncols):
		# 第一行是Key
		if(row == 0):
			arrKey.append(str(sheet.cell(row, col).value))
		# 第二行是类型
		elif(row == 1):
			arrType.append(str(sheet.cell(row, col).value))
		# 第三行是描述
		elif(row == 2):
			arrDes.append(str(sheet.cell(row, col).value))
		# 从第四行开始是值
		else:
			arrValue.append(str(sheet.cell(row, col).value))

	if(row > 2):
		writeLine(row - 2, arrKey, arrType, arrValue)


writeEnd()

luaFile.close( )
