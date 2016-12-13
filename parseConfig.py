# -*- coding: utf-8 -*-

import xdrlib, sys
import xlrd
import codecs

global sheet
global MergePoints
global text

class node:
	def __init__(self, des, x, y):
		self._des = des
		self._coordinate = [x, y]
		self._children = []
	def getDes(self):
		return self._des
	def getCoordinate(self):
		return self._coordinate
	def getChildren(self):
		return self._children
	def addChild(self, node):
		self._children.append(node)

class tree:
	def __init__(self):
		self._head = node("head", -1, -1)
	def addChild(self, node):
		self._head.addChild(node)
	def getHead(self):
		return self._head

def generateTree(rowMax, colMax):
	structTree = tree()
	row = 0
	for col in range(0, colMax):
		point = [row, col]
		endCol = col
		treeAddNodes(structTree, row, col, endCol)
	return structTree

def treeAddNodes(father, row, col, endCol):
	global sheet
	global MergePoints
	point = [row, col]
	if not isPointInMergePoints(point):
		if col < endCol:
			# 遍历叶子的邻接点
			addNodes(father, row, col)
			treeAddNodes(father, row, col + 1, endCol)
		else :
			addNodes(father, row, col)
			return
	else:
		if not isHeadOfMergePoints(point):
			if col < endCol:
				# 合并单元格的非头部结点只遍历不解析
				treeAddNodes(father, row, col + 1, endCol)
		else:
			# update endCol
			newEndCol = getMergePointsEndCol(point)
			newNode = addNodes(father, row, col)
			# 合并单元格的头部结点子节点遍历
			treeAddNodes(newNode, row + 1, col, newEndCol)
			# 合并单元格的头部结点的邻接点遍历
			treeAddNodes(father, row, col + 1, endCol)

def addNodes(father, row, col):
	global sheet
	des = sheet.cell_value(row, col)
	newNode = node(des, row, col)
	father.addChild(newNode)
	return newNode

def generateMergePoints(merges):
	points = []
	for (x, xMax, y, yMax) in merges:
		for i in range(x,xMax):
			for j in range(y,yMax):
				points.append([i, j])
	return points

def isPointInMergePoints(point):
	global MergePoints
	for p in MergePoints:
		if point[0] == p[0] and point[1] == p[1]:
			return True
	return False

def isHeadOfMergePoints(point):
	global sheet
	if isPointInMergePoints(point) and sheet.cell(point[0], point[1]).ctype != 0:
		return True
	return False

def getStartRow():
	global MergePoints
	rowMin = 0
	for [row, col] in MergePoints:
		if row > rowMin:
			rowMin = row
	# 最后一行叶子节点不属于合并单元格
	return rowMin + 2

# 获取当前单元格所在合并单元格的结束列坐标
def getMergePointsEndCol(point):
	global sheet
	for [row, rowMax, col, colMax] in sheet.merged_cells:
		if point[0] == row and point[1] == col:
			return colMax - 1

def traversingByTree(node, row_values):
	global sheet
	global text
	if node.getChildren() == None:
		return
	else :
		for child in node.getChildren():
			if child.getChildren():
				text = text + child.getDes() + " = {"
				traversingByTree(child, row_values)
			else :
				[row, col] = child.getCoordinate()
				if row_values[col] != "":
					if isinstance(row_values[col], float):
						text = text + child.getDes() + " = " + str(int(row_values[col])) + ", "
					else :
						text = text + child.getDes() + " = " + row_values[col] + ","

		text = text + "}, "
def process_excel(src):
	global sheet
	global MergePoints
	global text

	data = open_excel(src)
	sheet = data.sheets()[0]
	n_rows = sheet.nrows #总行数
	n_cols = sheet.ncols #总列数
	MergePoints = generateMergePoints(sheet.merged_cells)
	# 生成树状解析目录
	tree = generateTree(n_rows, n_cols)	
	rowMin = getStartRow()
	for row in range(rowMin, n_rows):
		text = text + "\t{"
		traversingByTree(tree.getHead(), sheet.row_values(row))
		if row < n_rows - 1:
			text = text + "\n"

def open_excel(file):
	try:
		data = xlrd.open_workbook(file)
		return data
	except Exception as e:
		print str(e)

# 此处配置excel文件名和导出的lua文件名
configDic = {
	"test.xlsx": "test.lua",
}

def main():
	global text
	for key in configDic:
		text = "local config = {\n"
		process_excel(key)
		text = text + "\n}\nreturn config"
		file = codecs.open(configDic[key], "w", "utf-8")
		file.write(text)
		file.close()
		print "------->" + configDic[key] + " :导出成功"

if __name__ == '__main__':
	main()