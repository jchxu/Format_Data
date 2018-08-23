# coding=utf-8
import xlrd

### 返回字典{合并的单元格坐标:合并单元格左上角坐标(即合并前数值所在单元格的坐标)}
def UnMergeCell(MergedCellsList):
    UnMergeIndexDict = {}
    for item in MergedCellsList:
        rowlow = item[0]
        rowhigh = item[1]
        collow = item[2]
        colhigh = item[3]
        for i in range(rowlow,rowhigh):
            for j in range(collow,colhigh):
                UnMergeIndexDict[(i,j)] = (rowlow,collow)
    return (UnMergeIndexDict)

### 查找标题行、货主列、品种列、数量列、到港日期列的index ###
def FindIndex(Sheet):
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = (0,0,0,0,0)
    for i in range(Sheet.nrows):
        if '货主' in Sheet.row_values(i):
            print(Sheet.row_values(i))

            break


    return (TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex)






PortFilename = '连云港港存贸易矿8.12.xls'
PortFilename = '岚山现货表 08.14.xls'
PortFile = xlrd.open_workbook(PortFilename, formatting_info=True)


Sheets = PortFile.sheets()
for i in range(len(Sheets)):
    Sheet = PortFile.sheet_by_index(i)
    MergedCells = Sheet.merged_cells
    print(Sheet.name)
    print(MergedCells)
    UnMergeIndexDict = UnMergeCell(MergedCells)
    TitleRowIndex, OwnColIndex, GoodColIndex, AmountColIndex, DateColIndex = FindIndex(Sheet)


PortFile.release_resources()