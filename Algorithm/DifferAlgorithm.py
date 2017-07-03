#coding=utf-8
"""
区分算法 核心应用LCS 最长公共子序列匹配
两次修改匹配结果，修复拼接行列引发的断层问题，以及修改获得最大匹配结果
"""
from GUI.TraceProgress import emitToProgressBar

class DifferAlgorithm(object):
    """
    ExcelDiff核心算法
    """
    @staticmethod
    def getDifferCellMatrix(workbook1,workbook2,sheet1,sheet2):
        """
        判断两个excel sheet不同部分
        :param sheet1: 表格1
        :param sheet2: 表格2
        :return CellMatrix: 返回4个包含原始信息的字典  preMatrixByRow,postMatrixByRow,preMatrixByCol,postMatrixByCol
        preMatrixByRow  原excel行区分信息
        postMatrixByRow 改动后excel行区分信息
        preMatrixByCol 原excel 列区分信息
        postMatrixByCol 改动后excel列区分信息
        GUIMessage 跟踪算法到GUI，装饰器不能实现，因为没办法emit发出信号
        """
        maxRow1,maxCol1 = workbook1.getSheetMaxSize(sheet1)
        maxRow2,maxCol2 = workbook2.getSheetMaxSize(sheet2)
        Arr1 = []
        Arr2 = []
        Arr3 = []
        Arr4 = []

        if((maxRow1 == 0 or maxCol1 == 0) and (maxRow2 == 0 or maxCol2 == 0)):
            return False,False,False,False,maxRow1,maxCol1,maxRow2,maxCol2 #全为空表
        elif(maxRow1 == 0 or maxCol1 == 0):   # 第一个excel为空
            sheet = []
            for i in range(maxRow2):
                col = []
                for j in range(maxCol2):
                    col.append((workbook2.getCellData(sheet2, i, j),(i,j),'add'))
                sheet.append(col)
            return False,sheet,False,sheet,maxRow1,maxCol1,maxRow2,maxCol2
        elif(maxRow2 == 0 or maxCol2 == 0):  #第二个excel为空
            sheet = []
            for i in range(maxRow1):
                col = []
                for j in range(maxCol1):
                    col.append((workbook1.getCellData(sheet1, i, j),(i,j),'del'))
                sheet.append(col)
            return sheet,False,sheet,False,maxRow1,maxCol1,maxRow2,maxCol2


        for i in range(maxRow1):
            for j in range(maxCol1):
                Arr1.insert(0,workbook1.getCellData(sheet1,i,j))
        for i in range(maxRow2):
            for j in range(maxCol2):
                Arr2.insert(0,workbook2.getCellData(sheet2,i,j))

        matrix = [[0 for col in range(maxRow2 * maxCol2 + 1)] for row in range(maxRow1 * maxCol1 + 1)]
        flag = [[0 for col in range(maxRow2 * maxCol2 + 1)] for row in range(maxRow1 * maxCol1 + 1)]

        DifferAlgorithm.LCS(Arr1,Arr2,matrix,flag)  #调用LCS算法求矩阵
        moveStatusByRow = DifferAlgorithm.moveTrailByLoop(len(Arr1),len(Arr2),Arr1,Arr2,flag,maxCol1, maxCol2) #获得运动轨迹
        preMatrixByRow,postMatrixByRow = DifferAlgorithm.classifyByRow(moveStatusByRow,maxCol1,maxCol2,len(Arr1),len(Arr2))

        #print moveStatusByRow

        DifferAlgorithm.cutoff_rule()  #行列算法分割线  --------------------------


        for i in range(maxCol1):
            for j in range(maxRow1):
                Arr3.insert(0,workbook1.getCellData(sheet1,j,i))
        for i in range(maxCol2):
            for j in range(maxRow2):
                Arr4.insert(0,workbook2.getCellData(sheet2,j,i))
        matrix = [[0 for _ in range(maxRow2 * maxCol2 + 1)] for _ in range(maxRow1 * maxCol1 + 1)]
        flag = [[0 for _ in range(maxRow2 * maxCol2 + 1)] for _ in range(maxRow1 * maxCol1 + 1)]

        DifferAlgorithm.LCS(Arr3, Arr4, matrix, flag)  # 调用LCS算法求矩阵
        moveStatusByCol = DifferAlgorithm.moveTrailByLoop(len(Arr3),len(Arr4),Arr3,Arr4,flag,maxRow1, maxRow2)#获得运动轨迹
        preMatrixByCol,postMatrixByCol = DifferAlgorithm.classifyByCol(moveStatusByCol,maxRow1,maxRow2,len(Arr3),len(Arr4))

        #print moveStatusByCol

        return preMatrixByRow,postMatrixByRow,preMatrixByCol,postMatrixByCol,maxRow1,maxCol1,maxRow2,maxCol2

    @staticmethod
    @emitToProgressBar
    def LCS(Arr1,Arr2,matrix,flag):
        for i in range(len(Arr1)+1):
            matrix[i][0] = 0
            flag[i][0] = 'up'   #方向也需要初始化，需要得到所有信息，而只不是匹配行列信息
        for j in range(len(Arr2)+1):
            matrix[0][j] = 0
            flag[0][j] = 'left'
        for i in range(1,len(Arr1)+1):
            for j in range(1,len(Arr2)+1):
                if(Arr1[i-1] == Arr2[j-1] != ''):
                    matrix[i][j] = matrix[i - 1][j - 1] + 1
                    flag[i][j] = "left_up"
                else:
                    if(matrix[i - 1][j] >= matrix[i][j - 1]):
                        matrix[i][j] = matrix[i - 1][j]
                        flag[i][j] = "up"
                    else:
                        matrix[i][j] = matrix[i][j - 1]
                        flag[i][j] = "left"

    @staticmethod
    @emitToProgressBar
    def moveTrailByLoop(i, j, Arr1, Arr2, flag,preMaxColOrRowCount,postMaxColOrRowCount):
        """
        找出箭头的运动轨迹，采用非递归方式找到回溯路径，同时对断行问题进行修正
        :param i: 输入的回溯法开始坐标
        :param j:  输入的回溯法开始坐标
        :param Arr1: 数据list
        :param Arr2: 数据list
        :param flag: 箭头标记方向
        :param preMaxColOrRowCount: 换行大小
        :param postMaxColOrRowCount: 换行大小
        :return: 返回具有方向运动轨迹的数组(值，矩阵Row，矩阵Col，方向)
        """
        nRow = i
        nCol = j
        lastLeft_up = ()  #记录上一个Left_up坐标，同时用来解决断行问题
        ret = []
        while(i != 0 or j != 0):
            if (flag[i][j] == "left_up"):
                if(lastLeft_up):
                    if ((nRow - i) / preMaxColOrRowCount != (nRow - lastLeft_up[0]) / preMaxColOrRowCount
                        and (nCol - j) / postMaxColOrRowCount == (nCol - lastLeft_up[1]) / postMaxColOrRowCount):
                        flag[i][j] = 'left'
                        continue
                    elif ((nRow - i) / preMaxColOrRowCount == (nRow - lastLeft_up[0]) / preMaxColOrRowCount
                          and (nCol - j) / postMaxColOrRowCount != (nCol - lastLeft_up[1]) / postMaxColOrRowCount):
                        flag[i][j] = 'up'
                        continue
                lastLeft_up = (i,j)
                moveStatus = (Arr2[j - 1], i - 1, j - 1, flag[i][j])
                i = i - 1
                j = j - 1
            else:
                if(flag[i][j] == "up"):
                    moveStatus = (Arr1[i - 1], i - 1, j - 1, flag[i][j])
                    i = i - 1
                else:
                    moveStatus = (Arr2[j - 1], i - 1, j - 1, flag[i][j])
                    j = j - 1
            ret.append(moveStatus)
        return ret

    @staticmethod
    @emitToProgressBar
    def classifyByRow(moveTrail,maxCol1,maxCol2,nRow,nCol,reverse = False):
        """
        以行匹配为例
        :param moveTrail: 匹配路径
        :param maxCol1: 原excel 列数
        :param maxCol2: 原excel 行数
        :param nRow: 组成动态规划矩阵的行数
        :param nCol: 组成动态规划矩阵的列数
        :param reverse:
        :return:
        """
        countCol1 = 0    #前excel列计数
        countCol2 = 0    #后excel列计数
        preDisplayMatrix = {}   #原excel展示矩阵,字典存储为{index:ArrayRow}的形式
        postDisplayMatrix = {}  #后excel展示矩阵
        currentPreDisplayRow = -1   #展示矩阵当前行索引
        currentPostDisplayRow = -1
        delRowCount = 0   #删除行统计初始化
        addRowCount = 0   #增加行统计初始化


        def matrixPos2ExcelPos(matrixPos,maxCol,nRowOrCol,reverse = False):
            """
            :param matrixPos: 行列当前长度位置
            :param maxCol:   最大列数
            :param nRowOrCol 组成矩阵最大行或最大列，这个是len(arr)必须减去1等到当前最大下标，用于转换到excel真实坐标
            :return:  cell在excel中的行号列号
            """
            realPos = nRowOrCol - 1 - matrixPos
            row = realPos / maxCol
            col = realPos % maxCol
            if(reverse):
                return col,row
            else:
                return row,col

        def isAllRowChanged(arr,mode):
            """
            全行都改变返回True，否则返回False
            :param arr: 行矩阵
            :param mode:  是否是add或del
            :return:
            """
            flag = False
            for cell in arr:
                if(cell[2] == mode):
                    flag = True
                else:
                    flag = False
                    break
            return flag

        def isAllRowSpace(arr,nCol):
            """
            判断是否整个行都是空行
            :param arr:输入数组
            :param nCol:列数，用于判断是否是满的状态
            :return: 是返回True 否则是False
            """
            if(len(arr) < nCol):
                return False
            flag = False
            for cell in arr:
                if(cell[0] == '' and (cell[2] == 'add' or cell[2] == 'del')):
                    flag = True
                else:
                    flag = False
                    break
            return flag

        def rowSpaceMatch(arr):
            """
            修改此行为remain
            :param arr:
            :return:
            """
            for i in range(len(arr)):
                arr[i] = (arr[i][0],arr[i][1],'remain')


        for index,status in enumerate(moveTrail):
            value,matrixRowPos,matrixColPos,direction = status
            preCellPos = matrixPos2ExcelPos(matrixRowPos,maxCol1,nRow,reverse)
            postCellPos = matrixPos2ExcelPos(matrixColPos,maxCol2,nCol,reverse) #转换为excel中的坐标

            if(direction == 'left_up'):
                if (countCol1 == 0):  # 换行 在展示字典中初始化新的一行
                    countCol1 = maxCol1
                    currentPreDisplayRow = currentPreDisplayRow + 1
                    preDisplayMatrix[currentPreDisplayRow] = []
                if (countCol2 == 0):
                    countCol2 = maxCol2
                    currentPostDisplayRow = currentPostDisplayRow + 1
                    postDisplayMatrix[currentPostDisplayRow] = []

                countCol1 = countCol1 - 1
                countCol2 = countCol2 - 1

                if(currentPreDisplayRow != currentPostDisplayRow): #强制修改匹配行为同一行
                    if(currentPreDisplayRow > currentPostDisplayRow):
                        checkRowExist = currentPostDisplayRow
                        postDisplayMatrix[currentPreDisplayRow] = postDisplayMatrix[currentPostDisplayRow]
                        del postDisplayMatrix[currentPostDisplayRow]
                        currentPostDisplayRow = currentPreDisplayRow
                        if (checkRowExist not in preDisplayMatrix.keys() and checkRowExist not in postDisplayMatrix):
                            # 由于以前被占据，可能已经将另一边挤下去了，此时自己再跳行，导致中间空出一行，补回
                            for key in range(checkRowExist + 1, currentPreDisplayRow + 1):
                                if (key in preDisplayMatrix.keys()):
                                    preDisplayMatrix[key - 1] = preDisplayMatrix[key]
                                    del preDisplayMatrix[key]
                                if (key in postDisplayMatrix.keys()):
                                    postDisplayMatrix[key - 1] = postDisplayMatrix[key]
                                    del postDisplayMatrix[key]
                            currentPostDisplayRow = currentPostDisplayRow - 1
                            currentPreDisplayRow = currentPreDisplayRow - 1

                    elif(currentPreDisplayRow < currentPostDisplayRow):
                        checkRowExist = currentPreDisplayRow
                        preDisplayMatrix[currentPostDisplayRow] = preDisplayMatrix[currentPreDisplayRow]
                        del preDisplayMatrix[currentPreDisplayRow]
                        currentPreDisplayRow = currentPostDisplayRow
                        if(checkRowExist not in preDisplayMatrix.keys() and checkRowExist not in postDisplayMatrix):
                            # 由于以前被占据，可能已经将另一边挤下去了，此时自己再跳行，导致中间空出一行，补回
                            for key in range(checkRowExist + 1,currentPreDisplayRow + 1):
                                if(key in preDisplayMatrix.keys()):
                                    preDisplayMatrix[key - 1] = preDisplayMatrix[key]
                                    del preDisplayMatrix[key]
                                if(key in postDisplayMatrix.keys()):
                                    postDisplayMatrix[key - 1] = postDisplayMatrix[key]
                                    del postDisplayMatrix[key]
                            currentPostDisplayRow = currentPostDisplayRow - 1
                            currentPreDisplayRow = currentPreDisplayRow - 1

                preDisplayMatrix[currentPreDisplayRow].append((value,preCellPos,'remain'))
                postDisplayMatrix[currentPostDisplayRow].append((value,postCellPos,'remain'))

            elif(direction == 'up'):
                if (countCol1 == 0):  # 换行 在展示字典中初始化新的一行
                    countCol1 = maxCol1
                    currentPreDisplayRow = currentPreDisplayRow + 1
                    preDisplayMatrix[currentPreDisplayRow] = []
                countCol1 = countCol1 - 1
                preDisplayMatrix[currentPreDisplayRow].append((value, preCellPos, 'del'))
                if(countCol1 == 0 and isAllRowChanged(preDisplayMatrix[currentPreDisplayRow],'del')):
                    isAllSpace = isAllRowSpace(preDisplayMatrix[currentPreDisplayRow], maxCol1)
                    while(currentPreDisplayRow in postDisplayMatrix.keys()): #当需要全部del时，发现此行已经被占据，将自己移下去
                        if(isAllSpace and isAllRowSpace(postDisplayMatrix[currentPreDisplayRow],maxCol2)):
                            rowSpaceMatch(postDisplayMatrix[currentPostDisplayRow])
                            rowSpaceMatch(preDisplayMatrix[currentPostDisplayRow])
                            break
                        else:
                            preDisplayMatrix[currentPreDisplayRow + 1] = preDisplayMatrix[currentPreDisplayRow]
                            del preDisplayMatrix[currentPreDisplayRow]
                            currentPreDisplayRow = currentPreDisplayRow + 1
            else:
                if (countCol2 == 0):
                    countCol2 = maxCol2
                    currentPostDisplayRow = currentPostDisplayRow + 1
                    postDisplayMatrix[currentPostDisplayRow] = []
                countCol2 = countCol2 - 1
                postDisplayMatrix[currentPostDisplayRow].append((value, postCellPos, 'add'))
                if(countCol2 == 0 and isAllRowChanged(postDisplayMatrix[currentPostDisplayRow],'add')):
                    isAllSpace = isAllRowSpace(postDisplayMatrix[currentPostDisplayRow],maxCol2)
                    while(currentPostDisplayRow in preDisplayMatrix.keys()):#当全部为add，发现此行被占据，自己挤下去！
                        if(isAllSpace and isAllRowSpace(preDisplayMatrix[currentPostDisplayRow],maxCol1)):
                            rowSpaceMatch(postDisplayMatrix[currentPostDisplayRow])
                            rowSpaceMatch(preDisplayMatrix[currentPostDisplayRow])
                            break
                        else:
                            postDisplayMatrix[currentPostDisplayRow + 1] = postDisplayMatrix[currentPostDisplayRow]
                            del postDisplayMatrix[currentPostDisplayRow]
                            currentPostDisplayRow = currentPostDisplayRow + 1
        return preDisplayMatrix,postDisplayMatrix

    @staticmethod
    @emitToProgressBar
    def classifyByCol(moveTrail,maxRow1,maxRow2,nRow,nCol):
        return DifferAlgorithm.classifyByRow(moveTrail,maxRow1,maxRow2,nRow,nCol,True)

    @staticmethod
    @emitToProgressBar
    def cutoff_rule():
        """
        行列信息的分割线，算法中无意义，主要用于在progressBar中输出label信息
        :return:
        """
        pass

    @staticmethod
    @emitToProgressBar
    def getOriginalMatrix(workbook,sheetName):
        """
        直接返回excel原始数据构造的矩阵
        :param sheetName:
        :return:
        """
        sheet = []
        nRow,nCol = workbook.getSheetMaxSize(sheetName)
        for i in range(nRow):
            col = []
            for j in range(nCol):
                col.append(workbook.getCellData(sheetName, i, j))
            sheet.append(col)
        return sheet
















