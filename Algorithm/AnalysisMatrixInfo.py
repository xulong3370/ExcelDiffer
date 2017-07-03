#coding=utf-8
"""
解析字典模块
"""
import string
from GUI.TraceProgress import emitToProgressBar
class AnalysisMatrixInfo(object):
    """
    解析字典成成份，构建展示矩阵
    """
    @staticmethod
    @emitToProgressBar
    def analysis(preMatrixByRow,postMatrixByRow,preMatrixByCol,postMatrixByCol,preNRow,preNCol,postNRow,postNCol):
        """
        通过两个输入由行和列构成的字典信息，构建观测矩阵而不是字典
        :param preMatrixByRow:   原excel矩阵的行字典
        :param preMatrixByCol:   列字典
        :param postMatrixByRow:  后excel矩阵的行字典
        :param postMatrixByCol:  列字典
        :param preNRow: 原excel的 size 注意不是构建的matrix
        :param preNCol:
        :param postNRow:后excel的size
        :param postNCol:
        :return:
        """
        if((not preMatrixByRow and not preMatrixByCol) and (not postMatrixByRow and not postMatrixByCol)):
            return CorrelationMatrix([[]],[[]],[],[],[],[],[],-1,-1)   #两个excel为空
        elif(not preMatrixByRow and not preMatrixByCol):  #第一个为空
            postMatrix = postMatrixByCol  #此时返回的postMatrixByCol  postMatrixByRow 都是正常矩阵结果，不再是字典
            nRow = len(postMatrix)
            nCol = len(postMatrix[0])
            if(nRow <= nCol):
                addRowList = [i for i in range(nRow)]
                addColList = []
            else:
                addColList = [j for j in range(nCol)]
                addRowList = []
            preMatrix = [[('CHANGED', ('CHANGED', 'CHANGED'), 'CHANGED') for _ in range(nCol)] for _ in
                         range(nRow)]
            return CorrelationMatrix(preMatrix,postMatrix,[],[],addRowList,addColList,[],nRow - 1,nCol - 1)
        elif(not postMatrixByRow and not postMatrixByCol):  #第二个为空
            preMatrix =  preMatrixByCol
            nRow = len(preMatrix)
            nCol = len(preMatrix[0])
            if(nRow <= nCol):
                delRowList = [i for i in range(nRow)]
                delColList = []
            else:
                delColList = [j for j in range(nCol)]
                delRowList = []
            postMatrix = [[('CHANGED', ('CHANGED', 'CHANGED'), 'CHANGED') for _ in range(nCol)] for _ in
                         range(nRow)]
            return CorrelationMatrix(preMatrix,postMatrix,delRowList,delColList,[],[],[],nRow - 1,nCol - 1)

        maxRow = max(max(preMatrixByRow),max(postMatrixByRow))
        maxCol = max(max(preMatrixByCol),max(postMatrixByCol))

        #获得前后展示矩阵
        preMatrix = AnalysisMatrixInfo.makeMatrix(preMatrixByRow,preMatrixByCol,maxRow,maxCol)
        postMatrix = AnalysisMatrixInfo.makeMatrix(postMatrixByRow, postMatrixByCol, maxRow, maxCol)

        delRowList, delColList, addRowList, addColList = \
            AnalysisMatrixInfo.getChangedRowAndCol(preMatrixByRow,postMatrixByRow,
                                                   preMatrixByCol,postMatrixByCol) #获得增删坐标list

        if(len(delRowList)  == preNRow and len(addRowList) == postNRow
           and len(delColList) == preNCol and  len(addColList)  == postNCol): #表示全部删除又重建了，所以行列信息冲突，必须舍去一个
            nCol = len(preMatrix[0])
            if((preNRow + postNRow) <= (preNCol + postNCol)):
                for j in reversed(delColList):
                    for cell in postMatrix:
                        cell.pop(j)  #删除列
                for j in reversed(addColList):
                    for cell in preMatrix:
                        cell.pop(j)
                if (preNCol > postNCol):
                    for cell in postMatrix:
                        for _ in range(preNCol - postNCol):
                            cell.append(('CHANGED', ('CHANGED', 'CHANGED'), 'CHANGED'))
                else:
                    for cell in preMatrix:
                        for _ in range(postNCol - preNCol):
                            cell.append(('CHANGED', ('CHANGED', 'CHANGED'), 'CHANGED'))
                delColList = []
                addColList = []
                maxCol = max(preNCol,postNCol) - 1#减一的原因是maxCol参数指的是最大值，而不是长度。最大值等于长度减一
            else:
                for i in reversed(delRowList):
                    postMatrix.pop(i)
                for i in reversed(addRowList):
                    preMatrix.pop(i)

                if (preNRow > postNRow):
                    for _ in range(preNRow - postNRow):
                        postMatrix.append(
                            [('CHANGED', ('CHANGED', 'CHANGED'), 'CHANGED') for _ in range(nCol)])
                else:
                    for _ in range(postNRow - preNRow):
                        preMatrix.append(
                            [('CHANGED', ('CHANGED', 'CHANGED'), 'CHANGED') for _ in range(nCol) ])
                delRowList = []
                addRowList = []
                maxRow = max(preNRow,postNRow) - 1  #减一的原因是maxRow参数指的是最大值，而不是长度。最大值等于长度减一

        modifyList = AnalysisMatrixInfo.getChangedCellList(preMatrix,postMatrix,delRowList,addRowList,delColList,addColList,
                                                           maxRow,maxCol) #通过原矩阵字典求单独修改的坐标list

        return CorrelationMatrix(preMatrix,postMatrix,delRowList,delColList,addRowList,addColList,
                                 modifyList,maxRow,maxCol)


    @staticmethod
    def makeMatrix(matrixByRow,matrixByCol,maxRow,maxCol):
        """
        将字典行列信息 构建成一个矩阵
        :param matrixByRow:
        :param matrixByCol:
        :param maxRow:
        :param maxCol:
        :return:返回一个展示所用矩阵，删除行或列信息
        """
        retMatrix = [[('CHANGED',('CHANGED','CHANGED'),'CHANGED') for _ in range(maxCol+1)] for _ in range(maxRow+1)]  #构建最大矩阵
        def getColKey(rowValue):
            """
            根据excel唯一坐标值返回在矩阵中的列
            :param value:
            :return:第几列
            """
            for key,col in matrixByCol.iteritems():
                for colValue in col:
                    if(colValue[1] == rowValue[1]):
                        return key
            raise KeyError,u'算法出现未知错误，两个矩阵信息字典不同'

        for key,row in matrixByRow.iteritems():
            for value in row:
                colIndex = getColKey(value)
                rowIndex = key
                retMatrix[rowIndex][colIndex] = value

        return retMatrix

    @staticmethod
    def getChangedRowAndCol(preMatrixByRow,postMatrixByRow,preMatrixByCol,postMatrixByCol):
        """
        获得删除，增加的行和列在展示矩阵中的index
        :param preMatrixByRow:
        :param postMatrixByRow:
        :param preMatrixByCol:
        :param postMatrixByCol:
        :return:
        """
        addRowList = []
        addColList = []
        delRowList = []
        delColList = []
        def isChanged(array,mode):
            flag = False
            for cell in array:
                if(cell[2] == mode):
                    flag =True
                else:
                    flag = False
                    break
            return flag

        for key,array in preMatrixByRow.iteritems():
            if(isChanged(array,'del')):
                delRowList.append(key)
        for key,array in preMatrixByCol.iteritems():
            if(isChanged(array,'del')):
                delColList.append(key)
        for key,array in postMatrixByRow.iteritems():
            if(isChanged(array,'add')):
                addRowList.append(key)
        for key,array in postMatrixByCol.iteritems():
            if(isChanged(array,'add')):
                addColList.append(key)

        return delRowList,delColList,addRowList,addColList

    @staticmethod
    def getChangedCellList(preMatrix,postMatrix,delRowList,addRowList,delColList,addColList,maxRow,maxCol):
        modifyList = []
        for i in range(maxRow+1):
            if(i in delRowList or i in addRowList):
                continue
            else:
                for j in range(maxCol+1):
                    if(j in delColList or j in addColList):
                        continue
                    else:
                        if(preMatrix[i][j][0] != postMatrix[i][j][0]):
                            modifyList.append((i,j))

        return modifyList

class CorrelationMatrix(object):
    """
    相关矩阵类型，保存最终结果，包含原矩阵和结果矩阵，删除行列信息，修改单独Cell信息
    """
    def __init__(self,preMatrix,postMatrix,delRow,delCol,addRow,addCol,modifyCell,maxRow,maxCol):
        self.preMatrix = preMatrix
        self.postMatrix = postMatrix
        self.modifyCellInfo = {}
        self.maxNRow = maxRow + 1    #展示矩阵行列数目
        self.maxNCol = maxCol + 1

        self.preRowHeadList,self.preColHeadList,self.postRowHeadList,\
        self.postColHeadList = self.__headListAndColList(maxRow,maxCol,preMatrix,postMatrix,
                                                         delRow,delCol,addRow,addCol)

        self.delColInfo, self.addColInfo,self.delRowInfo, self.addRowInfo = \
            self.__changedRowAndColList(delRow,delCol,addRow,addCol)

        for pos in modifyCell:
            i,j = pos
            rowPre, colPre = self.__matrixPos2ExcelPos(pos,self.preRowHeadList,self.preColHeadList)
            rowPost, colPost = self.__matrixPos2ExcelPos(pos, self.postRowHeadList, self.postColHeadList)

            self.modifyCellInfo[pos] = (self.preMatrix[i][j][0],(rowPre, colPre),
                                        self.postMatrix[i][j][0],(rowPost, colPost))
            #self.modifyCellInfo字典中的value依次为 改变前数据，改变前（x1,y1）坐标，改变后数据，改变后(x2,y2)坐标

    def __headListAndColList(self,maxRow,maxCol,preMatrix,postMatrix,delRow,delCol,addRow,addCol):
        """
        :param maxRow:
        :param maxCol:
        :param preMatrix:
        :param postMatrix:
        :param delRow:
        :param delCol:
        :param addRow:
        :param addCol:
        :return:  返回表头信息，以及excel原表的size大小
        """
        preColList = []
        preRowList = []
        postColList = []
        postRowList = []

        for i in range(maxRow + 1):
            flag = False
            if(i not in delRow):
                for j in range(maxCol + 1):
                    postHead = postMatrix[i][j][1][0]
                    if(postHead != 'CHANGED'):
                        postColList.append(str(postHead + 1))
                        flag = True
                        break
                if(not flag):
                    postColList.append('CHANGED')
            else:
                postColList.append('CHANGED')

            flag = False
            if(i not in addRow):
                for j in range(maxCol + 1):
                    preHead = preMatrix[i][j][1][0]
                    if(preHead != 'CHANGED'):
                        preColList.append(str(preHead + 1))
                        flag = True
                        break
                if(not flag):
                    preColList.append('CHANGED')
            else:
                preColList.append('CHANGED')


        for j in range(maxCol + 1):
            flag = False
            if(j not in addCol):
                for i in range(maxRow + 1):
                    preHead = preMatrix[i][j][1][1]
                    if(preHead != 'CHANGED'):
                        preRowList.append(string.uppercase[preHead])
                        flag = True
                        break
                if(not flag):
                    preRowList.append('CHANGED')
            else:
                preRowList.append('CHANGED')

            flag = False
            if(j not in delCol):
                for i in range(maxRow + 1):
                    postHead = postMatrix[i][j][1][1]
                    if(postHead != 'CHANGED'):
                        postRowList.append(string.uppercase[postHead])
                        flag = True
                        break
                if(not flag):
                    postRowList.append('CHANGED')
            else:
                postRowList.append('CHANGED')

        return preRowList,preColList,postRowList,postColList

    def __changedRowAndColList(self,delRow,delCol,addRow,addCol):
        """
        将展示矩阵行列删除信息 与excel 行列名绑定
        :param delRow:
        :param delCol:
        :param addRow:
        :param addCol:
        :return: 返回展示矩阵与excel 行列原位置的对应信息
        """
        delColInfo = {}
        addColInfo = {}
        delRowInfo = {}
        addRowInfo = {}
        for i in delCol:
            delColInfo[i] = self.preRowHeadList[i]
        for i in addCol:
            addColInfo[i] = self.postRowHeadList[i]
        for i in delRow:
            delRowInfo[i] = self.preColHeadList[i]
        for i in addRow:
            addRowInfo[i] = self.postColHeadList[i]
        return delColInfo, addColInfo, delRowInfo, addRowInfo

    def __matrixPos2ExcelPos(self,pos,rowHeadList,colHeadList):
        """
        将matrix中删除和增加的位置信息转化为excel中位置信息
        :param pos:
        :param rowHeadList:
        :param colHeadList:
        :return:
        """
        return colHeadList[pos[0]], rowHeadList[pos[1]]



    def __repr__(self):
        preRowHeadStr = '           '
        for j in range(self.maxNCol):
            preRowHeadStr = preRowHeadStr + "    {0}    ||".format(str(self.preMatrix[0][j][1][1]))

        postRowHeadStr = '           '
        for j in range(self.maxNCol):
            postRowHeadStr = postRowHeadStr + "    {0}    ||".format(str(self.postMatrix[0][j][1][1]))


        preMatrixStr = ''
        postMatrixStr = ''
        for i in range(self.maxNRow):
            preMatrixStr = preMatrixStr + "    {0}    ||".format(str(self.preMatrix[i][0][1][0]))
            for j in range(self.maxNCol):
                preMatrixStr = preMatrixStr + "  {0}  ||".format(str(self.preMatrix[i][j][0]))
            preMatrixStr = preMatrixStr + "\n"
        for i in range(self.maxNRow):
            postMatrixStr = postMatrixStr + "    {0}    ||".format(str(self.postMatrix[i][0][1][0]))
            for j in range(self.maxNCol):
                postMatrixStr = postMatrixStr + "  {0}  ||".format(str(self.postMatrix[i][j][0]))
            postMatrixStr = postMatrixStr + "\n"


        delCellStr = ''
        for index,value in  sorted(self.modifyCellInfo.iteritems()):
            i,j = index
            delCellStr = delCellStr + "原坐标{0}的{1}变为坐标{2}的{3}|".format(str(value[1]),str(value[0]),
                                                          str(value[3]),str(value[2]))

        Str = """
对比结果
原Excel：
{8}
{0}
{7}
现在Excel：
{9}
{1}
{7}
删除的行为：{2}
{7}
增加的行为：{3}
{7}
删除的列为：{4}
{7}
增加的列为：{5}
{7}
修改的几个部分元素为：{6}""".format(preMatrixStr,postMatrixStr,str(self.delRowInfo),str(self.addRowInfo),
                         str(self.delColInfo),str(self.addColInfo),str(delCellStr),'-'*50,preRowHeadStr,postRowHeadStr)
        return Str
