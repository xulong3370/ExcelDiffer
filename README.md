# ExcelDiffer
## ExcelDiffer based on Python and LCS algorithm <br>
require pyQt4 py2exe xlrd <br>
基于python2.7 实现的 excel differ <br>
UI 使用到了pyQt4<br>
读写excel 使用到了 xlrd<br>
打包使用 py2exe <br>

算法基于LCS最长公共子序列，分别拼接行、列单独进行LCS算法遍历<br>
寻找最优路径解决断层引发的bug<br>
最后行列还原数组获得最终结果
