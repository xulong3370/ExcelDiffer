#coding=utf-8
"""
跟踪算法中的代码，一个装饰器，将代码运行状态实时反映到进度条中，并且控制算法尽快正常在线程中结束
"""
from PyQt4 import QtCore
instanceGUIMessage = None
progressLabel = {'LCS':u"算法匹配中...",'moveTrailByLoop':u'寻优路径中...'
    ,'classifyByRow':u'获得匹配矩阵...','analysis':u'即将完成当前计算...'
    ,'getOriginalMatrix':u'统计sheet级别增删结果'}
rowOrCol = u'行匹配：'

def setGUIMessage(GUIMessage):
    """
    传入当前线程的实例对象，启动装饰器的功能
    :param GUIMessage:
    :return:
    """
    global instanceGUIMessage
    instanceGUIMessage = GUIMessage
    global rowOrCol
    rowOrCol = u'行匹配：'

def emitToProgressBar(func):
    """
    装饰器，控制progress响应和信息，控制线程尽快正常退出，防止线程过多卡顿
    :param func:
    :return:
    """
    def _deco(*args, **kwargs):
        if ((not instanceGUIMessage) or (instanceGUIMessage and not instanceGUIMessage.cancelFlag)):
            if(instanceGUIMessage and not instanceGUIMessage.cancelFlag):
                if(func.__name__ == 'cutoff_rule'):
                    global rowOrCol
                    rowOrCol = u'列匹配：'
                if(func.__name__ == 'analysis'):
                    labelStr = progressLabel.setdefault(func.__name__,u'计算中...')
                else:
                    labelStr = rowOrCol + progressLabel.setdefault(func.__name__,u'计算中...')
                instanceGUIMessage.emit(QtCore.SIGNAL('updateProgressMessage(QString)'),labelStr)#修改进度条，追踪当前算法

            ret = func(*args, **kwargs)  #不存在GUIMessage 或者 存在message但Flag不为False才执行算法
            if(instanceGUIMessage and func.__name__ in progressLabel.keys()
               and not instanceGUIMessage.cancelFlag ):
                instanceGUIMessage.emit(QtCore.SIGNAL('updateProgressbar(int)'),10) #进度条增加10
            return ret
        else:
            return None,None  #非正常结束算法，调用返回None值，不执行func
    return _deco
