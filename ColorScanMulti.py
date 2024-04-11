# 2023年2月21日13:37:05 根据之前写的ColorScanMulti改一个颜色检测软件出来。
# 检测列表中列出：
# 1.当前行是否检测的开关，当前检测点坐标，
# 2.检测区域的左上角坐标LX,LY  右下角坐标RX,RY。 颜色RGB值和RGB值的偏移量ROffset, GOffset, BOffset
# 3.检测时间间隔，闪屏开关， 声音开关， 选择声音文件的按钮， 声音文件路径信息, 备注。
# 按表格中的内容进行颜色检测，并进行声音和闪屏提醒。
# 可单线程顺序执行，也可多线程并发执行，线程数量可用spinBox选择。如果线程数量大于检测区域的数量，则选择较小值来创建线程

import sys,threading,os,json,datetime,time
from PyQt5.QtWidgets import QMainWindow,QApplication,QMessageBox,QFileDialog,QTableWidgetItem,QCheckBox,QPushButton,QAbstractItemView
from PyQt5.QtCore import pyqtSignal,QUrl,QTimer
from PyQt5.QtGui import QColor
import pyautogui as pag
from PyQt5 import QtMultimedia
from ctypes import windll
from Ui_ColorScanMulti import Ui_ColorScanMulti
 
# 定义全局常量
pag.FAILSAFE = True
ISRUN = 0   # 0是停止  1是运行

# 定义全局函数
# 判断颜色在指定区域是否存在, 存在返回True, 不存在返回False
# hdc屏幕设备上下文句柄, hdc = windll.user32.GetDC(None)
# X,Y是当前检测点的坐标,RGB是需要判定的颜色值,ROffset,GOffset,BOffset是颜色值的偏移量，
def isColorExist( hdc, X, Y, R, ROffset, G, GOffset, B, BOffset ):
    pixel = windll.gdi32.GetPixel( hdc, X, Y)
    r = pixel & 0x0000ff
    g = (pixel & 0x00ff00) >> 8
    b = pixel >> 16
    if r > R + ROffset or r < R - ROffset:
        return False
    if g > G + GOffset or g < G - GOffset:
        return False
    if b <= B + BOffset and b >= B - BOffset:
        return True

#获得鼠标位置的坐标和颜色   返回字典
def getMouseParam():
    posX, posY = pag.position()     #获取位置
    hdc = windll.user32.GetDC(None)  # 获取hdc 屏幕设备上下文句柄
    pixel = windll.gdi32.GetPixel(hdc, posX, posY)  # 提取RGB值
    r = pixel & 0x0000ff
    g = (pixel & 0x00ff00) >> 8
    b = pixel >> 16
    return {'x':posX, 'y':posY, 'r':r, 'g':g, 'b':b}



class ColorScanMulti( QMainWindow, Ui_ColorScanMulti): 
    #定义一个信号,用于子线程给主线程发信号
    signalCrossThread = pyqtSignal(str, str)     #两个str参数,第一个接收信号类型,第二个接收信号内容

    def __init__(self,parent =None):
        super( ColorScanMulti,self).__init__(parent)
        self.setupUi(self)

        # 初始化保存信息的表格  twColors
        tempHeaderText = ['开关','当前点', '左上X', '左上Y', '右下X', '右下Y', 'R', 'R色差', 'G', 'G色差', 'B', 'B色差',
                          '间隔(s)', '闪屏', '声音', '选择声音', '声音文件路径', '备注']
        self.twColors.setColumnCount( len(tempHeaderText))
        self.twColors.setHorizontalHeaderLabels( tempHeaderText)

        # 调整列宽
        self.twColors.setColumnWidth( 0, 40)
        self.twColors.setColumnWidth( 1, 100)
        self.twColors.setColumnWidth( 13, 40)
        self.twColors.setColumnWidth( 14, 40)
        self.twColors.setColumnWidth( 15, 80)
        self.twColors.setColumnWidth( 16, 120)
        self.twColors.setColumnWidth( 17, 160)
        for col in range(2,13):
            self.twColors.setColumnWidth( col, 60)

        #打开配置文件,初始化界面数据  配置文件为ColorScanMulti.ini  检测列表的内容保存为.csm后缀， 是ColorScanMulti的简写
        if os.path.exists( "./ColorScanMulti.ini"):
            try:
                iniFileDir = os.getcwd() + "\\"+ "ColorScanMulti.ini"
                with open( iniFileDir, 'r', encoding="utf-8") as iniFile:
                    iniDict = json.loads( iniFile.read())
                if iniDict:
                    if os.path.exists( iniDict['csmDir']):   # 如果检测列表内容文件位置存在,则用mfRefresh()初始化界面
                        self.labelDir.setText( iniDict['csmDir'])       # 显示配置文件的路径
                        self.sbThreadCount.setValue( iniDict['threadCount'])    # ini文件中保存设置的线程数量
                        self.sbAutoGetDelay.setValue( iniDict['autoGetDelay'])      # 自动获取鼠标信之前的延迟时间
                        self.teLog.setText( iniDict['log'])     # 显示teLog中的内容
                        if iniDict['noEdit'] == True:       # 设置检测列表是否可以编辑 和 cbNoEdit的状态
                            self.cbNoEdit.setChecked( True)
                            self.twColors.setEditTriggers(QAbstractItemView.NoEditTriggers)
                        else:
                            self.cbNoEdit.setChecked( False)
                            self.twColors.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)

                        if iniDict['threadMode'] == 'multi':    # 线程模式如果不是multi多线程，则都设置为单线程
                            self.rbMultiThread.setChecked( True)
                        else:
                            self.rbSingleThread.setChecked( True)

                        self.mfRefresh( iniDict['csmDir'])    
                    else:
                        QMessageBox.about( self, '提示', '配置文件ColorScanMulti.ini中的csmDir脚本文件位置错误或文件不存在')
            except:
                QMessageBox.about( self, "提示", "打开初始化文件ColorScanMulti.ini异常, 软件关闭时会自动重新创建ColorScanMulti.ini文件")

        # 绑定槽函数
        self.btnOpen.clicked.connect( self.mfOpen)
        self.btnSave.clicked.connect( self.mfSave)
        self.btnAdd.clicked.connect( self.mfAdd)
        self.btnDelete.clicked.connect( self.mfDelete)
        self.btnDeleteRows.clicked.connect( self.mfDeleteRows)
        self.btnClearActions.clicked.connect( self.mfClearColors)
        self.btnQuit.clicked.connect( self.mfQuit)
        self.btnStart.clicked.connect( self.mfStart)
        self.btnStop.clicked.connect( self.mfStop)
        self.btnHelp.clicked.connect( self.mfHelp)
        self.btnClearLog.clicked.connect( self.mfClearLog)
        self.btnInsert.clicked.connect( self.mfInsert)
        self.btnCopyRows.clicked.connect( self.mfCopyRows)
        self.btnNew.clicked.connect( self.mfNew)
        self.cbNoEdit.stateChanged.connect( self.mfNoEdit)
        self.btnAutoGet.clicked.connect( self.mfAutoGet)


        self.signalCrossThread.connect( self.mfSignal)       # 处理子线程给主线程发的信号

    # 槽函数定义
    # 点击使用帮助
    def mfHelp( self):
        QMessageBox.about( self, '使用帮助', 
        '点击退出按钮会保存检测表格中的数据,点击右上角叉号退出不会保存表格数据。  \
        \n\
        \n检测列表中数据不能为空,没有数据的要填上0,否则保存时会出现异常    \
        \n\
        \n表格中的时间间隔(单位秒s)是当前行检测完成后暂停多少秒   \
        \n\
        \n多线程模式下,如果指定线程数大于需要检测的行数,则线程数自动改为行数  \
        \n\
        \n软件思路:判断屏幕指定区域内是否存在指定颜色. \
        \n如果指定颜色存在则播放声音 闪屏报警, \
        \n检测区域坐标点和颜色RGB值可用Snipaste或微信截图工具来获得. \
        \n也可以点击自动获取按钮,指定时间后会自动获取鼠标位置的坐标或颜色RGB值.\
        \n色差±值表示指定颜色的范围.例如R=200  ±30,则判断R是否在170--230之间\
        \n被检测的屏幕区域不要遮挡. 检测区域不要选的太大, 否则一次检测会用很长时间.\
        \n报警声音可以用自己的wav音频.\
        '
        )


    # 处理子线程给主线程发的信号, 信号signalType是字符串'QMessageBox' 'Log' 'Flash去掉了，直接在新创建的线程里执行闪屏' 'Sound' 
    def mfSignal( self, signalType, content):
        if signalType == 'currentPos':
            row, pos = content.split('-')
            self.twColors.setItem( int(row), 1, QTableWidgetItem( pos))

        # elif signalType == 'Flash':       在主线程中闪屏会卡住
        #     row = int(content.split('-')[0])
        #     for i in range( 0, 15):
        #         print( 'in flash', i)
        #         self.twColors.item( row, 17).setForeground( QColor(255, 0, 0))
        #         time.sleep( 0.2)
        #         self.twColors.item( row, 17).setForeground( QColor(0, 0, 255))
        #         time.sleep( 0.2)

        elif signalType == 'Sound':
            # 播放警报声音
            row = int(content.split('-')[0])
            soundDir = self.twColors.item( row, 16).text()
            if soundDir == '...' or soundDir == '':
                return
            #这里应该加一个判断文件是否存在
            
            url = QUrl.fromLocalFile( soundDir)
            content = QtMultimedia.QMediaContent(url)
            # --------------------------------------------------------------------------------------------------
            # 注意这里的player前面要加self.   把player加入到这个类中
            # 因为QtMultimedia.QMediaPlayer() 必须在app = QApplication(sys.argv)之后才能播放声音
            # QtMultimedia会在主线程里加一个timer来播放声音, 没有app = QApplication(sys.argv)就没有主线程,没有信号队列
            # 单独写一个QtMultimedia.QMediaPlayer()是不能播放的
            #---------------------------------------------------------------------------------------------------
            self.player = QtMultimedia.QMediaPlayer()
            self.player.setMedia(content)
            self.player.setVolume(100)
            self.player.play()
            pass

        elif signalType == 'QMessageBox':
            QMessageBox.about( self, "提示", content)

        elif signalType == 'Display':
            self.teLog.append( content)


    # 定义 传入检测列表内容文件.csm位置,打开csm文件,刷新软件界面
    def mfRefresh( self, paramDir):
        # self.twColors.clearContents()    这个只清空数据,会留下空白的行. 这里需要清空所有行
        for i in range( 0, self.twColors.rowCount()):
            self.twColors.removeRow(0)

        with open( paramDir, 'r', encoding="utf-8") as csmFile:
            csmDict = json.loads( csmFile.read())
            if csmDict:
                i = 0 
                for key in csmDict:
                    self.twColors.insertRow( self.twColors.rowCount())
                    # 第一列是判断当前行是否检测的开关
                    tempCheckBoxSwitch = QCheckBox()
                    self.twColors.setCellWidget( i, 0, tempCheckBoxSwitch)
                    if csmDict[key]['switch'] == 1:
                        tempCheckBoxSwitch.setChecked( True)
                    else:
                        tempCheckBoxSwitch.setChecked( False)
                    #第二列是当前检测点， 默认是 0,0    表示还没有进行过检测
                    self.twColors.setItem( i, 1, QTableWidgetItem( '0,0'))
                    # LX,LY,RX,RY,R,ROffset,G,GOffset,B,BOffset,interval
                    self.twColors.setItem( i, 2, QTableWidgetItem( csmDict[key]['LX']))
                    self.twColors.setItem( i, 3, QTableWidgetItem( csmDict[key]['LY']))
                    self.twColors.setItem( i, 4, QTableWidgetItem( csmDict[key]['RX']))
                    self.twColors.setItem( i, 5, QTableWidgetItem( csmDict[key]['RY']))
                    self.twColors.setItem( i, 6, QTableWidgetItem( csmDict[key]['R']))
                    self.twColors.setItem( i, 7, QTableWidgetItem( csmDict[key]['ROffset']))
                    self.twColors.setItem( i, 8, QTableWidgetItem( csmDict[key]['G']))
                    self.twColors.setItem( i, 9, QTableWidgetItem( csmDict[key]['GOffset']))
                    self.twColors.setItem( i, 10, QTableWidgetItem( csmDict[key]['B']))
                    self.twColors.setItem( i, 11, QTableWidgetItem( csmDict[key]['BOffset']))
                    self.twColors.setItem( i, 12, QTableWidgetItem( csmDict[key]['interval']))
                    # 闪屏开关
                    tempCheckBoxFlash = QCheckBox()
                    self.twColors.setCellWidget( i, 13, tempCheckBoxFlash)
                    if csmDict[key]['flashSwitch'] == 1:
                        tempCheckBoxFlash.setChecked( True)
                    else:
                        tempCheckBoxFlash.setChecked( False)
                    # 声音开关
                    tempCheckBoxSound = QCheckBox()
                    self.twColors.setCellWidget( i, 14, tempCheckBoxSound)
                    if csmDict[key]['soundSwitch'] == 1:
                        tempCheckBoxSound.setChecked( True)
                    else:
                        tempCheckBoxSound.setChecked( False)
                    # 第15列是选择声音按钮
                    tempPushButton = QPushButton()
                    tempPushButton.setText('选择')
                    self.twColors.setCellWidget( i, 15, tempPushButton)
                    tempPushButton.clicked.connect( self.mfSelectSound)
                    # 第16列是声音路径
                    self.twColors.setItem( i, 16, QTableWidgetItem( csmDict[key]['soundDir']))
                    # 第17列是备注
                    self.twColors.setItem( i, 17, QTableWidgetItem( csmDict[key]['note']))

                    i = i + 1


    # 定义 打开(检测颜色文件.csm), 获得文件位置,并传递文件位置给mfRefresh(),用来刷新界面
    def mfOpen( self):
        try:
            tempDir, uselessFilt = QFileDialog.getOpenFileName( self, '选择检测颜色文件', os.getcwd(), '检测文件(*.csm)', '检测文件(*.csm)')
            if tempDir != '':
                self.labelDir.setText( tempDir)
                self.mfRefresh( tempDir)
            else:
                QMessageBox.about( self, "提示", "请选择后缀名为 .csm 的检测颜色文件。")
        except:
            QMessageBox.about( self, "提示", "打开检测颜色.csm文件失败,请重新选择。")


    #定义 保存界面上的数据
    def mfSave(self):
        saveDict = {}
        for i in range( 0, self.twColors.rowCount()):
            saveKey = 'colorScan' + str(i)
            tempDict = {}
            if self.twColors.cellWidget( i, 0).isChecked():
                tempDict['switch'] = 1
            else:
                tempDict['switch'] = 0
            # 当前检测点这一列不用保存，所以没有self.twColors.item( i, 1).text()
            tempDict['LX'] = self.twColors.item( i, 2).text()
            tempDict['LY'] = self.twColors.item( i, 3).text()
            tempDict['RX'] = self.twColors.item( i, 4).text()
            tempDict['RY'] = self.twColors.item( i, 5).text()
            tempDict['R'] = self.twColors.item( i, 6).text()
            tempDict['ROffset'] = self.twColors.item( i, 7).text()
            tempDict['G'] = self.twColors.item( i, 8).text()
            tempDict['GOffset'] = self.twColors.item( i, 9).text()
            tempDict['B'] = self.twColors.item( i, 10).text()
            tempDict['BOffset'] = self.twColors.item( i, 11).text()
            tempDict['interval'] = self.twColors.item( i, 12).text()

            if self.twColors.cellWidget( i, 13).isChecked():
                tempDict['flashSwitch'] = 1
            else:
                tempDict['flashSwitch'] = 0

            if self.twColors.cellWidget( i, 14).isChecked():
                tempDict['soundSwitch'] = 1
            else:
                tempDict['soundSwitch'] = 0
            # 没有第15列的操作  第15列是选择声音按钮
            tempDict['soundDir'] = self.twColors.item( i, 16).text()
            tempDict['note'] = self.twColors.item( i, 17).text()

            saveDict[saveKey] = tempDict
        
        saveJson = json.dumps( saveDict, indent=4)
        try:
            with open( self.labelDir.text(), 'w', encoding="utf-8") as saveFile:
                saveFile.write( saveJson)
            # QMessageBox.about( self, '提示', '保存成功  ' + self.labelDir.text())
        except:
            QMessageBox.about( self, "提示", "保存检测颜色文件失败。保存之前要新建文件或者打开已有的文件")

    # 定义 新建按钮,打开一个文件对话框获得保存路径,再打开一个对话框用于输入名称
    def mfNew( self):
        tempName = pag.prompt( text='输入新建检测文件.csm的名称', title='新建')
        if tempName == '' or tempName == None:
            pag.alert( text='脚本名称不可用')
        else:
            try:
                tempDir = QFileDialog.getExistingDirectory( self, '选择保存目录', os.getcwd(),QFileDialog.ShowDirsOnly)
                tempFileDir = tempDir + '/' + tempName + '.csm'
                f = open( tempFileDir, 'x')
                f.write('{}')
                f.close()
                self.labelDir.setText( tempFileDir)
                self.mfRefresh( tempFileDir)
            except FileExistsError:
                QMessageBox.about( self, '提示', '文件名称重复')
                return
            except:
                QMessageBox.about( self, '提示', '文件创建失败')
                return

    # 通过点击表格中的选择声音按钮，选择声音文件，把声音文件路径字符串填在声音文件路径单元格里    
    def mfSelectSound(self):
        # self.sender()是发送信号的QPushButton   .pos()是获得QPoint类和坐标   indexAt().row()获得按钮所在行的序号
        tempRow = self.twColors.indexAt(self.sender().pos()).row()
        # 打开选择文件对话框
        try:
            tempDir, uselessFilt = QFileDialog.getOpenFileName( self, '选择声音文件', os.getcwd(), '声音文件(*.wav)', '声音文件(*.wav)')
            if tempDir != '':
                self.twColors.setItem( tempRow, 16, QTableWidgetItem( tempDir))
            else:
                QMessageBox.about( self, "提示", "请选择后缀名为 .wav 的声音文件。")
        except:
            QMessageBox.about( self, "提示", "选择声音文件失败,请重新选择。")


    #给twColors表格添加一个空的行,用来添加新的检测
    def mfAdd( self):
        self.twColors.insertRow( self.twColors.rowCount())

        # 开关  判断是否检测当前行
        tempCheckBoxSwitch = QCheckBox()
        tempCheckBoxSwitch.setChecked( True)
        self.twColors.setCellWidget( self.twColors.rowCount()-1, 0, tempCheckBoxSwitch)
        # 当前检测点
        self.twColors.setItem( self.twColors.rowCount()-1, 1, QTableWidgetItem('0, 0'))
        # LX LY RX RY
        self.twColors.setItem( self.twColors.rowCount()-1, 2, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 3, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 4, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 5, QTableWidgetItem('0'))
        # RGB和偏移值 ROffset  GOffset  BOffset
        self.twColors.setItem( self.twColors.rowCount()-1, 6, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 7, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 8, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 9, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 10, QTableWidgetItem('0'))
        self.twColors.setItem( self.twColors.rowCount()-1, 11, QTableWidgetItem('0'))
        # 时间间隔  当前行检测完成后,间隔多少秒,再进行下一次当前行的检测
        self.twColors.setItem( self.twColors.rowCount()-1, 12, QTableWidgetItem('0'))
        # 闪屏开关
        tempCheckBoxFlash = QCheckBox()
        tempCheckBoxFlash.setChecked( True)
        self.twColors.setCellWidget( self.twColors.rowCount()-1, 13, tempCheckBoxFlash)
        # 声音开关
        tempCheckBoxSound = QCheckBox()
        tempCheckBoxSound.setChecked( True)
        self.twColors.setCellWidget( self.twColors.rowCount()-1, 14, tempCheckBoxSound)
        # 声音选择按钮
        tempPushButton = QPushButton()
        tempPushButton.setText('选择')
        self.twColors.setCellWidget( self.twColors.rowCount()-1, 15, tempPushButton)
        tempPushButton.clicked.connect( self.mfSelectSound)
        # 声音文件路径 和 备注
        self.twColors.setItem( self.twColors.rowCount()-1, 16, QTableWidgetItem('...'))
        self.twColors.setItem( self.twColors.rowCount()-1, 17, QTableWidgetItem('...'))

        self.twColors.scrollToBottom()     # twColors滚动到最后一行


    # 删除twActions当前所选的行
    def mfDelete( self):
        self.twColors.removeRow( self.twColors.currentRow())

    # 删除twActionis中 sbRowStart到sbRowEnd之间的行
    def mfDeleteRows( self):
        rowStart = self.sbRowDeleteStart.value()
        rowEnd = self.sbRowDeleteEnd.value()
        if rowStart > rowEnd or rowStart < 1:
            QMessageBox.about( self, '提示', '起止行参数错误')
        for i in range( rowStart, rowEnd + 1):
            self.twColors.removeRow( rowStart - 1)

    # 在当前行上面插入新行
    def mfInsert( self):
        currentRow = self.twColors.currentRow()
        self.twColors.insertRow( currentRow)

        # 开关  判断是否检测当前行
        tempCheckBoxSwitch = QCheckBox()
        tempCheckBoxSwitch.setChecked( True)
        self.twColors.setCellWidget( currentRow, 0, tempCheckBoxSwitch)
        # 当前检测点
        self.twColors.setItem( currentRow, 1, QTableWidgetItem('0, 0'))
        # LX LY RX RY
        self.twColors.setItem( currentRow, 2, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 3, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 4, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 5, QTableWidgetItem('0'))
        # RGB和偏移值 ROffset  GOffset  BOffset
        self.twColors.setItem( currentRow, 6, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 7, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 8, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 9, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 10, QTableWidgetItem('0'))
        self.twColors.setItem( currentRow, 11, QTableWidgetItem('0'))
        # 时间间隔  当前行检测完成后,间隔多少秒,再进行下一次当前行的检测
        self.twColors.setItem( currentRow, 12, QTableWidgetItem('0'))
        # 闪屏开关
        tempCheckBoxFlash = QCheckBox()
        tempCheckBoxFlash.setChecked( True)
        self.twColors.setCellWidget( currentRow, 13, tempCheckBoxFlash)
        # 声音开关
        tempCheckBoxSound = QCheckBox()
        tempCheckBoxSound.setChecked( True)
        self.twColors.setCellWidget( currentRow, 14, tempCheckBoxSound)
        # 声音选择按钮
        tempPushButton = QPushButton()
        tempPushButton.setText('选择')
        self.twColors.setCellWidget( currentRow, 15, tempPushButton)
        tempPushButton.clicked.connect( self.mfSelectSound)
        # 声音文件路径 和 备注
        self.twColors.setItem( currentRow, 16, QTableWidgetItem('...'))
        self.twColors.setItem( currentRow, 17, QTableWidgetItem('...'))


    #复制rowStart行到rowEnd行之间的操作列表到rowMark行后面
    def mfCopyRows( self):
        rowStart = self.sbRowCopyStart.value()
        rowEnd = self.sbRowCopyEnd.value()
        rowMark = self.sbRowCopyMark.value()
        # 这里需要一个参数值的合理性检测
        if rowStart < 1 or rowStart > rowEnd:
            QMessageBox.about( self, '提示', '复制操作的起止点参数错误')
            return

        tempColorsDict = {}       # 定义一个字典，存储所有要复制的行
        i = 0
        for k in range( rowStart - 1, rowEnd):
            tempDict = {}
            if self.twColors.cellWidget( k, 0).isChecked():
                tempDict['switch'] = 1
            else:
                tempDict['switch'] = 0
            # 当前检测点这一列不用保存，所以没有self.twColors.item( i, 1).text()
            tempDict['LX'] = self.twColors.item( k, 2).text()
            tempDict['LY'] = self.twColors.item( k, 3).text()
            tempDict['RX'] = self.twColors.item( k, 4).text()
            tempDict['RY'] = self.twColors.item( k, 5).text()
            tempDict['R'] = self.twColors.item( k, 6).text()
            tempDict['ROffset'] = self.twColors.item( k, 7).text()
            tempDict['G'] = self.twColors.item( k, 8).text()
            tempDict['GOffset'] = self.twColors.item( k, 9).text()
            tempDict['B'] = self.twColors.item( k, 10).text()
            tempDict['BOffset'] = self.twColors.item( k, 11).text()
            tempDict['interval'] = self.twColors.item( k, 12).text()

            if self.twColors.cellWidget( k, 13).isChecked():
                tempDict['flashSwitch'] = 1
            else:
                tempDict['flashSwitch'] = 0

            if self.twColors.cellWidget( k, 14).isChecked():
                tempDict['soundSwitch'] = 1
            else:
                tempDict['soundSwitch'] = 0
            # 没有第15列的操作  第15列是选择声音按钮
            tempDict['soundDir'] = self.twColors.item( k, 16).text()
            tempDict['note'] = self.twColors.item( k, 17).text()

            tempColorsDict[i] = tempDict
            i += 1

        # 把tempColorsDict中保存的操作，复制到rowMark指定的位置
        for k in range( 0, len(tempColorsDict)):
            self.twColors.insertRow( rowMark)
            # 第一列是判断当前行是否检测的开关
            tempCheckBoxSwitch = QCheckBox()
            self.twColors.setCellWidget( rowMark, 0, tempCheckBoxSwitch)
            if tempColorsDict[k]['switch'] == 1:
                tempCheckBoxSwitch.setChecked( True)
            else:
                tempCheckBoxSwitch.setChecked( False)
            #第二列是当前检测点， 默认是 0,0    表示还没有进行过检测
            self.twColors.setItem( rowMark, 1, QTableWidgetItem( '0,0'))
            # LX,LY,RX,RY,R,ROffset,G,GOffset,B,BOffset,interval
            self.twColors.setItem( rowMark, 2, QTableWidgetItem( tempColorsDict[k]['LX']))
            self.twColors.setItem( rowMark, 3, QTableWidgetItem( tempColorsDict[k]['LY']))
            self.twColors.setItem( rowMark, 4, QTableWidgetItem( tempColorsDict[k]['RX']))
            self.twColors.setItem( rowMark, 5, QTableWidgetItem( tempColorsDict[k]['RY']))
            self.twColors.setItem( rowMark, 6, QTableWidgetItem( tempColorsDict[k]['R']))
            self.twColors.setItem( rowMark, 7, QTableWidgetItem( tempColorsDict[k]['ROffset']))
            self.twColors.setItem( rowMark, 8, QTableWidgetItem( tempColorsDict[k]['G']))
            self.twColors.setItem( rowMark, 9, QTableWidgetItem( tempColorsDict[k]['GOffset']))
            self.twColors.setItem( rowMark, 10, QTableWidgetItem( tempColorsDict[k]['B']))
            self.twColors.setItem( rowMark, 11, QTableWidgetItem( tempColorsDict[k]['BOffset']))
            self.twColors.setItem( rowMark, 12, QTableWidgetItem( tempColorsDict[k]['interval']))
            # 闪屏开关
            tempCheckBoxFlash = QCheckBox()
            self.twColors.setCellWidget( rowMark, 13, tempCheckBoxFlash)
            if tempColorsDict[k]['flashSwitch'] == 1:
                tempCheckBoxFlash.setChecked( True)
            else:
                tempCheckBoxFlash.setChecked( False)
            # 声音开关
            tempCheckBoxSound = QCheckBox()
            self.twColors.setCellWidget( rowMark, 14, tempCheckBoxSound)
            if tempColorsDict[k]['soundSwitch'] == 1:
                tempCheckBoxSound.setChecked( True)
            else:
                tempCheckBoxSound.setChecked( False)
            # 第15列是选择声音按钮
            tempPushButton = QPushButton()
            tempPushButton.setText('选择')
            self.twColors.setCellWidget( rowMark, 15, tempPushButton)
            tempPushButton.clicked.connect( self.mfSelectSound)
            # 第16列是声音路径
            self.twColors.setItem( rowMark, 16, QTableWidgetItem( tempColorsDict[k]['soundDir']))
            self.twColors.setItem( rowMark, 17, QTableWidgetItem( tempColorsDict[k]['note']))

            rowMark += 1

    # 清空twColors中的所有内容
    def mfClearColors( self):
        for i in range( 0, self.twColors.rowCount()):
            self.twColors.removeRow(0)

    # 清空日志
    def mfClearLog( self):
        self.teLog.clear()

    # 退出程序
    def mfQuit( self):
        self.mfSave()
        app = QApplication.instance()
        app.quit()

    # 通过cbNoEdit checkBox控件 设置twColors检测列表是否可以被修改编辑
    def mfNoEdit( self):
        if self.cbNoEdit.isChecked():
            self.twColors.setEditTriggers(QAbstractItemView.NoEditTriggers)
        else:
            self.twColors.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)

    # 点击自动获取按钮btnAutoGet 在sbAutoGetDelay时间之后， 自动获得鼠标所在坐标的信息，并用QMessageBox和labelAutoGetParam提示
    def mfAutoGet( self):
        time.sleep( self.sbAutoGetDelay.value())
        mouseParamDict = getMouseParam()
        self.labelAutoGetParam.setText( '坐标('+ str(mouseParamDict['x'])+', '+ 
                        str(mouseParamDict['y'])+')  RGB('+str(mouseParamDict['r'])+', '+str(mouseParamDict['g'])+', '
                        +str(mouseParamDict['b'])+')')
        QMessageBox.about( self, '提示', '鼠标当前坐标 ('+ str(mouseParamDict['x'])+', '+ 
                        str(mouseParamDict['y'])+')  RGB('+str(mouseParamDict['r'])+', '+str(mouseParamDict['g'])+', '
                        +str(mouseParamDict['b'])+')')


    # 点击执行,创建线程，开始检测
    def mfStart( self):
        global ISRUN
        ISRUN = 1
        if self.rbSingleThread.isChecked():     # 单线程模式
            inRunThreading = threading.Thread( target= self.mfRun, args=(0, self.twColors.rowCount()))
            inRunThreading.start()
            self.signalCrossThread.emit( 'Display', datetime.datetime.now().strftime('%H:%M:%S') + '  开始检测--单线程')

        elif self.rbMultiThread.isChecked():        # 多线程模式
            if self.twColors.rowCount() <= self.sbThreadCount.value():
                self.signalCrossThread.emit( 'Display', datetime.datetime.now().strftime('%H:%M:%S') + '  开始检测--多线程 每一个检测行创建一个线程')
                for i in range( 0, self.twColors.rowCount()):
                    inRunThreading = threading.Thread( target= self.mfRun, args=(i, i+1))
                    inRunThreading.start()

            elif self.twColors.rowCount() > self.sbThreadCount.value():
                self.signalCrossThread.emit( 'Display', datetime.datetime.now().strftime('%H:%M:%S') + '  开始检测--多线程 每一个线程分配多个检测行')
                rowCount = self.twColors.rowCount()
                threadCount = self.sbThreadCount.value()
                n = rowCount // threadCount     # 为每个线程分配的行数
                a = rowCount % threadCount      # 分配完成后的余数
                # 第一个线程 需要执行的行数    平均分配的行数+余数 n+a
                inRunThreading = threading.Thread( target= self.mfRun, args=(0, n+a))
                inRunThreading.start()
                # 第二个线程及后面的线程平均分配检测行
                for i in range( 0, threadCount-1):
                    inRunThreading = threading.Thread( target= self.mfRun, args=((n+a)+i*n, (n+a)+(i+1)*n))
                    inRunThreading.start()
        else:
            QMessageBox.about( self, '提示', '请选择线程模式threadMode,单线程或多线程')

# 这里测试一段代码  闪屏提醒labelAlert总是卡住，颜色不会变化。 chatgpt说是子线程和主线程共用一个消息队列导致的
# 加入下面的代码不知道有没有用，也不理解实现原理
# 下面是chatgpt关于这个总是的描述：
# 用pyqt写的程序，在创建子线程并运行后，主线程的窗口界面总是卡住不动，需要点击主线程的界面，窗口才会有反应   这是哪里的问题？
# 这种情况一般是因为主线程和子线程同时使用了同一个Qt事件循环（QEventLoop），导致主线程的事件无法及时处理，从而导致主线程界面“卡住”。
# 为了避免这种情况，可以考虑在主线程中使用QTimer或者QApplication.processEvents()定期处理事件循环，这样就能够保证主线程界面不会“卡住”。
# 在这个示例代码中，创建了一个定时器并定期处理事件循环，可以解决主线程界面“卡住”的问题。
# 需要注意的是，定时器的时间间隔不能太短，否则会导致CPU占用率过高，影响程序性能。

        self.timer = QTimer(self)
        self.timer.timeout.connect(lambda: None)
        self.timer.start(100)
# 这三行代码是创建了一个 QTimer 对象并设置了一个超时时间间隔为100毫秒，
# 然后将 QTimer 的 timeout 信号连接到一个匿名函数上，并启动 QTimer。
# 该 QTimer 对象将在启动后每隔 100 毫秒发送一次 timeout 信号。
# 这里的匿名函数用于占位，因为 QTimer 的 timeout 信号必须连接一个槽函数，
# 但是如果我们不需要执行任何操作，只是需要使用 QTimer 来定时发送信号，就可以连接一个占位的空函数。
# 一般情况下，我们会在这个占位的空函数中写一些需要定时执行的代码。例如，我们可以在定时器超时后更新界面、刷新数据等。


    # 点击停止,停止脚本运行
    def mfStop( self):
        global ISRUN
        ISRUN = 0
        self.signalCrossThread.emit( 'Display', datetime.datetime.now().strftime('%H:%M:%S') + '  当前检测行完成后停止')

    # 运行脚本, 核心代码**************************************************
    def mfRun( self, rowStart, rowEnd):
        global ISRUN
        if ISRUN == 0:
            return
        
        hdc = windll.user32.GetDC(None)
        self.signalCrossThread.emit( 'Display', datetime.datetime.now().strftime('%H:%M:%S') + '  当前线程检测 ' 
                                    + str(rowStart + 1) + ' 到 ' + str( rowEnd) + '行' )
        
        for i in range( rowStart, rowEnd):
            if ISRUN == 0:
                return
            tempDict = {}       # 先把当前行的内容存在字典里
            if self.twColors.cellWidget( i, 0).isChecked():
                tempDict['switch'] = 1
            else:
                tempDict['switch'] = 0
            # 当前检测点这一列不用保存，所以没有self.twColors.item( i, 1).text()
            tempDict['LX'] = self.twColors.item( i, 2).text()
            tempDict['LY'] = self.twColors.item( i, 3).text()
            tempDict['RX'] = self.twColors.item( i, 4).text()
            tempDict['RY'] = self.twColors.item( i, 5).text()
            tempDict['R'] = self.twColors.item( i, 6).text()
            tempDict['ROffset'] = self.twColors.item( i, 7).text()
            tempDict['G'] = self.twColors.item( i, 8).text()
            tempDict['GOffset'] = self.twColors.item( i, 9).text()
            tempDict['B'] = self.twColors.item( i, 10).text()
            tempDict['BOffset'] = self.twColors.item( i, 11).text()
            tempDict['interval'] = self.twColors.item( i, 12).text()

            if self.twColors.cellWidget( i, 13).isChecked():
                tempDict['flashSwitch'] = 1
            else:
                tempDict['flashSwitch'] = 0

            if self.twColors.cellWidget( i, 14).isChecked():
                tempDict['soundSwitch'] = 1
            else:
                tempDict['soundSwitch'] = 0
            # 没有第15列的操作  第15列是选择声音按钮
            tempDict['soundDir'] = self.twColors.item( i, 16).text()
            # 第17列是备注
            tempDict['note'] = self.twColors.item( i, 17).text()
            
            if tempDict['switch'] == 0:     # 判断当前行是否要检测
                continue

            breakLoop = False
            for posX in range( int(tempDict['LX']), int(tempDict['RX'])):
                if ISRUN == 0:
                    return
                for posY in range( int(tempDict['LY']), int(tempDict['RY'])):
                    # 向主线程发送消息的字符串， 格式为 行号i-(当前检测点X坐标,当前检测点Y坐标)
                    posStr = str(i) + '-(' + str(posX) + ',' + str(posY) + ')'    # 向主线程发送消息的字符串
                    self.signalCrossThread.emit( 'currentPos', posStr)
                    # 判断颜色是否存在
                    if isColorExist( hdc, posX, posY, int(tempDict['R']), int(tempDict['ROffset']),
                                    int(tempDict['G']), int(tempDict['GOffset']), int(tempDict['B']), int(tempDict['BOffset'])) == True:
                        if tempDict['soundSwitch'] == 1:        # 判断是否播放声音
                            self.signalCrossThread.emit( 'Sound', posStr)
                        
                        if tempDict['flashSwitch'] == 1:        # 判断是否闪屏
                            self.labelAlert.setText( tempDict['note'])
                            for flashCount in range( 0, 15):
                                self.twColors.cellWidget( i, 0).setStyleSheet("background-color: rgb(255, 0, 0)")
                                self.labelAlert.setStyleSheet("background-color: rgb(255, 0, 0)")
                                time.sleep(0.2)
                                self.twColors.cellWidget( i, 0).setStyleSheet("background-color: rgb(0, 0, 255)")
                                self.labelAlert.setStyleSheet("background-color: rgb(0, 0, 255)")
                                time.sleep(0.2)

                        breakLoop = True
                        break
                if breakLoop:
                    break
            time.sleep( int(tempDict['interval']))

        inRunThreading = threading.Thread( target= self.mfRun, args=(rowStart, rowEnd))
        inRunThreading.start()           

                        


#主程序入口
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = ColorScanMulti()
    myWin.show()

    appExit = app.exec_()
    #退出程序之前,保存界面上的设置
    tempDict = { 'csmDir':myWin.labelDir.text(), 'threadCount':myWin.sbThreadCount.value(), 
                'log':myWin.teLog.toHtml(), 'autoGetDelay':myWin.sbAutoGetDelay.value() }
    if myWin.rbMultiThread.isChecked():
        tempDict['threadMode'] = 'multi'
    else:
        tempDict['threadMode'] = 'single'

    if myWin.cbNoEdit.isChecked():
        tempDict['noEdit'] = True
    else:
        tempDict['noEdit'] = False

    saveIniJson = json.dumps( tempDict, indent=4)
    try:
        saveIniFile = open( "./ColorScanMulti.ini", "w",  encoding="utf-8")
        saveIniFile.write( saveIniJson)
        saveIniFile.close()
    except:
        QMessageBox.about( myWin, "提示", "保存配置文件ColorScanMulti.ini失败")

    # 这一句特别重要, 程序是两个线程在运行, 关闭窗口只能结束主线程, 子线程还在运行. 
    # 创建子线程的标志ISRUN 一定要改成0, 子线程在检测ISRUN==0之后,就不再用Timer创建新的线程了
    ISRUN = 0

    sys.exit( appExit)
# sys.exit(app.exec_())  