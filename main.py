from pystray import Icon,Menu,MenuItem
from PIL import Image
#导入必须的库
from PyQt5 import QtWidgets, uic,QtCore
from PyQt5.QtGui import QFont ,QIcon, QPainter, QColor, QBrush, QPen,QPixmap
from PyQt5.QtCore import  Qt, QRect, QPoint, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton,QMainWindow,QMessageBox,QFrame,QHBoxLayout,QTextEdit,QLineEdit
from qfluentwidgets import NavigationItemPosition, FluentWindow, SubtitleLabel, setFont,Flyout,FlyoutAnimationType
from qfluentwidgets import FluentIcon as FIF
import sys
import os
import random
import time
import pygetwindow as gw
import psutil
import easygui
import json
import secrets
import numpy as np
import webbrowser
import base64

# PyInstaller -w(无控制台)打包后 sys.stdout/stderr 为 None,
# 裸 print 会抛 AttributeError 导致程序崩溃;这里重定向到空设备保护
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')
#获取本地运行路径
bin_dir = os.path.join(os.path.dirname(__file__),'bin')
#点名文件目录
#数据文件(名单/配置)固定放在程序所在目录,避免因启动目录不同而漂移
if getattr(sys, 'frozen', False):
    base_dir = os.path.dirname(sys.executable)  # PyInstaller 打包后:exe 所在目录
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))  # 开发时:脚本所在目录

name_file = os.path.join(base_dir, 'name.wow')
name_pro = os.path.join(base_dir, 'name.pro')
config_file = os.path.join(base_dir, 'config.json')
name_list = []
name_password = ''
counted_list = []
#判断启动模式
quiet_boot = False

try:
    for i in sys.argv:
        if i == "-quiet":
            quiet_boot = True
except Exception:
    quiet_boot = False

#检查文件是否存在
def checkfile(path):
    return os.path.exists(path)

#自己配置文件实现(多开检测:比较 PID 所属进程名,避免 PID 复用导致的误报)
try:
    if checkfile(config_file):
        with open(config_file, mode='r', encoding='utf-8') as f:
            config = json.loads(f.read())

        last_pid = int(config['last_pid'])
        duplicate = False
        if last_pid != os.getpid() and psutil.pid_exists(last_pid):
            try:
                #仅当 PID 对应的是同名程序时才认为是重复实例
                duplicate = (psutil.Process(last_pid).name().lower()
                             == psutil.Process().name().lower())
            except psutil.NoSuchProcess:
                duplicate = False

        if duplicate:
            easygui.msgbox('检测到已有开启的实例，PID：' + str(last_pid) + '，\n程序即将退出', '随机点名-不能重复启动实例', '退出')
            os._exit(0)

        config['last_pid'] = os.getpid()
        with open(config_file, mode='w', encoding='utf-8') as f:
            f.write(json.dumps(config))
    else:
        with open(config_file, mode='w', encoding='utf-8') as f:
            f.write(json.dumps({'last_pid': os.getpid()}))
except Exception as q:
    print(q)
#基础变量
press_time = 0.0 #悬浮球按下时间戳(用于区分点击与拖动)


# 创建高质量随机数生成器实例
rng = np.random.Generator(np.random.SFC64(seed=secrets.randbelow(2**64)))

def advanced_shuffle(items):
    """使用现代算法的高质量 shuffle"""
    arr = np.array(items)
    rng.shuffle(arr)
    return arr.tolist()
class FloatingBall(QMainWindow):  #浮动球


    def __init__(self):


        super().__init__()


        # 设置窗口的属性


        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.SubWindow)


        self.setAttribute(Qt.WA_TranslucentBackground, True)

        self.setFixedSize(50, 50) #设置大小

        # 设置窗口的初始大小

        screen_geometry = QApplication.primaryScreen().geometry() #获取屏幕分辨率
        self.setGeometry(screen_geometry.left() + 30, screen_geometry.bottom() - self.height() - 55, self.width(),
                         self.height())

        # 初始化鼠标按下的位置


        self.mouse_press_position = None


        self.corner_radius = 12   # 圆角半径

        # 加载图标
        self.icon_pixmap = QPixmap(bin_dir + '/icon_mini.png').scaled(self.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)


        # 设置悬浮球的样式（背景色通过paintEvent设置）
        self.setStyleSheet("border-radius: {0}px;".format(self.corner_radius))

        # 添加关闭按钮


        #self.close_button = QPushButton('X', self)


        #self.close_button.setGeometry(75, 10, 20, 20)


        #self.close_button.clicked.connect(self.close)


    def paintEvent(self, event):


        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 设置背景色画笔和画刷
        bg_brush = QBrush(QColor(255, 255, 255, 220))  # 白色背景，带透明度
        painter.setBrush(bg_brush)

        # 绘制圆角方形背景
        rect = QRect(0, 0, self.width(), self.height())
        painter.drawRoundedRect(rect, self.corner_radius, self.corner_radius)

        # 绘制图标，保持图标中心对齐
        icon_rect = QRect(QPoint(0, 0), self.icon_pixmap.size())
        icon_rect.moveCenter(self.rect().center())
        painter.drawPixmap(icon_rect, self.icon_pixmap)


    def mousePressEvent(self, event):
        global press_time

        if event.button() == Qt.LeftButton:
            press_time = time.monotonic()
            self.mouse_press_position = event.globalPos() - self.frameGeometry().topLeft()


    def mouseMoveEvent(self, event):


        if (event.buttons() & Qt.LeftButton) and self.mouse_press_position:


            self.move(event.globalPos() - self.mouse_press_position)


    def mouseReleaseEvent(self, event):
        global press_time
        #按下时间不超过 0.5 秒视为"点击",否则视为"拖动",拖动后不弹窗
        if time.monotonic() - press_time <= 0.5:
            print("开启窗口")
            #！！！！开启窗口
            mWindow.hide()
            mWindow.show()
            mWindow.showNormal()

        self.mouse_press_position = None

class SEEWO_Tools(): #SEEWO 用途相关工具（托盘工具、PPT检测[已弃用]）
    def __init__(self):
        self.FLOAT_KEEPOPEN = False
    def showIcon(self):
        self.icon = Icon("my_icon",title="随机点名")

        menu = Menu(
            MenuItem('显示主界面(可能会有问题)',lambda:self.showWindow()),
        MenuItem('退出程序',lambda:self.exitProgram())
        )

        self.icon.icon = Image.open(os.path.join(bin_dir, 'icon.ico'))

        self.icon.menu = menu
        # run_detached 在独立线程运行托盘消息循环,避免阻塞主线程的 Qt 事件循环(Windows 支持)
        try:
            self.icon.run_detached()
        except Exception as e:
            # 个别平台(如 macOS 开发环境)不支持 run_detached,托盘不可用但不影响主程序
            print('托盘图标启动失败(不影响主程序):', e)
    def showMessage(self,messages,titles='随机点名'):
        self.icon.notify(title=titles,message=messages)
    def exitProgram(self):
        global app
        app.quit()
        try:
            self.icon.stop()
        except Exception:
            pass
        os._exit(0)
    def showWindow(self):
        mWindow.hide()
        mWindow.show()
        mWindow.showNormal()
    
#新版名单逻辑
class NewList():
    def file_load(self): #用来读取本地的名单文件，返回类型json
        
        try: #尝试打开名单文件
            with open(name_pro,mode='rb') as f:
                file_data = f.read() #读入文件 
                file_decode = base64.b64decode(file_data).decode('utf-8') #b64 解码 然后用utf-8解码
                #尝试读取数量，顺便确保是存在的
                print('解码后文件',file_decode)
                file_sss = json.loads(file_decode)
                print(file_sss['num'])

                return(json.loads(file_decode)) #返回文件内容(json转成字典)
            
        except Exception as q:#打开失败
            print('打开失败',q)
            #先备份损坏文件，避免名单被静默清空
            try:
                if os.path.exists(name_pro):
                    os.rename(name_pro, name_pro + '.bak')
            except Exception:
                pass
            with open(name_pro,mode='wb') as f:
                #定义空文件内容
                name_content = {
                    'num':0,
                    'password_exist':False
                                }
                file_data = json.dumps(name_content).encode('utf-8') #转换成json并准备bas64编码的bytes
                file_encode = base64.b64encode(file_data)  #bas64 编码
                f.write(file_encode) #写文件
                f.close() #关闭文件
                return(name_content) #返回文件内容(json转成字典)

    def load(self): #用来载入读入的文件
        global name_content
        name_content = self.file_load()
        print('读取到的文件内容',name_content)
        global name_list
        global name_password
        name_list = [] #清空名单列表
        if name_content['num'] > 0:
            for i in range(0,name_content['num']):
                key = 'student' + str(i)
                if key in name_content: #防御:键缺失时跳过,避免 KeyError 崩溃
                    name_list.append(name_content[key])
        if name_content['password_exist']:
            name_password = name_content['password']
        else:
            name_password = ''
        print('读取到的密码',name_password)
        global counted_list
        counted_list = [] #清空点过列表
        print('载入的名单',name_list)
    
    def save(self,name:list):
        
        name_content = self.file_load()
        #清理旧的 student 键,避免残留脏数据
        for key in [k for k in name_content.keys() if k.startswith('student')]:
            del name_content[key]
        name_content['num'] = len(name)
        print('新名字列表',name)
        count = 0
        for i in name:
            name_content['student'+str(count)] = i
            count = count + 1
        with open(name_pro,mode='wb') as f:
                f.write(base64.b64encode(json.dumps(name_content).encode('utf-8')))
                f.close()
        file_manager.load()

        
    def passwd(self,password:str,new_password:str):
        if name_password == password:
            file_content = self.file_load()
            if new_password != '':
                file_content['password_exist'] = True
                file_content['password'] = new_password
                print(file_content)
            elif new_password == '':
                file_content['password'] = ''
                file_content['password_exist'] = False
        
            with open(name_pro,mode='wb') as f:
                    wdnmd = json.dumps(file_content)
                    print(wdnmd)
                    f.write(base64.b64encode(wdnmd.encode('utf-8')))
                    f.close()
            return(True)
        else:
            return(False)


class ChooseWorker(QThread):
    """后台抽选线程:在非 GUI 线程执行抽取逻辑,通过信号把结果送回主线程,
    避免跨线程直接操作 Qt 控件(原实现直接跨线程 setText,会导致崩溃/未定义行为)"""
    name_ready = pyqtSignal(str)  #滚动显示的名字
    done = pyqtSignal()           #抽选完成

    def run(self):
        try:
            ok = get_list_new()
            for i in ok:
                self.name_ready.emit(i.replace('\n', '').replace('\r', ''))
                time.sleep(0.1)
        finally:
            #无论是否异常都通知主线程恢复按钮,避免按钮永久禁用
            self.done.emit()


def get_list_new():
    """取得新的列表:洗牌后滚动展示候选,最后落点即被抽中者(标记为已点)。
    无论名单人数多少,都会真实消耗 1 人,保证"不重复点名"始终生效"""
    global name_list
    global counted_list

    if not name_list:
        return []

    shuffled = advanced_shuffle(name_list)  #洗牌(新列表,不改动 name_list)
    need_count_num = random.randint(13, 17)

    #候选滚动序列:人数足够取前 N 个,不足则全部滚动展示
    if len(shuffled) >= need_count_num:
        display = shuffled[:need_count_num]
    else:
        display = shuffled

    #从滚动序列中随机选一个作为最终落点(追加到末尾,滚动最后停在这里)
    pianyi = random.randint(0, len(display) - 1)
    chosen = display[pianyi]
    display.append(chosen)

    #标记已点:从剩余名单中移除一个匹配项(用索引删除,语义明确)
    try:
        idx = name_list.index(chosen)
        name_list.pop(idx)
        counted_list.append(chosen)
    except ValueError:
        pass

    print('抽到的', chosen)
    print('抽过的', counted_list)
    print('没抽的', name_list)
    return display


class MainWindow(QMainWindow): #主功能实现窗口
    def __init__(self):
        super(MainWindow, self).__init__()
        #加载.ui 文件
        uic.loadUi(bin_dir + '/main.ui',self)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint) #禁用最大化按钮
        #设置/固定宽度
        self.setFixedSize(self.width(), self.height()) #固定宽度和高度
        self.setWindowIcon(QIcon(bin_dir + '/icon.ico'))
        #self.setStyleSheet('background-color:white')

        #查找按钮对象
        self.Start_button = self.findChild(QPushButton,'Start_button')
        self.Open_File_Button = self.findChild(QPushButton,'OpenFile_button')
        self.Reset_button = self.findChild(QPushButton,"Reset_Button")

        self.name_label = self.findChild(QLabel,'Name_label')
        self.status_label = self.findChild(QLabel,'status_label')

        #连接信号和槽
        self.Start_button.clicked.connect(self.StartButton_do)
        self.Open_File_Button.clicked.connect(self.Open_File_button_do)
        self.Reset_button.clicked.connect(self.reset_button_do)

        ######   设置按钮字体
        self.name_label.setAlignment(Qt.AlignCenter)
        font = self.name_label.font()  # 获取当前字体
        font.setPointSize(72)  # 设置新的字体大小
        self.name_label.setFont(font)  # 应用新的字体

        font = self.name_label.font()  # 获取当前字体
        font.setPointSize(24)
        self.Start_button.setFont(font)

        self.name_label.setText('未选定')
        
        #设置FluentWindow 独立对象名
        self.setObjectName('home')
        
        #######
        
        self.refresh_status()
        
        


    def StartButton_do(self):
        print('开始按钮被按下')
        if len(name_list) == 0:
            self.name_label.setText("文件为空")
        else:
            #禁用按钮，避免抽选期间重复点击或重置造成的并发修改
            self.Start_button.setEnabled(False)
            self.Start_button.setText('正在抽取……')
            self.Reset_button.setEnabled(False)

            self.worker = ChooseWorker()
            self.worker.name_ready.connect(self.name_label.setText)
            self.worker.done.connect(self.on_choose_done)
            self.worker.start()

    def Open_File_button_do(self):
        print("打开文件按钮被按下")
        Flyout.create(icon=FIF.INFO,title='想要修改名单文件？',content='版本已更新，请转到左侧设置页面修改名单',target=self.Open_File_Button,parent=self,isClosable=True,aniType=FlyoutAnimationType.PULL_UP)
    def reset_button_do(self):
        print("重置按钮被按下")
        reset_App()
        self.refresh_status()
        self.name_label.setText('未选定')
    
    def on_choose_done(self):
        #抽选结束,恢复按钮(通过信号在主线程执行)
        self.Start_button.setEnabled(True)
        self.Start_button.setText('开始')
        self.Reset_button.setEnabled(True)
        print('开始刷新显示')
        self.refresh_status()
    def get_old(self): #旧的函数，用不到，但是还是想留着
        global name_list
        global counted_list
        if not name_list:
            return
        t = 0 
        while t <= 15:
            ok = random.randint(0,len(name_list)-1)
            self.name_label.setText(name_list[ok].replace('\n','').replace('\r',''))
            t = t + 1
            time.sleep(0.1)
        
        #移动最后的学生到点过列表:
        counted_list.append(name_list[ok])
        del name_list[ok]
    def closeEvent(self, a0): #处理关闭信号
        print('退出按钮被按下')
        #保存信息
        mWindow.hide()
        SEEWO_Tool.showMessage('窗口已最小化到托盘')
        a0.ignore()
    
    
    def refresh_status(self):
        global name_list
        sb1 = '总人数：'
        sb2 = ',已点人数:'   
        sb3 = ',剩余人数：'
        Total_SB =str(len(name_list) + len(counted_list))
        self.status_label.setText(sb1 + Total_SB + sb2 + str(len(counted_list)) + sb3 + str(len(name_list)))
        #random.shuffle(name_list)
        print('刷新结束')
class about(QWidget):
    def __init__(self):
        super().__init__()
        uic.loadUi(bin_dir + '/about.ui',self) 
        self.setObjectName('about')
class settings(QWidget):
    def loadtext(self):
        for i in name_list:
            text = self.name_text.toPlainText() + i + '\n'
            self.name_text.setPlainText(text)
    def Locked(self):
        self.lock_status_label.setText('已锁定')
        self.lock_status_label.setStyleSheet('color:red;font:bold 14px')
        self.name_text.setReadOnly(True)
        self.password_lineedit.setEnabled(True)
        self.unlock_button.setText("解锁")
        self.islock = True
        self.saveButton.setEnabled(False)
    def Unlocked(self):
        self.lock_status_label.setText('已解锁')
        self.lock_status_label.setStyleSheet('color:green;font:bold 14px')
        self.name_text.setReadOnly(False)
        self.password_lineedit.setEnabled(False)
        self.unlock_button.setText("锁定")
        self.islock = False
        self.saveButton.setEnabled(True)
    def __init__(self):
        super().__init__()
        uic.loadUi(bin_dir + '/settings.ui',self) 
        self.setObjectName('settings')
        
        #查找对象
        self.name_text = self.findChild(QTextEdit,'contentEdit')
        self.unlock_button = self.findChild(QPushButton,'unlock_button')
        self.password_lineedit = self.findChild(QLineEdit,'lineEdit')
        self.change_password_button = self.findChild(QPushButton,'change_password_button')
        self.lock_status_label = self.findChild(QLabel,'Lock_status')
        self.saveButton = self.findChild(QPushButton,'SaveButton')
        #连接信号和槽
        self.unlock_button.clicked.connect(self.unlock_button_do)
        self.change_password_button.clicked.connect(self.change_password_button_do)
        self.saveButton.clicked.connect(self.save_button_do)

        #加载文本
        self.loadtext()


        self.islock = True
        #设置显示逻辑
        file_content = file_manager.file_load()
        if file_content['password_exist']:
            self.Locked()
        else:
            self.Unlocked()
    
    def save_button_do(self):
        print('保存按钮点击')
        print('密码内容',self.password_lineedit.text())
        text = self.name_text.toPlainText()
        wow = list(filter(None,text.split('\n')))
        
        file_manager.save(wow)
    def unlock_button_do(self):
        print('解锁/锁定 按钮点击',self.islock)
        if self.islock:
            if self.password_lineedit.text() == name_password:
                self.Unlocked()
            else:
                Flyout.create(icon=FIF.CAFE,title='你干嘛~',content='密码错误',target=self.unlock_button,parent=self,isClosable=True,aniType=FlyoutAnimationType.DROP_DOWN)
        else:
            self.Locked()
    def change_password_button_do(self):
        print('修改密码按钮点击')
        oldpassword = easygui.passwordbox('请输入旧密码（没有就留空）')
        if oldpassword != name_password:
            easygui.msgbox('旧密码错误')
            return
        new1 = easygui.passwordbox('请输入新密码')
        new2 = easygui.passwordbox('请再次输入')
        if new1 != new2:
            easygui.msgbox('两次密码不一致，已取消修改')
            return
        if new1 == '':
            #留空表示清除密码保护,需要二次确认,避免误操作
            if not easygui.ccbox('新密码为空，将清除密码保护，是否继续？'):
                return
        if file_manager.passwd(oldpassword,new1):
            easygui.msgbox('修改成功')
        else:
            easygui.msgbox('修改失败')
        file_manager.load()


class WelcomeWindow(FluentWindow): #多合一窗口
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowMaximizeButtonHint) #禁用最大化按钮
        #设置/固定宽度
        #self.setFixedSize(self.width(), self.height()) #固定宽度和高度
        self.setFixedSize(820, 530) 
        self.setWindowIcon(QIcon(bin_dir + '/icon.ico'))
        
        
        #添加 MainWindow 作为子窗口
        self.homeInterface = MainWindow()
        self.addSubInterface(self.homeInterface,icon=FIF.HOME,text='随机点名',position=NavigationItemPosition.TOP)

        #添加 Settings 作为子窗口
        self.setInterface = settings()
        self.addSubInterface(self.setInterface,icon=FIF.SETTING,text='设置',position=NavigationItemPosition.BOTTOM)
    
        #添加 about 作为子窗口
        self.aboutInterface = about()
        self.addSubInterface(self.aboutInterface,icon=FIF.INFO,text='关于',position=NavigationItemPosition.BOTTOM)
                
        
        #标题栏定制
        self.titleBar.maxBtn.hide() #禁用最大化按钮
        self.titleBar.setDoubleClickEnabled(False) #禁用双击最大化
        self.setWindowTitle('随机点名')
        
        #查找按钮对象
        self.author_button = self.findChild(QPushButton,'pushButton')
        self.software_info_button = self.findChild(QPushButton,'pushButton_2')

        #连接按钮
        self.author_button.clicked.connect(self.author_button_do)
        self.software_info_button.clicked.connect(self.software_info_button_do)


        
    def closeEvent(self, a0): #处理关闭信号
        print('退出按钮被按下')
        #保存信息
        mWindow.hide()
        SEEWO_Tool.showMessage('窗口已最小化到托盘')
        a0.ignore()
    def author_button_do(self):
        webbrowser.open('https://github.com/Xiaoxiaoyu1321')
    def software_info_button_do(self):
        webbrowser.open('https://github.com/Xiaoxiaoyu1321/Random-roll-call')

def reset_App1(): #旧版文件读取
    global name_list
    global counted_list
    #读入点名文件
    with open(name_file,mode='r',encoding='utf-8') as f:
        file_content = f.readlines()
        f.close()
    #分析点名文件
    name_list = [] #清空点名文件
    counted_list = [] #清空点过列表
    print(len(name_list))
    print(len(file_content))
    try:
        for i in file_content:
            if i.startswith('#'):
                print('检测到注释',i)
                

            else:
                name_list.append(i)
                #print(i)
                
    except Exception as q:
        print(q)
        easygui.msgbox('载入文件时遇到错误,'+q,title='Error!!!')

#新版文件读取
def reset_App():
    file_manager.load()



if __name__ == "__main__":
    file_manager = NewList()
    #检查所需文件是否存在
    #    if not checkfile(name_file):
    #        print('找不到点名文件，现在创建一个')
    #        with open(name_file,mode='w+',encoding='utf-8') as f:
    #            default_text = '#这是一个点名文件，它采用txt文件形式保存。  \n#像这样子，以“#” 开头的文本不会被当做姓名处理，如果您希望添加注释，也可以在注释的文本前添加“#” \n#请直接将姓名每行一个粘贴到下面的空白区域，请确保没有多余的空行!!!!!'
    #            f.write(default_text)
    #            f.close()
        
    #调用Reset_APP 方法重置应用程序
    reset_App()
    
   
    
    
    
    


    #PyQT 操作
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)
    Float_Ball = FloatingBall()
    Float_Ball.show()

    #mWindow = MainWindow()
    mWindow = WelcomeWindow()
    mWindow.show()

    #加载托盘
    SEEWO_Tool=SEEWO_Tools()
    SEEWO_Tool.showIcon()

    #判断是否Quiet Boot
    if quiet_boot:
        mWindow.hide()
    
    



    sys.exit(app.exec_())
    pass
