
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QWidget
import sys
from PyQt5.QtGui import QPixmap, QIcon
from themes import *



class App(QWidget):
	def __init__(self):
		super().__init__()
		self.setWindowTitle("KRYPTT")
		self.setWindowIcon(QIcon("logo.png"))
		self.resize(1324, 771)
		self.setStyleSheet("background:rgb(30,30,20);")
		self.title = QtWidgets.QLabel(self)
		self.title.setGeometry(QtCore.QRect(80, 30, 581, 111))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(48)
		font.setBold(True)
		font.setWeight(75)
		self.title.setFont(font)
		self.title.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.title.setObjectName("title")
		self.bio = QtWidgets.QLabel(self)
		self.bio.setGeometry(QtCore.QRect(80, 130, 1231, 71))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(14)
		font.setBold(True)
		font.setWeight(75)
		self.bio.setFont(font)
		self.bio.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.bio.setWordWrap(True)
		self.bio.setObjectName("bio")
		self.input_topic = QtWidgets.QLineEdit(self)
		self.input_topic.setGeometry(QtCore.QRect(70, 230, 411, 71))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(16)
		font.setBold(True)
		font.setWeight(75)
		self.input_topic.setFont(font)
		self.input_topic.setStyleSheet("background:rgba(255, 255, 255, 0.7);;\n"
	"border:0px;\n"
	"border-radius:15px;\n"
	"color:black;\n"
	"padding-left:10px;\n"
	"")
		self.input_topic.setObjectName("input_topic")
		self.input_name = QtWidgets.QLineEdit(self)
		self.input_name.setGeometry(QtCore.QRect(70, 330, 411, 71))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(16)
		font.setBold(True)
		font.setWeight(75)
		self.input_name.setFont(font)
		self.input_name.setStyleSheet("background:rgba(255, 255, 255, 0.7);;\n"
	"border:0px;\n"
	"border-radius:15px;\n"
	"color:black;\n"
	"padding-left:10px;\n"
	"")
		self.input_name.setObjectName("input_name")
		self.input_slides = QtWidgets.QSlider(self)
		self.input_slides.setGeometry(QtCore.QRect(60, 500, 511, 31))
		self.input_slides.setStyleSheet("QSlider::handle:horizontal {background-color: white; border-radius: 5px; height: \n"
	"50px; width: 50px;}")
		self.input_slides.setMaximum(9)
		self.input_slides.setPageStep(11)
		self.input_slides.setOrientation(QtCore.Qt.Horizontal)
		self.input_slides.setTickPosition(QtWidgets.QSlider.TicksBelow)
		self.input_slides.setObjectName("input_slides")
		self.head1 = QtWidgets.QLabel(self)
		self.head1.setGeometry(QtCore.QRect(60, 430, 531, 71))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(14)
		font.setBold(True)
		font.setWeight(75)
		self.head1.setFont(font)
		self.head1.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.head1.setWordWrap(True)
		self.head1.setObjectName("head1")
		self.submit = QtWidgets.QPushButton(self)
		self.submit.setGeometry(QtCore.QRect(770, 680, 481, 61))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(22)
		font.setBold(True)
		font.setWeight(75)
		self.submit.setFont(font)
		self.submit.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
		self.submit.setMouseTracking(True)
		self.submit.setStyleSheet("QPushButton{background:rgba(255, 255, 255, 0.7);;\n"
	"border:0px;\n"
	"border-radius:15px;\n"
	"color:black;\n"
	"padding-left:10px;\n"
	"}\n"
	"QPushButton::hover{background:white;}")
		self.submit.setObjectName("submit")
		self.input_theme = QtWidgets.QSlider(self)
		self.input_theme.setGeometry(QtCore.QRect(60, 660, 511, 31))
		self.input_theme.setStyleSheet("QSlider::handle:horizontal {background-color: white; border-radius: 5px; height: \n"
	"50px; width: 50px;}")
		self.input_theme.setMaximum(3)
		self.input_theme.setPageStep(11)
		self.input_theme.setOrientation(QtCore.Qt.Horizontal)
		self.input_theme.setTickPosition(QtWidgets.QSlider.TicksBelow)
		self.input_theme.setObjectName("input_theme")
		self.head2 = QtWidgets.QLabel(self)
		self.head2.setGeometry(QtCore.QRect(60, 590, 531, 71))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(14)
		font.setBold(True)
		font.setWeight(75)
		self.head2.setFont(font)
		self.head2.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.head2.setWordWrap(True)
		self.head2.setObjectName("head2")
		self.preview_pane = QtWidgets.QLabel(self)
		self.preview_pane.setGeometry(QtCore.QRect(780, 270, 451, 280))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(48)
		font.setBold(True)
		font.setWeight(75)
		self.preview_pane.setFont(font)
		self.preview_pane.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.preview_pane.setScaledContents(True)
		self.preview_pane.setAlignment(QtCore.Qt.AlignCenter)
		self.preview_pane.setObjectName("preview_pane")
		self.logo_instagram = QtWidgets.QLabel(self)
		self.logo_instagram.setGeometry(QtCore.QRect(1230, 30, 51, 51))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(48)
		font.setBold(True)
		font.setWeight(75)
		self.logo_instagram.setFont(font)
		self.logo_instagram.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.logo_instagram.setText("")
		self.logo_instagram.setScaledContents(True)
		self.logo_instagram.setAlignment(QtCore.Qt.AlignCenter)
		self.logo_instagram.setObjectName("logo_instagram")
		self.logo_twitter = QtWidgets.QLabel(self)
		self.logo_twitter.setGeometry(QtCore.QRect(1155, 35, 51, 43))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(48)
		font.setBold(True)
		font.setWeight(75)
		self.logo_twitter.setFont(font)
		self.logo_twitter.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.logo_twitter.setText("")
		self.logo_twitter.setScaledContents(True)
		self.logo_twitter.setAlignment(QtCore.Qt.AlignCenter)
		self.logo_twitter.setObjectName("logo_twitter")
		self.logo_github = QtWidgets.QLabel(self)
		self.logo_github.setGeometry(QtCore.QRect(1068, 30, 51, 51))
		font = QtGui.QFont()
		font.setFamily("Gotham")
		font.setPointSize(48)
		font.setBold(True)
		font.setWeight(75)
		self.logo_github.setFont(font)
		self.logo_github.setStyleSheet("background:transparent;\n"
	"color:white;")
		self.logo_github.setText("")
		self.logo_github.setScaledContents(True)
		self.logo_github.setAlignment(QtCore.Qt.AlignCenter)
		self.logo_github.setObjectName("logo_github")
		_translate = QtCore.QCoreApplication.translate
		self.setWindowTitle(_translate("self", "KRYPPT - PPT MAKER"))
		self.title.setText(_translate("self", "KRYPPT"))
		self.bio.setText(_translate("self", "<html><head/><body><p><span style=\" font-weight:400;\">Welcome to Project KRYPPT, It allows you to create PPT for your school project in less than 5 seconds. Just enter the topic &amp; your name, select the number of informational slides you want .</span></p></body></html>"))
		self.input_topic.setPlaceholderText(_translate("self", "Enter Topic"))
		self.input_name.setPlaceholderText(_translate("self", "Enter Your Name"))
		self.head1.setText(_translate("self", "<html><head/><body><p><span style=\" color:#ffffff;\">Number of Slides (excluding the title, index, acknowledgment, bibliography &amp; Thankyou)</span></p></body></html>"))
		self.submit.setText(_translate("self", "Create My PPT"))
		self.head2.setText(_translate("self", "<html><head/><body><p><span style=\" color:#ffffff;\">Select The PPT Design</span></p></body></html>"))
		self.preview_pane.setText(_translate("self", "Preview"))
		self.submit.clicked.connect(lambda:create_ppt_button())
		self.input_theme.valueChanged.connect(lambda: change_preview())
		def create_ppt_button():
			topic = self.input_topic.text()
			name = self.input_name.text()
			if topic!="" and name!="":
				theme = self.input_theme.value()
				slides = self.input_slides.value() + 1
				if int(theme) == 0:
					create_ppt_2(topic, name, int(slides))
				elif int(theme) == 1:
					create_ppt_3(topic, name, int(slides))
				elif int(theme) == 2:
					create_ppt_4(topic, name, int(slides))
				elif int(theme) == 3:
					create_ppt_5(topic, name, int(slides))
		self.preview_pane.setPixmap(QPixmap("preview/"+str(1)+".png"))
		self.logo_github.setPixmap(QPixmap("logo/github.png"))
		self.logo_twitter.setPixmap(QPixmap("logo/twitter.png"))
		self.logo_instagram.setPixmap(QPixmap("logo/instagram.png"))
		def change_preview():
			theme = self.input_theme.value()
			self.preview_pane.setPixmap(QPixmap("preview/"+str(theme+1)+".png"))

if __name__=="__main__":
	app = QApplication(sys.argv)
	a = App()
	a.show()
	sys.exit(app.exec_())