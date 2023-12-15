import os
import shutil
import sys
import pandas as pd
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import *
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Font, Border, Side
from openpyxl.styles import numbers


class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('My Auto Application')
        self.setICon()
        self.setGridLayout()
        self.setWindowCenter()
        self.show()

    def setICon(self):
        self.setWindowIcon(QIcon('./img/Geostory_Logo.png'))
        self.setGeometry(600, 300, 500, 300)
        self.show()

    def setGridLayout(self):
        grid = QGridLayout()
        self.setLayout(grid)

        # 객체 생성
        self.filePath = QLineEdit('E:/geostory/2023/타 부서 업무협조/해양사업부/데이터/총괄표/(신양식) 장애물 관리대장(무안공항)_231214.xlsx')
        self.folderPath = QLineEdit('E:/geostory/2023/타 부서 업무협조/해양사업부/데이터/FolderTree_v2')
        self.formPath = QLineEdit('E:/geostory/2023/타 부서 업무협조/해양사업부/데이터/장애물 관리대장 양식')

        self.fileSelectBtn = QPushButton('열기')
        self.folderSelectBtn = QPushButton('열기')
        self.formSelectBtn = QPushButton('열기')
        self.accessBtn = QPushButton('적용')

        # 객체 위치 지정
        grid.addWidget(QLabel('관리대장 : '), 1, 1)
        grid.addWidget(QLabel('이미지 경로 : '), 2, 1)
        grid.addWidget(QLabel('관리대장 양식 : '), 3, 1)

        grid.addWidget(self.filePath, 1, 2)
        grid.addWidget(self.folderPath, 2, 2)
        grid.addWidget(self.formPath, 3, 2)

        grid.addWidget(self.fileSelectBtn, 1, 3)
        grid.addWidget(self.folderSelectBtn, 2, 3)
        grid.addWidget(self.formSelectBtn, 3, 3)

        grid.addWidget(self.accessBtn, 4, 2)

        self.fileSelectBtn.clicked.connect(self.selectFile)
        self.folderSelectBtn.clicked.connect(self.selectFolder)
        self.formSelectBtn.clicked.connect(self.selectForm)
        self.accessBtn.clicked.connect(self.accessLogic)

    def setWindowCenter(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def selectFile(self):
        fName = QFileDialog.getOpenFileName(self, '파일 선택', '', 'excel file(*.xls *.xlsx')
        self.filePath.setText(fName[0])

    def selectFolder(self):
        fName = QFileDialog.getExistingDirectory(self, '폴더 선택', '')
        self.folderPath.setText(fName[:])

    def selectForm(self):
        fName = QFileDialog.getExistingDirectory(self, '폴더 선택', '')
        self.formPath.setText(fName[:])

    # 필요한 파일 및 폴더 입력 후 처리 함수
    def accessLogic(self):
        if not self.filePath.text() or not self.folderPath.text() or not self.formPath.text():
            print('파일 또는 폴더를 선택해 주세요.')
        else:
            print('파일, 폴더 입력 완료', self.filePath.text(), self.formPath.text())
            self.readExcel()

    # 관리대장 읽어오는 함수
    def readExcel(self):
        self.pf = pd.read_excel(self.filePath.text(), header=3, usecols='B,C,E,I:L,P:AF,AJ,AL,AN:AP,AZ:BA,BF,BY')
        self.pf.columns = ['순번', '신규연도', '연번', '세부종류', '장애물 용도', '명칭', '건축주', '특례 장애물 구분', '차폐 기준 장애물 및 지정일',
                           '주소1', '주소2',
                           '주소3',
                           '주소4', '도로명주소', '위치구역', '위도', '경도', 'X축', 'Y축', '지반높이', '건물/시설물/수목 높이', '전체높이',
                           '제한표면',
                           '제한표면 침범높이',
                           '협의높이', '위반여부', '기관명', '연락처', '관리번호', '건축허가일', '준공승인일', '장애물 등재일', '좌표/높이 결정방법']
        print(self.pf)
        self.makeFolder()
        self.copyExcel()

    # 결과 저장 폴더 만드는 함수
    def makeFolder(self):
        print('MyApp - makeFolder()')
        self.path = os.getcwd()
        self.savePath = self.path + "/result"
        try:
            if not os.path.exists(self.savePath):
                os.mkdir(self.savePath)
        except:
            print('Error : Creating directory.' + self.savePath)

    def copyExcel(self):
        print('MyApp - copyExcel()')
        print(self.pf.columns, len(self.pf))
        try:
            for i in range(len(self.pf)):
                imgPath = self.folderPath.text() + "/" + str(self.pf['연번'][i])
                print('imgPath : ',imgPath)
                if self.pf['세부종류'][i] == '나무':
                    fileName = '/' + str(self.pf['연번'][i]) + ".xlsx"
                    save = self.savePath + fileName
                    form = self.formPath.text() + "/장애물 관리대장 및 상세표 양식(나무)_v2.xlsx"
                    shutil.copy(form, save)
                    self.inputDataToExcel(savePath=save, data=self.pf.loc[i], imgPath=imgPath)
                elif self.pf['세부종류'][i] == '산':
                    fileName = '/' + str(self.pf['연번'][i]) + ".xlsx"
                    save = self.savePath + fileName
                    form = self.formPath.text() + "/장애물 관리대장 및 상세표 양식(산)_v2.xlsx"
                    shutil.copy(form, save)
                    self.inputDataToExcel(savePath=save, data=self.pf.loc[i], imgPath=imgPath)
                elif self.pf['세부종류'][i] == '건물':
                    fileName = '/' + str(self.pf['연번'][i]) + ".xlsx"
                    save = self.savePath + fileName
                    form = self.formPath.text() + "/장애물 관리대장 및 상세표 양식(건물)_v2.xlsx"
                    shutil.copy(form, save)
                    self.inputDataToExcel(savePath=save, data=self.pf.loc[i], imgPath=imgPath)
                else:
                    fileName = '/' + str(self.pf['연번'][i]) + ".xlsx"
                    save = self.savePath + fileName
                    form = self.formPath.text() + "/장애물 관리대장 및 상세표 양식(기타)_v2.xlsx"
                    shutil.copy(form, save)
                    self.inputDataToExcel(savePath=save, data=self.pf.loc[i], imgPath=imgPath)
        except Exception as e:
            print('Error : Copy Excel. ' + e)
        finally:
            print('MyApp - copyExcel() done!')

    def inputDataToExcel(self, savePath, data, imgPath):
        try:
            print('inputDataToExcel : ', data['연번'])
            wb = load_workbook(savePath)
            ws1 = wb["연번"]
            ws2 = wb["연번(장애물 관리대장 상세표)"]
            ws1.title = str(data['연번'])
            ws2.title = str(data['연번'])+"(장애물 관리대장 상세표)"
            print("sheet 이름", wb.sheetnames)

            ws1['A5'].value=data['연번']
            ws1['B5'].value=data['명칭']
            ws1['C5'].value =data['주소1'] + " " + data['주소2'] + " " + data['주소3'] + " " + data['주소4']
            ws1['D6'].value =data['도로명주소']
            ws1['F5'].value =data['위치구역']
            ws1['A9'].value =data['위도']
            ws1['B9'].value =data['경도']
            ws1['C9'].value =data['X축']
            ws1['C9'].number_format = '00"˚"00"′"00.00"″"'
            ws1['D9'].value =data['Y축']
            ws1['D9'].number_format = '00"˚"00"′"00.00"″"'
            ws1['E9'].value =data['좌표/높이 결정방법']
            ws1['F9'].value =data['건축허가일']
            ws1['F9'].number_format = "yy/mm/dd"
            ws1['G9'].value =data['준공승인일']
            ws1['G9'].number_format = "yy/mm/dd"
            ws1['A13'].value =data['지반높이']
            ws1['B13'].value =data['건물/시설물/수목 높이']
            ws1['C13'].value =data['전체높이']
            ws1['D13'].value =data['제한표면']
            ws1['E13'].value =data['제한표면 침범높이']
            ws1['F13'].value =data['협의높이']
            ws1['G13'].value =data['위반여부']
            ws1['A18'].value =data['신규연도']
            ws1['A18'].number_format = "yy/mm/dd"
            ws1['B18'].value =data['특례 장애물 구분']
            ws1['C18'].value =data['차폐 기준 장애물 및 지정일']
            ws1['E18'].value =data['장애물 등재일']
            ws1['A23'].value =data['장애물 용도']
            ws1['B23'].value =data['건축주']
            ws1['C23'].value =data['기관명']
            ws1['D23'].value =data['연락처']
            ws1['F23'].value =data['관리번호']

            if data['순번'] != '제거':
                self.typeImage(data['세부종류'], imgPath, data, wb, ws1, ws2, savePath)

            ws1['G20'].border = Border(right=Side(border_style='medium', color="000000"),bottom=Side(border_style='thin', color="000000"),left=Side(border_style='thin', color="000000"),top=Side(border_style='thin', color="000000"))
            ws1['G21'].border = Border(right=Side(border_style='medium', color="000000"),bottom=Side(border_style='thin', color="000000"),left=Side(border_style='thin', color="000000"),top=Side(border_style='thin', color="000000"))
            ws1['G22'].border = Border(right=Side(border_style='medium', color="000000"),bottom=Side(border_style='thin', color="000000"),left=Side(border_style='thin', color="000000"),top=Side(border_style='thin', color="000000"))
            ws1['G23'].border = Border(right=Side(border_style='medium', color="000000"),bottom=Side(border_style='thin', color="000000"),left=Side(border_style='thin', color="000000"),top=Side(border_style='thin', color="000000"))
            ws1['G24'].border = Border(right=Side(border_style='medium', color="000000"),bottom=Side(border_style='thin', color="000000"),left=Side(border_style='thin', color="000000"),top=Side(border_style='thin', color="000000"))
            ws1.font = Font(name='돋움', size=11)
            wb.save(savePath)
        except Exception as e:
            print("inputDataToExcel() - Error! : ", e)



    def typeImage(self, type, imgPath, data,wb, ws1, ws2, savePath):
        print('MyApp - typeImage()')
        try:
            self.setImage(wb=wb,ws=ws1, width=757.23, height=426.27, imgPath=imgPath, data=data, position="A25",
                          imgType="/현장사진_", savePath=savePath)
            if type == '나무':
                self.setImage(wb=wb, ws=ws2, width=738.62, height=434.17, imgPath=imgPath, data=data,
                              position="A5", imgType="/단면도_", savePath=savePath)
                self.setImage(wb=wb, ws=ws2, width=378.95, height=437.57, imgPath=imgPath, data=data,
                              position="A20", imgType="/포인트클라우드_", savePath=savePath)
                self.setImage(wb=wb, ws=ws2, width=384.25, height=434.55, imgPath=imgPath, data=data,
                              position="E20", imgType="/지상라이다_", savePath=savePath)
            elif type == '산':
                self.setImage(wb=wb, ws=ws2, width=732.95, height=431.14, imgPath=imgPath, data=data,
                              position="A5", imgType="/단면도_", savePath=savePath)
                self.setImage(wb=wb, ws=ws2, width=367.98, height=436.82, imgPath=imgPath, data=data,
                              position="A20", imgType="/포인트클라우드_", savePath=savePath)
                self.setImage(wb=wb, ws=ws2, width=365.34, height=435.30, imgPath=imgPath, data=data,
                              position="E20", imgType="/수치표고자료_", savePath=savePath)
            elif type == '건물':
                self.setImage(wb=wb, ws=ws2, width=393.32, height=414.88, imgPath=imgPath, data=data,
                              position="A5", imgType="/정사영상_", savePath=savePath)
                self.setImage(wb=wb, ws=ws2, width=385.00, height=434.17, imgPath=imgPath, data=data,
                              position="E5", imgType="/3D모델링_", savePath=savePath)
                self.setImage(wb=wb, ws=ws2, width=754.50, height=435.30, imgPath=imgPath, data=data,
                              position="A20", imgType="/단면도_", savePath=savePath)
            elif type == '기타':
                self.setImage(wb=wb, ws=ws2, width=662.98, height=456.10, imgPath=imgPath, data=data,
                              position="A5", imgType="/위치도_", savePath=savePath)
                self.setImage(wb=wb, ws=ws2, width=742.02, height=433.03, imgPath=imgPath, data=data,
                              position="A20", imgType="/단면도_", savePath=savePath)
        except Exception as e:
            print("typeImage - Error : ", e)

    def setImage(self, wb, ws, width, height, imgPath, data, position, imgType, savePath):
        try:
            imgPath = imgPath + imgType + str(data['연번']) + ".jpg"
            img = Image(imgPath)
            ws.add_image(img, position)
            img.width, img.height = width, height
            wb.save(savePath)
        except Exception as e:
            print("setImage - Error : ",imgType," 이미지가 없습니다.", e)


if __name__ == '__main__':
    # Application 객체 생성
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
