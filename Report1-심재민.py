import pymysql
import sys
from PyQt5.QtWidgets import *
import csv
import json
import xml.etree.ElementTree as ET
from decimal import *


class JSONEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, Decimal):
            return str(obj)
        return json.JSONEncoder.default(self, obj)


class DB_Utils:
    # SQL 검색문(sql과 params)을 전달받아, 실행하는 메소드
    def queryExecutor(self, db, sql, params):
        conn = pymysql.connect(host='localhost', user='guest', password='bemyguest', db=db, charset='utf8')

        try:
            with conn.cursor(pymysql.cursors.DictCursor) as cursor:
                cursor.execute(sql, params)

                rows = cursor.fetchall()
                return rows
        except Exception as e:
            print(e)
            print(type(e))
        finally:
            cursor.close()
            conn.close()


class DB_Queries:
    # 검색문을 각각 하나의 메소드로 정의
    def selectCustomersName(self):
        sql = "SELECT name FROM customers"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectCustomersCountry(self):
        sql = "SELECT country FROM customers"
        params = ()

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectCustomersCity(self, country=None):
        sql = "SELECT city FROM customers"
        params = ()

        if country == "ALL":
            country = None

        if country:
            sql += " WHERE country = %s"
            params = (country)

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectSearchedOrder(self, key=None, value=None):
        sql = """SELECT o.orderNo, o.orderDate, o.requiredDate, o.shippedDate, o.status, c.name AS customer, o.comments
                 FROM customers as c INNER JOIN orders as o ON c.customerId = o.customerId"""

        params = ()
        if value == "ALL":
            pass

        elif key == "name":
            sql += " WHERE name = %s"
            params = (value)

        elif key == "country":
            sql += " WHERE country = %s"
            params = (value)

        elif key == "city":
            sql += " WHERE city = %s"
            params = (value)

        sql += " ORDER BY o.orderNo"

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows

    def selectOrderDetail(self, orderNo):
        sql = """SELECT d.orderLineNo, p.productCode, p.name AS productName,
                 d.quantity, d.priceEach, d.quantity * d.priceEach as 상품주문액
                 FROM orders as o INNER JOIN orderDetails as d ON o.orderNo = d.orderNo
                 INNER JOIN products as p on d.productCode = p.productCode
                 WHERE o.orderNo = %s ORDER BY orderLineNo"""
        params = (orderNo)

        util = DB_Utils()
        rows = util.queryExecutor(db="classicmodels", sql=sql, params=params)
        return rows


# 주문 검색 화면
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.key = ""
        self.value = ""
        self.setupUI()

    def setupUI(self):
        # 윈도우 설정
        self.setWindowTitle("주문 검색")
        self.setGeometry(0, 0, 1000, 700)

        # 필요한 모든 위젯 생성
        # 첫 번째 줄: 고객, 국가, 도시
        self.labelCustomer = QLabel("고객: ")
        self.comboBoxCustomer = QComboBox(self)
        query = DB_Queries()
        self.setComboBoxData(query.selectCustomersName(), "name", self.comboBoxCustomer)

        self.labelCountry = QLabel("국가: ")
        self.comboBoxCountry = QComboBox(self)
        self.setComboBoxData(query.selectCustomersCountry(), "country", self.comboBoxCountry)

        self.labelCity = QLabel("도시: ")
        self.comboBoxCity = QComboBox(self)
        self.setComboBoxData(query.selectCustomersCity(), "city", self.comboBoxCity)

        # 두 번째 줄: 검색된 주문의 개수
        self.labelNumOrderText = QLabel("검색된 주문의 개수: ")
        self.labelNumOrder = QLabel()

        # 검색, 초기화 버튼
        self.btnSearch = QPushButton("검색")
        self.btnSearch.clicked.connect(self.btnSearchClicked)
        self.btnClear = QPushButton("초기화")
        self.btnClear.clicked.connect(self.btnClearClicked)

        groupBoxSearch = QGroupBox("주문 검색")

        # 주문 리스트 테이블
        self.tableWidgetOrderList = QTableWidget(self)
        self.tableWidgetOrderList.doubleClicked.connect(
                lambda: self.tableWidgetDoubleClicked(self.tableWidgetOrderList.currentRow()))
        self.setTableWidgetData()

        # 레이아웃의 생성, 위젯 연결
        layoutDialog = QGridLayout()
        layoutNumOrder = QGridLayout()
        layoutSearch = QGridLayout()
        layoutSearchOrder = QVBoxLayout()
        layout = QVBoxLayout()

        layoutDialog.addWidget(self.labelCustomer, 0, 0)
        layoutDialog.addWidget(self.comboBoxCustomer, 0, 1)
        layoutDialog.addWidget(self.labelCountry, 0, 2)
        layoutDialog.addWidget(self.comboBoxCountry, 0, 3)
        layoutDialog.addWidget(self.labelCity, 0, 4)
        layoutDialog.addWidget(self.comboBoxCity, 0, 5)
        layoutNumOrder.addWidget(self.labelNumOrderText, 0, 0)
        layoutNumOrder.addWidget(self.labelNumOrder, 0, 1)
        layoutSearch.addLayout(layoutDialog, 0, 0)
        layoutSearch.addLayout(layoutNumOrder, 1, 0)
        layoutSearch.addWidget(self.btnSearch, 0, 1)
        layoutSearch.addWidget(self.btnClear, 1, 1)
        layoutSearchOrder.addLayout(layoutSearch)
        layoutSearchOrder.addWidget(self.tableWidgetOrderList)

        groupBoxSearch.setLayout(layoutSearchOrder)
        layout.addWidget(groupBoxSearch)

        # 레이아웃 설정
        self.setLayout(layout)

    def comboBoxChanged(self, key, value):
        if key == "country":
            query = DB_Queries()
            if value == "ALL":
                self.setComboBoxData(query.selectCustomersCity(), "city", self.comboBoxCity)
            else:
                self.setComboBoxData(query.selectCustomersCity(value), "city", self.comboBoxCity)
        self.key = key
        self.value = value

    def btnSearchClicked(self):
        if self.key == "city" and self.value == "ALL" and self.comboBoxCountry.currentText() != "ALL":
            self.key = "country"
            self.value = self.comboBoxCountry.currentText()
        self.setTableWidgetData(self.key, self.value)

    def btnClearClicked(self):
        self.comboBoxCustomer.setCurrentIndex(0)
        self.comboBoxCountry.setCurrentIndex(0)
        self.comboBoxCity.setCurrentIndex(0)
        self.setTableWidgetData()

    def setComboBoxData(self, rows, key, comboBox):
        items = list(set(str(row[columnName]) for row in rows for columnName in row))
        items.sort()
        items.insert(0, "ALL")
        comboBox.clear()
        comboBox.addItems(items)
        comboBox.currentTextChanged.connect(lambda: self.comboBoxChanged(key, comboBox.currentText()))

    def setTableWidgetData(self, key=None, value=None):
        query = DB_Queries()
        rows = query.selectSearchedOrder(key, value)
        self.labelNumOrder.setText(str(len(rows)))  # 검색된 주문의 개수 update

        columnNames = ["orderNo", "orderDate", "requiredDate", "shippedDate",
                       "status", "customer", "comments"]
        self.tableWidgetOrderList.clear()
        self.tableWidgetOrderList.setColumnCount(len(columnNames))
        self.tableWidgetOrderList.setRowCount(max(len(rows), 1))
        self.tableWidgetOrderList.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidgetOrderList.setHorizontalHeaderLabels(columnNames)

        for rowIdx, row in enumerate(rows):
            for columnIdx, (k, v) in enumerate(row.items()):
                item = QTableWidgetItem(str(v))
                self.tableWidgetOrderList.setItem(rowIdx, columnIdx, item)

        self.tableWidgetOrderList.resizeRowsToContents()
        self.tableWidgetOrderList.resizeColumnsToContents()

    def tableWidgetDoubleClicked(self, row):
        orderNo = self.tableWidgetOrderList.item(row, 0)
        if orderNo:
            dialogue = SubWindow(orderNo.text())
            dialogue.exec_()


class SubWindow(QDialog):
    def __init__(self, orderNo):
        super().__init__()
        self.orderNo = orderNo
        self.setupUI()

    def setupUI(self):
        # 윈도우 설정
        self.setWindowTitle("주문 상세 내역")
        self.setGeometry(100, 100, 700, 500)

        # 필요한 모든 위젯 생성
        # 주문번호, 상품개수, 주문액
        self.labelOrderNoText = QLabel("주문번호: ")
        self.labelNumProductText = QLabel("상품개수: ")
        self.labelPriceText = QLabel("주문액: ")
        self.labelOrderNo = QLabel(self.orderNo)
        self.labelNumProduct = QLabel()
        self.labelPrice = QLabel()

        # 파일 출력
        self.radioBtnCSV = QRadioButton("CSV", self)
        self.radioBtnCSV.setChecked(True)
        self.radioBtnJSON = QRadioButton("JSON", self)
        self.radioBtnXML = QRadioButton("XML", self)
        self.btnSave = QPushButton("저장")
        self.btnSave.clicked.connect(lambda: self.btnSaveClicked(self.orderNo))

        groupBoxDetail = QGroupBox("주문 상세 내역")
        groupBoxFileSave = QGroupBox("파일 출력")

        # 상품 리스트 테이블
        self.tableWidgetProductList = QTableWidget(self)
        self.setTableWidgetData()

        # 레이아웃의 생성, 위젯 연결
        layoutDetailInfo = QGridLayout()
        layoutProductList = QVBoxLayout()
        layoutDetail = QVBoxLayout()
        layoutFileSave = QGridLayout()
        layout = QVBoxLayout()

        layoutDetailInfo.addWidget(self.labelOrderNoText, 0, 0)
        layoutDetailInfo.addWidget(self.labelOrderNo, 0, 1)
        layoutDetailInfo.addWidget(self.labelNumProductText, 0, 2)
        layoutDetailInfo.addWidget(self.labelNumProduct, 0, 3)
        layoutDetailInfo.addWidget(self.labelPriceText, 0, 4)
        layoutDetailInfo.addWidget(self.labelPrice, 0, 5)
        layoutProductList.addWidget(self.tableWidgetProductList)
        layoutDetail.addLayout(layoutDetailInfo)
        layoutDetail.addLayout(layoutProductList)

        layoutFileSave.addWidget(self.radioBtnCSV, 0, 0)
        layoutFileSave.addWidget(self.radioBtnJSON, 0, 1)
        layoutFileSave.addWidget(self.radioBtnXML, 0, 2)
        layoutFileSave.addWidget(self.btnSave, 1, 3)

        groupBoxDetail.setLayout(layoutDetail)
        groupBoxFileSave.setLayout(layoutFileSave)
        layout.addWidget(groupBoxDetail)
        layout.addWidget(groupBoxFileSave)

        # 레이아웃 설정
        self.setLayout(layout)

    def setTableWidgetData(self):
        query = DB_Queries()
        rows = query.selectOrderDetail(self.orderNo)
        self.labelNumProduct.setText(str(len(rows)))  # 상품 개수 update

        columnNames = ["orderLineNo", "productCode", "productName",
                       "quantity", "priceEach", "상품주문액"]
        self.tableWidgetProductList.clear()
        self.tableWidgetProductList.setColumnCount(len(columnNames))
        self.tableWidgetProductList.setRowCount(max(len(rows), 1))
        self.tableWidgetProductList.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tableWidgetProductList.setHorizontalHeaderLabels(columnNames)

        totalPrice = 0
        for rowIdx, row in enumerate(rows):
            for columnIdx, (k, v) in enumerate(row.items()):
                item = QTableWidgetItem(str(v))
                self.tableWidgetProductList.setItem(rowIdx, columnIdx, item)
            totalPrice += row["quantity"] * row["priceEach"]

        self.tableWidgetProductList.resizeRowsToContents()
        self.tableWidgetProductList.resizeColumnsToContents()

        self.labelPrice.setText(str(totalPrice))  # 주문액 update

    def btnSaveClicked(self, orderNo):
        if self.radioBtnCSV.isChecked():
            with open(orderNo+'.csv', 'w', encoding='utf-8-sig', newline='') as f:
                wr = csv.writer(f)
                query = DB_Queries()
                rows = query.selectOrderDetail(orderNo)

                # 테이블 헤더를 출력
                columnNames = list(rows[0].keys())
                wr.writerow(columnNames)

                # 테이블 내용을 출력
                for row in rows:
                    values = list(row.values())
                    wr.writerow(values)

        elif self.radioBtnJSON.isChecked():
            query = DB_Queries()
            rows = query.selectOrderDetail(orderNo)

            newDict = dict(Orders=rows)

            # JSON 화일에 쓰기
            with open(orderNo+'.json', 'w', encoding='utf-8') as f:
                json.dump(newDict, f, indent=4, ensure_ascii=False, cls=JSONEncoder)

        elif self.radioBtnXML.isChecked():
            query = DB_Queries()
            rows = query.selectOrderDetail(orderNo)

            newDict = dict(Orders=rows)

            # XDM 트리 생성
            tableName = list(newDict.keys())[0]
            tableRows = list(newDict.values())[0]

            rootElement = ET.Element('TABLE')
            rootElement.attrib['name'] = tableName

            for row in tableRows:
                rowElement = ET.Element('ROW')
                rootElement.append(rowElement)

                for columnName in list(row.keys()):
                    if row[columnName] == None:
                        rowElement.attrib[columnName] = ''
                    else:
                        rowElement.attrib[columnName] = str(row[columnName])

            # XDM 트리를 화일에 출력
            ET.ElementTree(rootElement).write(orderNo+'.xml', encoding='utf-8', xml_declaration=True)



#########################################

if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    mainWindow.show()
    app.exec_()



