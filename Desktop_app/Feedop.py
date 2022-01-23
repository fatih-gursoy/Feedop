from PyQt5 import QtWidgets, uic, QtCore
import sys, os
import pandas as pd
import openpyxl
import firebase_admin
from firebase_admin import credentials, firestore, auth
from datetime import datetime, date
import time
import win32com.client as win32

def resource_path(relative_path):
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    return os.path.abspath(os.path.join(bundle_dir, relative_path))

class Arayuz(QtWidgets.QWidget):
    def __init__(self):
        super(Arayuz, self).__init__()
        uic.loadUi(resource_path("Feedop.ui"), self)
        self.baslangic.setDate(QtCore.QDate.currentDate())
        self.bitis.setDate(QtCore.QDate.currentDate())
        self.baslangic.dateChanged.connect(self.date_kontrol)
        self.getnews_btn.clicked.connect(self.get_click)
        self.toexcel_btn.clicked.connect(self.export_toExcel)
        self.mail_btn.clicked.connect(self.send_mail)
        self.delete_btn.clicked.connect(self.row_delete)
                
    def get_click(self):
        firebaseconfig ={      
                "type": "service_account",
                "project_id": "PROJECT_ID",
                "private_key_id": "KEY_ID",
                "private_key": "-----BEGIN PRIVATE KEY-----\n<KEY>\n-----END PRIVATE KEY-----\n",
                "client_email": "firebase-adminsdk-ngard@PROJECT_ID.iam.gserviceaccount.com",
                "client_id": "CLIENT_ID"
            }

        cred = credentials.Certificate(firebaseconfig)
        default_app = firebase_admin.initialize_app(cred)
        
        user = auth.get_user_by_email("user@gmail.com")
        
        
        db = firestore.client()

       
        baslangıc = (self.baslangic.dateTime()).toPyDateTime()
        bitis = (self.bitis.dateTime()).toPyDateTime()
        
        tarihrange = pd.date_range(baslangıc, bitis)

        docid, docs =[], []
        Haberler_Listesi = {"Haber":[],"Link":[], "Tarih":[]}
        
        #firebase'den datayı çekiyor..
        for trh in tarihrange:
            tarih = trh.strftime("%d.%#m.%Y")
            doc = db.collection('Haberler_Listesi').where('Tarih', 'array_contains', tarih).get()
            
            for item in doc:
                if not item.id in docid:
                    docid.append(item.id)
                    docs.append(item)
                        
        for doc in docs:
            for k, v in doc.to_dict().items():
                Haberler_Listesi[k] += v
        
        liste1 = list(zip(Haberler_Listesi["Haber"], Haberler_Listesi["Link"], Haberler_Listesi["Tarih"]))

        for key in Haberler_Listesi: Haberler_Listesi[key] = []
        
        for trh in tarihrange:
            tarih = trh.strftime("%d.%#m.%Y")
            for a,b,c in liste1:
                if c == tarih:
                    Haberler_Listesi["Haber"].append(a)
                    Haberler_Listesi["Link"].append(b)
                    Haberler_Listesi["Tarih"].append(c)

        #Tabloya yazıyor
        self.tablo.setRowCount(len(Haberler_Listesi["Haber"]))
        
        for column in range(0, self.tablo.columnCount(), 1):
            for col, key in enumerate(Haberler_Listesi):
                for row, value in enumerate(Haberler_Listesi[key]):            
                    if key == self.tablo.horizontalHeaderItem(column).text():
                        newitem = QtWidgets.QTableWidgetItem(value)
                        self.tablo.setItem(row, column, newitem)

        count = 0
        while count <= 100:
            time.sleep(0.001)
            count += 1
            self.progressBar.setValue(count)
            
        firebase_admin.delete_app(default_app)

    def date_kontrol(self):
        ilk_tarih = self.baslangic.date()
        self.bitis.setMinimumDate(ilk_tarih)

    def export_toExcel(self):
        xls_dict ={"Haber":[],"Link":[], "Tarih":[]}

        for column in range(0, self.tablo.columnCount(), 1):
            for row in range(0, self.tablo.rowCount(), 1):           
                header = self.tablo.horizontalHeaderItem(column).text()
                newitem = self.tablo.item(row, column).text()
                xls_dict[header].append(newitem)                

        if len(xls_dict["Haber"]) > 0:
            df = pd.DataFrame(xls_dict)
            path = ("~/Downloads/Haberler" + datetime.now().strftime("%d%m%Y_%H%M%S")+ ".xlsx")
            df.to_excel(path)
            msg = QtWidgets.QMessageBox()
            msg.setWindowTitle("Uyarı")
            msg.setText("Dosya Downloads'a eklendi." )
            msg.exec_()
        else:
            msg = QtWidgets.QMessageBox()
            msg.setWindowTitle("Uyarı")
            msg.setText("Lütfen istediğiniz tarih aralığındaki haberleri çekin")
            msg.exec_()

    def send_mail(self):
        
        Haberler = {"Haber":[],"Link":[], "Tarih":[]}

        for column in range(0, self.tablo.columnCount(), 1):
            for row in range(0, self.tablo.rowCount(), 1):           
                header = self.tablo.horizontalHeaderItem(column).text()
                newitem = self.tablo.item(row, column).text()
                Haberler[header].append(newitem) 

        mail_list = list(zip(Haberler["Haber"], Haberler["Link"]))
        mail_text ="Merhaba,<br> <br> Bu haftaki Sektör Haberlerini aşağıda görebilirsiniz. <br> <br>"

        for haber, link in mail_list:
            hyperlink = '<a href="'+ link +'">Haberin Detayı</a>'
            haber_txt = "<b>"+haber +"</b>"
            mail_text += "<ul>" + "<li>" + haber_txt + " " + hyperlink +"</li>"+ "</ul>"
        
        if len(mail_list)>0:
            outlook = win32.Dispatch("outlook.application")
            mail = outlook.Createitem(0)
            mail.to = ""
            mail.Subject = "Haftalık Sektör Haberleri"
            mail.HtmlBody = mail_text
            mail.Display(True)
        else:
            msg = QtWidgets.QMessageBox()
            msg.setWindowTitle("Uyarı")
            msg.setText("Lütfen istediğiniz tarih aralığındaki haberleri çekin")
            msg.exec_()

    def row_delete(self):
        rows = self.tablo.selectionModel().selectedRows()
        rowlist= []

        for r in rows:
            rowlist.append(r.row())

        rowlist = sorted(rowlist, reverse = True)

        for i in rowlist:
            self.tablo.removeRow(i)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle('Fusion')
    Pencere = Arayuz()
    Pencere.resize(1200,600)
    Pencere.show()
    sys.exit(app.exec_())