import pymongo,sys,os,openpyxl,re
import SoftwareInterfaceStyle as sis
from pathlib import Path
from datetime import datetime, timedelta
from PyQt6.QtWidgets import (QWidget,QApplication,QGridLayout,QVBoxLayout,QLabel,QLineEdit,QPushButton,QComboBox,QTableWidget,QAbstractItemView,
                             QHeaderView,QMessageBox,QTextEdit,QFileDialog,QCheckBox,QSpinBox,QMenu,QInputDialog,QTableWidgetItem)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap,QAction,QCursor,QIcon
import Memberships_Language as lang
import ItalyCityCode as ICC

# Versione 1.0.2-r4

# Variabili globali

heading = ""
dbclient = ""
interface = ""
logo_path = ""
icon_path = ""
language = ""
first_start_application = 0

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.db = dbclient["Memberships"] # Apertura database
        
        # *-*-* Impostazioni iniziali *-*-*
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"{heading}")
        self.setMinimumSize(640, 540) # Risoluzione minima per schermi piccoli
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.showMaximized()
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        
        # *-*-* Grafica dei widgets *-*-*
        
        self.setStyleSheet(sis.interface_style(interface))
        
        # *-*-* Widgets *-*-*
        
        # Logo (Parte superiore sinistra)
        
        L_logo = QLabel(self)
        IMG_Pixmap = QPixmap(logo_path)
        L_logo.setPixmap(IMG_Pixmap)
        L_logo.resize(IMG_Pixmap.width(), IMG_Pixmap.height())
        self.lay.addWidget(L_logo, 0, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Label titolo (Parte superiore centrale)
        
        L_title = QLabel(self, text=f"{heading}")
        L_title.setAccessibleName("an_title")
        self.lay.addWidget(L_title, 0, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Checkbox per ricerca automatica (Parte superiore centrale)
        
        self.CHB_auto_search = QCheckBox(self, text=lang.msg(language, 0, "MainWindow"))
        self.lay.addWidget(self.CHB_auto_search, 0, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
        
        # Pulsante Database
        
        self.B_database_management = QPushButton(self, text=lang.msg(language, 1, "MainWindow"))
        self.B_database_management.clicked.connect(self.database_management_open)
        self.lay.addWidget(self.B_database_management, 0, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Pulsante Opzioni (Parte superiore destra)
        
        self.B_options_menu = QPushButton(self, text=lang.msg(language, 2, "MainWindow"))
        self.B_options_menu.clicked.connect(self.options_menu_open)
        self.lay.addWidget(self.B_options_menu, 0, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Casella codice fiscale (Parte centrale sinistra)
        
        self.LE_tax_id_code = QLineEdit(self)
        self.LE_tax_id_code.setPlaceholderText(lang.msg(language, 3, "MainWindow"))
        self.LE_tax_id_code.textChanged.connect(self.auto_search)
        self.LE_tax_id_code.textChanged.connect(self.auto_field_change)
        self.lay.addWidget(self.LE_tax_id_code, 1, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella nome (Parte centrale sinistra)
        
        self.LE_name = QLineEdit(self)
        self.LE_name.setPlaceholderText(lang.msg(language, 4, "MainWindow"))
        self.LE_name.textChanged.connect(self.auto_search)
        self.lay.addWidget(self.LE_name, 2, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella cognome (Parte centrale sinistra)
        
        self.LE_surname = QLineEdit(self)
        self.LE_surname.setPlaceholderText(lang.msg(language, 5, "MainWindow"))
        self.LE_surname.textChanged.connect(self.auto_search)
        self.lay.addWidget(self.LE_surname, 3, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella data di nascita (Parte centrale sinistra)
        
        self.LE_date_of_birth = QLineEdit(self)
        self.LE_date_of_birth.setPlaceholderText(lang.msg(language, 6, "MainWindow"))
        self.lay.addWidget(self.LE_date_of_birth, 4, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella luogo di nascita (Parte centrale sinistra)
        
        self.LE_birth_place = QLineEdit(self)
        self.LE_birth_place.setPlaceholderText(lang.msg(language, 7, "MainWindow"))
        self.lay.addWidget(self.LE_birth_place, 5, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Box sesso (Parte centrale sinistra)
        
        self.CB_sex = QComboBox(self)
        self.CB_sex.addItems([lang.msg(language, 8, "MainWindow"), lang.msg(language, 9, "MainWindow")])
        self.lay.addWidget(self.CB_sex, 6, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella comune di residenza (Parte centrale sinistra)
        
        self.LE_city_of_residence = QLineEdit(self)
        self.LE_city_of_residence.setPlaceholderText(lang.msg(language, 10, "MainWindow"))
        self.lay.addWidget(self.LE_city_of_residence, 7, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella indirizzo di residenza (Parte centrale sinistra)
        
        self.LE_residential_address = QLineEdit(self)
        self.LE_residential_address.setPlaceholderText(lang.msg(language, 11, "MainWindow"))
        self.lay.addWidget(self.LE_residential_address, 8, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella CAP (Parte centrale sinistra)
        
        self.LE_postal_code = QLineEdit(self)
        self.LE_postal_code.setPlaceholderText(lang.msg(language, 12, "MainWindow"))
        self.lay.addWidget(self.LE_postal_code, 9, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella e-mail (Parte centrale sinistra)
        
        self.LE_email = QLineEdit(self)
        self.LE_email.setPlaceholderText(lang.msg(language, 13, "MainWindow"))
        self.lay.addWidget(self.LE_email, 10, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella numero tessera (Parte centrale sinistra)
        
        self.LE_card_number = QLineEdit(self)
        self.LE_card_number.setPlaceholderText(lang.msg(language, 14, "MainWindow"))
        self.lay.addWidget(self.LE_card_number, 11, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante inserisci (Parte bassa sinistra)
        
        self.B_insert = QPushButton(self, text=f"{lang.msg(language, 15, 'MainWindow')} >>")
        self.B_insert.clicked.connect(self.insert_into_db)
        self.lay.addWidget(self.B_insert, 12, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante cerca (Parte bassa sinistra)
        
        self.B_search = QPushButton(self, text=lang.msg(language, 16, "MainWindow"))
        self.B_search.clicked.connect(self.search_db)
        self.lay.addWidget(self.B_search, 12, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante esporta in Excel
        
        self.B_excel_export = QPushButton(self, text=lang.msg(language, 17, "MainWindow"))
        self.B_excel_export.clicked.connect(self.excel_export)
        self.lay.addWidget(self.B_excel_export, 12, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Box database (Parte centrale destra)
        
        self.TE_db_response = QTextEdit(self)
        self.TE_db_response.setReadOnly(True)
        self.TE_db_response.setPlainText(f"{lang.msg(language, 18, 'MainWindow')} {heading}\n\n{lang.msg(language, 19, 'MainWindow')}")
        self.lay.addWidget(self.TE_db_response, 1, 3, 12, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
    
    # *-*-* Funzioni per il ridimensionamento della finestra *-*-*

    def resizeEvent(self, event):
        W_width = self.width() / 3
        W_height = self.height() / 3
        
        try:
            # Parte sinistra
            self.LE_tax_id_code.setMinimumWidth(int(W_width)-50)
            self.LE_name.setMinimumWidth(int(W_width)-50)
            self.LE_surname.setMinimumWidth(int(W_width)-50)
            self.LE_date_of_birth.setMinimumWidth(int(W_width)-50)
            self.LE_birth_place.setMinimumWidth(int(W_width)-50)
            self.CB_sex.setMinimumWidth(int(W_width)-50)
            self.LE_city_of_residence.setMinimumWidth(int(W_width)-50)
            self.LE_residential_address.setMinimumWidth(int(W_width)-50)
            self.LE_postal_code.setMinimumWidth(int(W_width)-50)
            self.LE_email.setMinimumWidth(int(W_width)-50)
            self.LE_card_number.setMinimumWidth(int(W_width)-50)
            # Parte destra
            self.TE_db_response.setMinimumSize(int(W_width * 2), int(W_height * 3) - 100)
        except AttributeError: pass
    
    # *-*-* Funzione inserimento *-*-*
    
    def insert_into_db(self):
        date = datetime.now()
        # Trasformazione dati inseriti
        
        self.tax_id_code = self.LE_tax_id_code.text().upper().replace(" ", "")
        self.name = self.LE_name.text().upper().strip()
        self.surname = self.LE_surname.text().upper().strip()
        self.date_of_birth = self.LE_date_of_birth.text().replace(" ", "")
        self.birth_place = self.LE_birth_place.text().upper().strip()
        self.sex = self.CB_sex.currentText()
        self.city_of_residence = self.LE_city_of_residence.text().upper().strip()
        self.residential_address = self.LE_residential_address.text().upper().strip()
        self.postal_code = self.LE_postal_code.text().replace(" ", "")
        self.email = self.LE_email.text().replace(" ", "")
        self.card_number = self.LE_card_number.text().replace(" ", "")
        self.date_of_membership = date.strftime("%Y%m%d")
        
        # Controllo caselle obbigatorie vuote
        
        if self.tax_id_code == "" or self.name == "" or self.surname == "":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 21, "MainWindow"))
            return err_msg.exec()
        
        # Controllo codice fiscale
        
        if len(self.tax_id_code) != 16:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 22, "MainWindow"))
            return err_msg.exec()
        if self.has_numbers(self.tax_id_code[0:6]) == True or self.has_numbers(self.tax_id_code[8]) == True or self.has_numbers(self.tax_id_code[11]) == True or self.has_numbers(self.tax_id_code[15]) == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 22, "MainWindow"))
            return err_msg.exec()
        if self.tax_id_code[6:8].isdigit() == False or self.tax_id_code[9:11].isdigit() == False or self.tax_id_code[12:15].isdigit() == False:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 22, "MainWindow"))
            return err_msg.exec()
        tax_id_code_char_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9"]
        for char in self.tax_id_code:
            if char not in tax_id_code_char_list:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 22, "MainWindow"))
                return err_msg.exec()
        
        # Controllo nome
        
        if self.has_numbers(self.name) == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 23, "MainWindow"))
            return err_msg.exec()
        
        # Controllo cognome
        
        if self.has_numbers(self.surname) == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 24, "MainWindow"))
            return err_msg.exec()
        
        # Trasformazione caselle vuote e data tesseramento
        
        if self.date_of_birth == "": self.date_of_birth = "-"
        if self.birth_place == "": self.birth_place = "-"
        if self.city_of_residence == "": self.city_of_residence = "-"
        if self.residential_address == "": self.residential_address = "-"
        if self.postal_code == "": self.postal_code = "-"
        if self.email == "": self.email = "-"
        if self.card_number == "":
            self.card_number = "-"
            self.date_of_membership = "-"
        
        # Controllo presenza nel database
        
        col = self.db["cards"]
        if col.find_one({"tax_id_code": self.tax_id_code, "name": self.name, "surname": self.surname}, {"_id": 0}) != None:
            msg = QMessageBox(self)
            msg.setWindowTitle(lang.msg(language, 25, "MainWindow"))
            msg.setText(lang.msg(language, 26, "MainWindow"))
            msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.No)
            msg.buttonClicked.connect(self.modify_db)
            return msg.exec()
        
        # Inserimento nel database
        
        col.insert_one({
            "tax_id_code": self.tax_id_code,
            "name": self.name,
            "surname": self.surname,
            "date_of_birth": self.date_of_birth,
            "birth_place": self.birth_place,
            "sex": self.sex,
            "city_of_residence": self.city_of_residence,
            "residential_address": self.residential_address,
            "postal_code": self.postal_code,
            "email": self.email,
            "card_number": self.card_number,
            "date_of_membership": self.date_of_membership
        })
        
        self.TE_db_response.clear()
        date_of_membership = self.date_of_membership # Trasformazione data in formato leggibile
        if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
        self.TE_db_response.insertPlainText(f"""-*-* {lang.msg(language, 27, 'MainWindow')} *-*-
\n{lang.msg(language, 3, 'MainWindow')}: {self.tax_id_code}
{lang.msg(language, 4, 'MainWindow')}: {self.name}
{lang.msg(language, 5, 'MainWindow')}: {self.surname}
{lang.msg(language, 6, 'MainWindow')}: {self.date_of_birth}
{lang.msg(language, 7, 'MainWindow')}: {self.birth_place}
{lang.msg(language, 28, 'MainWindow')}: {self.sex}
{lang.msg(language, 10, 'MainWindow')}: {self.city_of_residence}
{lang.msg(language, 11, 'MainWindow')}: {self.residential_address}
{lang.msg(language, 12, 'MainWindow')}: {self.postal_code}
{lang.msg(language, 13, 'MainWindow')}: {self.email}
{lang.msg(language, 14, 'MainWindow')}: {self.card_number}
{lang.msg(language, 29, 'MainWindow')}: {date_of_membership}""")
        self.clear_box()
    
    # *-*-* Funzione modifica database *-*-*
    
    def modify_db(self, button):
        if button.text() == "&OK" or button.text() == "OK":
            col = self.db["cards"]
            
            # Impostazione dati lasciati vuoti
            
            person_data = col.find_one({"tax_id_code": self.tax_id_code, "name": self.name, "surname": self.surname}, {"_id": 0})
            if self.date_of_birth == "-": self.date_of_birth = person_data["date_of_birth"]
            if self.birth_place == "-": self.birth_place = person_data["birth_place"]
            if self.city_of_residence == "-": self.city_of_residence = person_data["city_of_residence"]
            if self.residential_address == "-": self.residential_address = person_data["residential_address"]
            if self.postal_code == "-": self.postal_code = person_data["postal_code"]
            if self.email == "-": self.email = person_data["email"]
            if self.card_number == "-":
                self.card_number = person_data["card_number"]
                self.date_of_membership = person_data["date_of_membership"]
            
            # Update del database
            
            col.update_one({"tax_id_code": self.tax_id_code, "name": self.name, "surname": self.surname}, {"$set": {
                "date_of_birth": self.date_of_birth,
                "birth_place": self.birth_place,
                "sex": self.sex,
                "city_of_residence": self.city_of_residence,
                "residential_address": self.residential_address,
                "postal_code": self.postal_code,
                "email": self.email,
                "card_number": self.card_number,
                "date_of_membership": self.date_of_membership
            }})
            
            self.TE_db_response.clear()
            date_of_membership = self.date_of_membership # Trasformazione data in formato leggibile
            if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
            self.TE_db_response.insertPlainText(f"""-*-* {lang.msg(language, 30, 'MainWindow')} *-*-
\n{lang.msg(language, 3, 'MainWindow')}: {self.tax_id_code}
{lang.msg(language, 4, 'MainWindow')}: {self.name}
{lang.msg(language, 5, 'MainWindow')}: {self.surname}
{lang.msg(language, 6, 'MainWindow')}: {self.date_of_birth}
{lang.msg(language, 7, 'MainWindow')}: {self.birth_place}
{lang.msg(language, 28, 'MainWindow')}: {self.sex}
{lang.msg(language, 10, 'MainWindow')}: {self.city_of_residence}
{lang.msg(language, 11, 'MainWindow')}: {self.residential_address}
{lang.msg(language, 12, 'MainWindow')}: {self.postal_code}
{lang.msg(language, 13, 'MainWindow')}: {self.email}
{lang.msg(language, 14, 'MainWindow')}: {self.card_number}
{lang.msg(language, 29, 'MainWindow')}: {date_of_membership}""")
            self.clear_box()
    
    # -*-* Funzione cambio automatico caselle
    
    def auto_field_change(self):
        tax_id_code = self.LE_tax_id_code.text().upper().replace(" ", "")
        if len(tax_id_code) == 16:
            tax_id_code_to_month = {"A":"01","B":"02","C":"03","D":"04","E":"05","H":"06","L":"07","M":"08","P":"09","R":"10","S":"11","T":"12"}
            # Casella data di nascita
            try:
                current_year = datetime.now()
                current_year = current_year.strftime("%Y")
                day = int(tax_id_code[9:11])
                if day > 40: day = day -40
                if day < 10: day = f"0{day}"
                month = tax_id_code_to_month[tax_id_code[8]]
                year = int(f"{current_year[:2]}{tax_id_code[6:8]}")
                if int(current_year[2:]) < int(tax_id_code[6:8]): year -= 100
                date_of_birth = f"{day}/{month}/{year}"
                self.LE_date_of_birth.setText(date_of_birth)
            except: pass
            
            # Casella sesso
            try:
                date = int(tax_id_code[9:11])
                if date > 40:
                    if language == "ITALIANO": self.CB_sex.setCurrentText("FEMMINA")
                    if language == "ENGLISH": self.CB_sex.setCurrentText("FEMALE")
                else:
                    if language == "ITALIANO": self.CB_sex.setCurrentText("MASCHIO")
                    if language == "ENGLISH": self.CB_sex.setCurrentText("MALE")
            except: pass
            
            # Casella luogo di nascita
            
            try:
                self.LE_birth_place.setText(ICC.GetCity(tax_id_code[11:15]))
            except: pass
    
    # -*-* Funzione ricerca automatica nel database *-*-
    
    def auto_search(self):
        # Controllo se la ricerca automatica è attiva
        
        if self.CHB_auto_search.isChecked() == False: return
        
        # Variabili
        
        tax_id_code = self.LE_tax_id_code.text().upper().replace(" ", "")
        name = self.LE_name.text().upper().strip()
        surname = self.LE_surname.text().upper().strip()
        
        # Controllo codice fiscale
        
        if len(tax_id_code) == 16:
            if self.has_numbers(tax_id_code[0:6]) == True or self.has_numbers(tax_id_code[8]) == True or self.has_numbers(tax_id_code[11]) == True or self.has_numbers(tax_id_code[15]) == True:
                tax_id_code = "-"
            if tax_id_code[6:8].isdigit() == False or tax_id_code[9:11].isdigit() == False or tax_id_code[12:15].isdigit() == False:
                tax_id_code = "-"
            tax_id_code_char_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9"]
            for char in tax_id_code:
                if char not in tax_id_code_char_list:
                    tax_id_code = "-"
                    break
        else: tax_id_code = "-"
        
        # Controllo nome
        
        if self.has_numbers(name) == True or len(name) < 3: name = "-"
        
        # Controllo cognome
        
        if self.has_numbers(surname) == True or len(surname) < 3: surname = "-"
        
        # Blocco della funzione se niente è inserito
        
        if tax_id_code == "-" and name == "-" and surname == "-": return
        
        # Ricerca dati nel database
        
        col = self.db["cards"]
        
        # Solo un inserimento
        
        if tax_id_code != "-" and name == "-" and surname == "-": people_founded = col.find({"tax_id_code": tax_id_code}, {"_id": 0}).limit(10)
        if tax_id_code == "-" and name != "-" and surname == "-": people_founded = col.find({"name": {"$regex": name}}, {"_id": 0}).limit(10)
        if tax_id_code == "-" and name == "-" and surname != "-": people_founded = col.find({"surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        
        # Due inserimenti
        
        if tax_id_code != "-" and name != "-" and surname == "-": people_founded = col.find({"tax_id_code": tax_id_code, "name": {"$regex": name}}, {"_id": 0}).limit(10)
        if tax_id_code != "-" and name == "-" and surname != "-": people_founded = col.find({"tax_id_code": tax_id_code, "surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        if tax_id_code == "-" and name != "-" and surname != "-": people_founded = col.find({"name": {"$regex": name}, "surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        
        # Tre inserimenti
        
        if tax_id_code != "-" and name != "-" and surname != "-": people_founded = col.find({"tax_id_code": tax_id_code, "name": {"$regex": name}, "surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        
        # Aggiunta alla casella di risposta
        
        people_founded = list(people_founded)
        
        if len(people_founded) != 0:
            self.TE_db_response.clear()
            self.TE_db_response.insertPlainText(f"-*-* {lang.msg(language, 31, 'MainWindow')} *-*-\n\n")
            for person in people_founded:
                date_of_membership = person["date_of_membership"] # Trasformazione data in formato leggibile
                if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
                self.TE_db_response.append(f"""{lang.msg(language, 3, 'MainWindow')}: {person['tax_id_code']}
{lang.msg(language, 4, 'MainWindow')}: {person['name']}
{lang.msg(language, 5, 'MainWindow')}: {person['surname']}
{lang.msg(language, 6, 'MainWindow')}: {person['date_of_birth']}
{lang.msg(language, 7, 'MainWindow')}: {person['birth_place']}
{lang.msg(language, 28, 'MainWindow')}: {person['sex']}
{lang.msg(language, 10, 'MainWindow')}: {person['city_of_residence']}
{lang.msg(language, 11, 'MainWindow')}: {person['residential_address']}
{lang.msg(language, 12, 'MainWindow')}: {person['postal_code']}
{lang.msg(language, 13, 'MainWindow')}: {person['email']}
{lang.msg(language, 14, 'MainWindow')}: {person['card_number']}
{lang.msg(language, 29, 'MainWindow')}: {date_of_membership}\n\n------------------------------\n\n""")
    
    # *-*-* Funzione ricerca manuale nel database *-*-*
    
    def search_db(self):
        # Controllo se la ricerca automatica è attiva
        
        if self.CHB_auto_search.isChecked() == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 32, "MainWindow"))
            return err_msg.exec()
        
        # Variabili
        
        tax_id_code = self.LE_tax_id_code.text().upper().replace(" ", "")
        name = self.LE_name.text().upper().strip()
        surname = self.LE_surname.text().upper().strip()
        
        # Controllo codice fiscale
        
        if len(tax_id_code) == 16:
            if self.has_numbers(tax_id_code[0:6]) == True or self.has_numbers(tax_id_code[8]) == True or self.has_numbers(tax_id_code[11]) == True or self.has_numbers(tax_id_code[15]) == True:
                tax_id_code = "-"
            if tax_id_code[6:8].isdigit() == False or tax_id_code[9:11].isdigit() == False or tax_id_code[12:15].isdigit() == False:
                tax_id_code = "-"
            tax_id_code_char_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9"]
            for char in tax_id_code:
                if char not in tax_id_code_char_list:
                    tax_id_code = "-"
                    break
        else: tax_id_code = "-"
        
        # Controllo nome
        
        if self.has_numbers(name) == True or len(name) < 3: name = "-"
        
        # Controllo cognome
        
        if self.has_numbers(surname) == True or len(surname) < 3: surname = "-"
        
        # Blocco della funzione se niente è inserito
        
        if tax_id_code == "-" and name == "-" and surname == "-":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 33, "MainWindow"))
            return err_msg.exec()
        
        # Ricerca dati nel database
        
        col = self.db["cards"]
        
        # Solo un inserimento
        
        if tax_id_code != "-" and name == "-" and surname == "-": people_founded = col.find({"tax_id_code": tax_id_code}, {"_id": 0}).limit(10)
        if tax_id_code == "-" and name != "-" and surname == "-": people_founded = col.find({"name": {"$regex": name}}, {"_id": 0}).limit(10)
        if tax_id_code == "-" and name == "-" and surname != "-": people_founded = col.find({"surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        
        # Due inserimenti
        
        if tax_id_code != "-" and name != "-" and surname == "-": people_founded = col.find({"tax_id_code": tax_id_code, "name": {"$regex": name}}, {"_id": 0}).limit(10)
        if tax_id_code != "-" and name == "-" and surname != "-": people_founded = col.find({"tax_id_code": tax_id_code, "surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        if tax_id_code == "-" and name != "-" and surname != "-": people_founded = col.find({"name": {"$regex": name}, "surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        
        # Tre inserimenti
        
        if tax_id_code != "-" and name != "-" and surname != "-": people_founded = col.find({"tax_id_code": tax_id_code, "name": {"$regex": name}, "surname": {"$regex": surname}}, {"_id": 0}).limit(10)
        
        # Aggiunta alla casella di risposta
        
        people_founded = list(people_founded)
        
        if len(people_founded) == 0:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 25, "MainWindow"))
            err_msg.setText(lang.msg(language, 34, "MainWindow"))
            return err_msg.exec()
        else:
            self.TE_db_response.clear()
            self.TE_db_response.insertPlainText(f"-*-* {lang.msg(language, 31, 'MainWindow')} *-*-\n\n")
            for person in people_founded:
                date_of_membership = person["date_of_membership"] # Trasformazione data in formato leggibile
                if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
                self.TE_db_response.append(f"""{lang.msg(language, 3, 'MainWindow')}: {person['tax_id_code']}
{lang.msg(language, 4, 'MainWindow')}: {person['name']}
{lang.msg(language, 5, 'MainWindow')}: {person['surname']}
{lang.msg(language, 6, 'MainWindow')}: {person['date_of_birth']}
{lang.msg(language, 7, 'MainWindow')}: {person['birth_place']}
{lang.msg(language, 28, 'MainWindow')}: {person['sex']}
{lang.msg(language, 10, 'MainWindow')}: {person['city_of_residence']}
{lang.msg(language, 11, 'MainWindow')}: {person['residential_address']}
{lang.msg(language, 12, 'MainWindow')}: {person['postal_code']}
{lang.msg(language, 13, 'MainWindow')}: {person['email']}
{lang.msg(language, 14, 'MainWindow')}: {person['card_number']}
{lang.msg(language, 29, 'MainWindow')}: {date_of_membership}\n\n------------------------------\n\n""")       

    # *-*-* Funzione pulizia campi *-*-*
    
    def clear_box(self):
        # Controllo della funzione di ricerca automatica
        
        auto_search = self.CHB_auto_search.isChecked()
        if auto_search == True: self.CHB_auto_search.setChecked(False)
        
        # Pulizia campi
        
        self.LE_tax_id_code.clear()
        self.LE_name.clear()
        self.LE_surname.clear()
        self.LE_date_of_birth.clear()
        self.LE_birth_place.clear()
        self.CB_sex.setCurrentIndex(0)
        self.LE_city_of_residence.clear()
        self.LE_residential_address.clear()
        self.LE_postal_code.clear()
        self.LE_email.clear()
        self.LE_card_number.clear()
        
        # Riattivazione autoricerca
        
        if auto_search == True: self.CHB_auto_search.setChecked(True)
    
    # *-*-* Funzione apertura esportazione in excel *-*-*
    
    def excel_export(self):
        self.excel_window = ExcelWindow()
        self.excel_window.show()
    
    # *-*-* Funzione apertura finestra database *-*-*
    
    def database_management_open(self):
        self.database_window = DatabaseWindow()
        self.database_window.show()
        
    # *-*-* Funzione apertura finestra opzioni *-*-*
    
    def options_menu_open(self):
        self.options_window = OptionsMenu()
        self.options_window.show()
        self.close()
        
    # *-*-* Funzione controllo numeri *-*-*
    
    def has_numbers(self, st:str):
        return any(char.isdigit() for char in st)
    
    # *-*-* Funzione pressione tasti *-*-*
    
    def keyPressEvent(self, event):
        if event.key() == 16777268: self.clear_box()
        
# -*-* Finestra Database *-*-

class DatabaseWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.db = dbclient["Memberships"] # Apertura database
        
        # *-*-* Impostazioni iniziali *-*-*
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"Database {heading}")
        self.setMinimumSize(640, 540) # Risoluzione minima per schermi piccoli
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        
        # *-*-* Grafica dei widgets *-*-*
        
        self.setStyleSheet(sis.interface_style(interface))
        
        # *-*-* Widgets *-*-*
        
        # Combobox chiave esportazione
        
        self.CB_database_key = QComboBox(self)
        self.CB_database_key.addItems([lang.msg(language, 3, "MainWindow"), lang.msg(language, 4, "MainWindow"), lang.msg(language, 5, "MainWindow"),
                                        lang.msg(language, 6, "MainWindow"), lang.msg(language, 7, "MainWindow"), lang.msg(language, 28, "MainWindow"),
                                        lang.msg(language, 10, "MainWindow"), lang.msg(language, 11, "MainWindow"), lang.msg(language, 12, "MainWindow"),
                                        lang.msg(language, 13, "MainWindow"), lang.msg(language, 14, "MainWindow"), lang.msg(language, 29, "MainWindow"),
                                        lang.msg(language, 0, "DatabaseWindow")])
        self.CB_database_key.currentTextChanged.connect(self.database_key_change)
        self.lay.addWidget(self.CB_database_key, 0, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Combobox esportazione
        
        self.CB_database = QComboBox(self)
        self.CB_database.currentTextChanged.connect(self.database_change)
        self.lay.addWidget(self.CB_database, 0, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Casella inserimento
        
        self.LE_database = QLineEdit(self)
        self.lay.addWidget(self.LE_database, 0, 2, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante cerca
        
        self.B_search = QPushButton(self, text=lang.msg(language, 16, "MainWindow"))
        self.B_search.clicked.connect(self.search_database)
        self.lay.addWidget(self.B_search, 1, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante elimina
        
        self.B_delete = QPushButton(self, text=lang.msg(language, 1, "DatabaseWindow"))
        self.B_delete.clicked.connect(self.delete_database)
        self.lay.addWidget(self.B_delete, 1, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Tabella risultati
        
        self.T_results = QTableWidget(self)
        self.T_results.setColumnCount(5)
        self.T_results.setHorizontalHeaderLabels([lang.msg(language, 3, "MainWindow"), lang.msg(language, 4, "MainWindow"), lang.msg(language, 5, "MainWindow"),
                                                  lang.msg(language, 14, "MainWindow"), lang.msg(language, 29, "MainWindow")])
        self.T_results.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_results.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_results.doubleClicked.connect(self.show_detailed_person)
        self.T_results.customContextMenuRequested.connect(self.T_results_CM)
        self.T_results_headers = self.T_results.horizontalHeader()
        self.lay.addWidget(self.T_results, 2, 0, 1, 3, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.database_key_change() # Avvio della funzione di riempimento Combobox
    
     # -*-* Funzione cambio Combobox chiave per database
    
    def database_key_change(self):
        self.LE_database.clear()
        if self.CB_database_key.currentText() == lang.msg(language, 3, "MainWindow") or self.CB_database_key.currentText() == lang.msg(language, 4, "MainWindow") or self.CB_database_key.currentText() == lang.msg(language, 5, "MainWindow")\
            or self.CB_database_key.currentText() == lang.msg(language, 7, "MainWindow") or self.CB_database_key.currentText() == lang.msg(language, 10, "MainWindow") or self.CB_database_key.currentText() == lang.msg(language, 11, "MainWindow")\
            or self.CB_database_key.currentText() == lang.msg(language, 12, "MainWindow") or self.CB_database_key.currentText() == lang.msg(language, 13, "MainWindow"):
            self.CB_database.clear()
            self.CB_database.addItems([lang.msg(language, 2, "DatabaseWindow"), lang.msg(language, 3, "DatabaseWindow")])
        if self.CB_database_key.currentText() == lang.msg(language, 29, "MainWindow") or self.CB_database_key.currentText() == lang.msg(language, 6, "MainWindow") or self.CB_database_key.currentText() == lang.msg(language, 14, "MainWindow"):
            self.CB_database.clear()
            self.CB_database.addItems([lang.msg(language, 2, "DatabaseWindow"), lang.msg(language, 4, "DatabaseWindow"), lang.msg(language, 5, "DatabaseWindow")])
        if self.CB_database_key.currentText() == lang.msg(language, 28, "MainWindow"):
            self.CB_database.clear()
            self.CB_database.addItems([lang.msg(language, 8, "MainWindow"), lang.msg(language, 9, "MainWindow")])
        if self.CB_database_key.currentText() == lang.msg(language, 0, "DatabaseWindow"):
            self.CB_database.clear()
    
    # -*-* Funzione cambio Combobox database
    
    def database_change(self):
        self.LE_database.clear()
        if self.CB_database_key.currentText() == lang.msg(language, 3, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 6, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 7, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 4, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 8, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 9, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 5, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 10, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 11, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 6, "MainWindow"):
            self.LE_database.setPlaceholderText(lang.msg(language, 12, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 7, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 13, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 14, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 28, "MainWindow"):
            self.LE_database.setPlaceholderText(lang.msg(language, 15, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 10, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 16, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 17, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 11, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 18, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 19, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 12, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 20, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 21, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 13, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 22, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 23, "DatabaseWindow"))
        
        if self.CB_database_key.currentText() == lang.msg(language, 14, "MainWindow"):
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 24, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 25, "DatabaseWindow"))
            if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"):
                self.LE_database.setPlaceholderText(lang.msg(language, 26, "DatabaseWindow"))
            
        if self.CB_database_key.currentText() == lang.msg(language, 29, "MainWindow"):
            date = datetime.now()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                date = date.strftime("%d/%m/%Y")
                self.LE_database.setText(date)
            if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"):
                date -= timedelta(days=365)
                date = date.strftime("%d/%m/%Y")
                self.LE_database.setText(date)
            if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"):
                date = date.strftime("%d/%m/%Y")
                self.LE_database.setText(date)
        
        if self.CB_database_key.currentText() == lang.msg(language, 0, "DatabaseWindow"):
            self.LE_database.setPlaceholderText(lang.msg(language, 15, "DatabaseWindow"))
    
    # -*-* Funzione di ricerca nel database *-*-
    
    def search_database(self):
        
        # Pulizia tabella
        if self.T_results.rowCount() != 0:
            for row in reversed(range(self.T_results.rowCount())):
                self.T_results.removeRow(row)
        
        # Interrogazione database
        col = self.db["cards"]
        
        if self.CB_database_key.currentText() == lang.msg(language, 3, "MainWindow"): # Tramite codice fiscale
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow") and len(self.LE_database.text()) != 16:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 22, "MainWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"tax_id_code": self.LE_database.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"tax_id_code": {"$regex": self.LE_database.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 4, "MainWindow"): # Tramite nome
            if self.has_numbers(self.LE_database.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 23, "MainWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"name": self.LE_database.text().upper().strip()}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"name": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 5, "MainWindow"): # Tramite cognome
            if self.has_numbers(self.LE_database.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 24, "MainWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"surname": self.LE_database.text().upper().strip()}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"surname": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 6, "MainWindow"): # Tramite data di nascita
            date = self.LE_database.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
            
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"date_of_birth": date}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"date_of_birth": {"$gte": date}}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"date_of_birth": {"$lte": date}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 7, "MainWindow"): # Tramite luogo di nascita
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"birth_place": self.LE_database.text().upper().strip()}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"birth_place": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 28, "MainWindow"): # Tramite sesso
            if self.CB_database.currentText() == lang.msg(language, 8, "MainWindow"): self.query_db = col.find({"sex": lang.msg(language, 8, "MainWindow")}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 9, "MainWindow"): self.query_db = col.find({"sex": lang.msg(language, 9, "MainWindow")}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 10, "MainWindow"): # Tramite città di residenza
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"city_of_residence": self.LE_database.text().upper().strip()}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"city_of_residence": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 11, "MainWindow"): # Tramite indirizzo di residenza
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"residential_address": self.LE_database.text().upper().strip()}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"residential_address": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 12, "MainWindow"): # Tramite CAP
            try: int(self.LE_database.text().replace(" ", ""))
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 28, "DatabaseWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"postal_code": self.LE_database.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"postal_code": {"$regex": self.LE_database.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 13, "MainWindow"): # Tramite e-mail
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"email": self.LE_database.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"email": {"$regex": self.LE_database.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 14, "MainWindow"): # Tramite numero tessera
            card_number = self.LE_database.text().replace(" ", "")
            numeric_card_number = True
            try: float(card_number)
            except: numeric_card_number = False
            if numeric_card_number == True:
                number_expression = re.compile(r"^\d+$")
                card_number = float(card_number)
                if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$regex": number_expression}, "$expr":{"$eq": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0})
                if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$regex": number_expression}, "$expr":{"$gte": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0})
                if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$regex": number_expression}, "$expr":{"$lte": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0})
            else:
                if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"card_number": card_number}, {"_id": 0})
                if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$ne": "-", "$gte": card_number}}, {"_id": 0})
                if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$ne": "-", "$lte": card_number}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 29, "MainWindow"): # Tramite data tessera
            date = self.LE_database.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
            date = date.split("/")
            date = f"{date[2]}{date[1]}{date[0]}"
            try: date = int(date)
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
            
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"date_of_membership": {"$ne": "-"}, "$expr":{"$eq": [{"$toDouble": "$date_of_membership"}, date]}}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"date_of_membership": {"$ne": "-"}, "$expr":{"$gte": [{"$toDouble": "$date_of_membership"}, date]}}, {"_id": 0})
            if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"date_of_membership": {"$ne": "-"}, "$expr":{"$lte": [{"$toDouble": "$date_of_membership"}, date]}}, {"_id": 0})
        
        if self.CB_database_key.currentText() == lang.msg(language, 0, "DatabaseWindow"): # Tutti i non tesserati
            self.query_db = col.find({"date_of_membership": "-"}, {"_id": 0})
            
        # Inserimento nella tabella
        
        for person in self.query_db:
            row = self.T_results.rowCount()
            self.T_results.insertRow(row)
            self.T_results.setItem(row, 0, QTableWidgetItem(person["tax_id_code"]))
            self.T_results.setItem(row, 1, QTableWidgetItem(person["name"]))
            self.T_results.setItem(row, 2, QTableWidgetItem(person["surname"]))
            self.T_results.setItem(row, 3, QTableWidgetItem(person["card_number"]))
            date = person["date_of_membership"] # Trasformazione data
            date = f"{date[6:]}/{date[4:6]}/{date[:4]}"
            self.T_results.setItem(row, 4, QTableWidgetItem(date))
    
     # -*-* Funzione di eliminazione nel database *-*-
    
    def delete_database(self):
        
        # Preparazione query per il database
        
        if self.CB_database_key.currentText() == lang.msg(language, 3, "MainWindow"): # Tramite codice fiscale
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow") and len(self.LE_database.text()) != 16:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 22, "MainWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"tax_id_code": self.LE_database.text().upper().replace(" ", "")}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"tax_id_code": {"$regex": self.LE_database.text().upper().replace(" ", "")}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 4, "MainWindow"): # Tramite nome
            if self.has_numbers(self.LE_database.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 23, "MainWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"name": self.LE_database.text().upper().strip()}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"name": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 5, "MainWindow"): # Tramite cognome
            if self.has_numbers(self.LE_database.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 24, "MainWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"surname": self.LE_database.text().upper().strip()}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"surname": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 6, "MainWindow"): # Tramite data di nascita
            date = self.LE_database.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
            
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"date_of_birth": date}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = {"date_of_birth": {"$gte": date}}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = {"date_of_birth": {"$lte": date}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 7, "MainWindow"): # Tramite luogo di nascita
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"birth_place": self.LE_database.text().upper().strip()}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"birth_place": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 28, "MainWindow"): # Tramite sesso
            if self.CB_database.currentText() == lang.msg(language, 8, "MainWindow"): self.query_db = {"sex": lang.msg(language, 8, "MainWindow")}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 9, "MainWindow"): self.query_db = {"sex": lang.msg(language, 9, "MainWindow")}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 10, "MainWindow"): # Tramite città di residenza
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"city_of_residence": self.LE_database.text().upper().strip()}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"city_of_residence": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 11, "MainWindow"): # Tramite indirizzo di residenza
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"residential_address": self.LE_database.text().upper().strip()}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"residential_address": {"$regex": self.LE_database.text().upper().strip()}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 12, "MainWindow"): # Tramite CAP
            try: int(self.LE_database.text().replace(" ", ""))
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 28, "DatabaseWindow"))
                return err_msg.exec()
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"postal_code": self.LE_database.text().upper().replace(" ", "")}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"postal_code": {"$regex": self.LE_database.text().upper().replace(" ", "")}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 13, "MainWindow"): # Tramite e-mail
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"email": self.LE_database.text().upper().replace(" ", "")}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = {"email": {"$regex": self.LE_database.text().upper().replace(" ", "")}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 14, "MainWindow"): # Tramite numero tessera
            numeric_card_number = True
            try: float(card_number)
            except: numeric_card_number = False
            if numeric_card_number == True:
                number_expression = re.compile(r"^\d+$")
                card_number = float(card_number)
                if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"card_number": {"$regex": number_expression}, "$expr":{"$eq": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0}
                if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = {"card_number": {"$regex": number_expression}, "$expr":{"$gte": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0}
                if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = {"card_number": {"$regex": number_expression}, "$expr":{"$lte": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0}
            else:
                if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"card_number": card_number}, {"_id": 0}
                if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = {"card_number": {"$ne": "-", "$gte": card_number}}, {"_id": 0}
                if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = {"card_number": {"$ne": "-", "$lte": card_number}}, {"_id": 0}
        
        if self.CB_database_key.currentText() == lang.msg(language, 29, "MainWindow"): # Tramite data tessera
            date = self.LE_database.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
            date = date.split("/")
            date = f"{date[2]}{date[1]}{date[0]}"
            
            if self.CB_database.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = {"date_of_membership": date}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = {"date_of_membership": {"$gte": date, "$ne": "-"}}, {"_id": 0}
            if self.CB_database.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = {"date_of_membership": {"$lte": date, "$ne": "-"}}, {"_id": 0}
            
        if self.CB_database_key.currentText() == lang.msg(language, 0, "DatabaseWindow"): # Tutti i non tesserati
            self.query_db = {"date_of_membership": "-"}, {"_id": 0}
        
        # Messagebox per eliminazione definitiva
        
        msg = QMessageBox(self)
        msg.setWindowTitle(lang.msg(language, 25, "MainWindow"))
        msg.setText(lang.msg(language, 30, "DatabaseWindow"))
        msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.No)
        msg.buttonClicked.connect(self.delete_database_confirm)
        return msg.exec()
    
    # -*-* Funzione di conferma eliminazione database *-*-
    
    def delete_database_confirm(self, button):
        if button.text() == "&OK" or button.text() == "OK":
            col = self.db["cards"] # Connessione alla collezione del database
            col.delete_many(self.query_db[0])
            
            # Pulizia tabella
            if self.T_results.rowCount() != 0:
                for row in reversed(range(self.T_results.rowCount())):
                    self.T_results.removeRow(row)
    
    # -*-* Custom menu tabella *-*-
    
    def T_results_CM(self):
        if self.T_results.currentRow() == -1: return
        menu = QMenu(self)
        detailed_show_action = QAction(lang.msg(language, 31, "DatabaseWindow"), self)
        delete_action = QAction(lang.msg(language, 1, "DatabaseWindow"), self)
        delete_action.triggered.connect(self.delete_person)
        detailed_show_action.triggered.connect(self.show_detailed_person)
        menu.addActions([detailed_show_action, delete_action])
        menu.popup(QCursor.pos())
    
    # Funzione elimina
    
    def delete_person(self):
        if self.T_results.currentRow() == -1: return
        # Ricerca della persona alla pressione del tasto
        
        row = self.T_results.currentRow()       
        tax_id_code = self.T_results.item(row, 0).text()
        name = self.T_results.item(row, 1).text()
        surname = self.T_results.item(row, 2).text()
        
        col = self.db["cards"] # Connessione alla collezione del database
        col.delete_one({"tax_id_code": tax_id_code, "name": name, "surname": surname})
        
        # Eliminazione dalla tabella
        
        self.T_results.removeRow(row)
        self.T_results.setCurrentCell(-1, -1)
    
    # Funzione visualizzazione dettagliata
    
    def show_detailed_person(self):
        if self.T_results.currentRow() == -1: return
        # Ricerca della persona alla pressione del tasto
        
        row = self.T_results.currentRow()       
        tax_id_code = self.T_results.item(row, 0).text()
        name = self.T_results.item(row, 1).text()
        surname = self.T_results.item(row, 2).text()
        
        col = self.db["cards"] # Connessione alla collezione del database
        person = col.find_one({"tax_id_code": tax_id_code, "name": name, "surname": surname})
        
        # Visualizzazione dettagliata della persona
        
        date_of_membership_st = "-"
        if person["date_of_membership"] != "-":
            date_of_membership_st = f"{person['date_of_membership'][6:]}/{person['date_of_membership'][4:6]}/{person['date_of_membership'][:4]}"
        detailed_person_st = f"""{lang.msg(language, 3, "MainWindow")}: {person["tax_id_code"]}
{lang.msg(language, 4, "MainWindow")}: {person["name"]}
{lang.msg(language, 5, "MainWindow")}: {person["surname"]}
{lang.msg(language, 6, "MainWindow")}: {person["date_of_birth"]}
{lang.msg(language, 7, "MainWindow")}: {person["birth_place"]}
{lang.msg(language, 28, "MainWindow")}: {person["sex"]}
{lang.msg(language, 10, "MainWindow")}: {person["city_of_residence"]}
{lang.msg(language, 11, "MainWindow")}: {person["residential_address"]}
{lang.msg(language, 12, "MainWindow")}: {person["postal_code"]}
{lang.msg(language, 13, "MainWindow")}: {person["email"]}
{lang.msg(language, 14, "MainWindow")}: {person["card_number"]}
{lang.msg(language, 29, "MainWindow")}: {date_of_membership_st}"""
        
        msg = QMessageBox(self)
        msg.setWindowTitle(lang.msg(language, 32, "DatabaseWindow"))
        msg.setText(detailed_person_st)       
        return msg.exec()
    
    # -*-* Funzione di ridimensionamento finestra *-*-
    
    def resizeEvent(self, event):
        W_width = self.width()
        W_height = self.height()
        
        try:
            self.LE_database.setMinimumWidth(int(W_width / 2) - 60)
            self.T_results.setMinimumSize(int(W_width - 15), int(W_height - 130))
            self.T_results_headers.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            self.T_results_headers.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            self.T_results_headers.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            self.T_results_headers.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
            self.T_results_headers.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        except AttributeError: pass
    
    # *-*-* Funzione controllo numeri *-*-*
    
    def has_numbers(self, st:str):
        return any(char.isdigit() for char in st)

# -*-* Finestra Excel *-*-

class ExcelWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.db = dbclient["Memberships"] # Apertura database
        
        # *-*-* Impostazioni iniziali *-*-*
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"{lang.msg(language, 0, 'ExcelWindow')} {heading}")
        self.setMinimumSize(640, 540) # Risoluzione minima per schermi piccoli
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        
        # *-*-* Grafica dei widgets *-*-*
        
        self.setStyleSheet(sis.interface_style(interface))
        
        # *-*-* Widgets *-*-*
        
        # Label numero colonne
        
        L_columns = QLabel(self, text=lang.msg(language, 1, "ExcelWindow"))
        self.lay.addWidget(L_columns, 0, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Spinbox numero colonne
        
        self.SB_colums = QSpinBox(self)
        self.SB_colums.setMinimum(0)
        self.SB_colums.valueChanged.connect(self.set_table_columns)
        self.lay.addWidget(self.SB_colums, 0, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Combobox template
        
        self.CB_template = QComboBox(self)
        self.CB_template.addItems([lang.msg(language, 2, "ExcelWindow")])
        self.lay.addWidget(self.CB_template, 0, 2, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante scelta template
        
        self.B_template = QPushButton(self, text=lang.msg(language, 3, "ExcelWindow"))
        self.B_template.clicked.connect(self.set_template)
        self.lay.addWidget(self.B_template, 0, 3, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante aggiungi riga
        
        self.B_add_row = QPushButton(self, text=lang.msg(language, 4, "ExcelWindow"))
        self.B_add_row.clicked.connect(self.add_row)
        self.lay.addWidget(self.B_add_row, 1, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante rimuovi riga
        
        self.B_remove_row = QPushButton(self, text=lang.msg(language, 5, "ExcelWindow"))
        self.B_remove_row.clicked.connect(self.remove_row)
        self.lay.addWidget(self.B_remove_row, 1, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Tabella di anteprima
        
        self.T_preview = QTableWidget(self)
        self.set_table_columns()
        self.T_preview_headers = self.T_preview.horizontalHeader()
        self.T_preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_preview.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_preview.customContextMenuRequested.connect(self.T_preview_CM)
        self.lay.addWidget(self.T_preview, 2, 0, 1, 4, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Combobox chiave esportazione
        
        self.CB_export_key = QComboBox(self)
        self.CB_export_key.addItems([lang.msg(language, 3, "MainWindow"), lang.msg(language, 4, "MainWindow"), lang.msg(language, 5, "MainWindow"),
                                     lang.msg(language, 6, "MainWindow"), lang.msg(language, 7, "MainWindow"), lang.msg(language, 28, "MainWindow"),
                                     lang.msg(language, 10, "MainWindow"), lang.msg(language, 11, "MainWindow"), lang.msg(language, 12, "MainWindow"),
                                     lang.msg(language, 13, "MainWindow"), lang.msg(language, 14, "MainWindow"), lang.msg(language, 29, "MainWindow")])
        self.CB_export_key.currentTextChanged.connect(self.export_key_change)
        self.lay.addWidget(self.CB_export_key, 3, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Combobox esportazione
        
        self.CB_export = QComboBox(self)
        self.CB_export.currentTextChanged.connect(self.export_change)
        self.lay.addWidget(self.CB_export, 3, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Casella inserimento
        
        self.LE_export = QLineEdit(self)
        self.lay.addWidget(self.LE_export, 3, 2, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante esportazione
        
        self.B_export = QPushButton(self, text=lang.msg(language, 6, "ExcelWindow"))
        self.B_export.clicked.connect(self.export_button)
        self.lay.addWidget(self.B_export, 3, 3, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        
        self.export_key_change() # Avvio della funzione di riempimento Combobox
        
    # -*-* Funzione impostazione colonne *-*-
    
    def set_table_columns(self):
        if self.SB_colums.value() == 0:
            for row in reversed(range(self.T_preview.rowCount())):
                self.T_preview.removeRow(row)
        self.T_preview.setColumnCount(self.SB_colums.value())
        for column in range(self.SB_colums.value()):
            self.T_preview_headers.setSectionResizeMode(column, QHeaderView.ResizeMode.Stretch)
    
    # -*-* Funzione aggiungi riga *-*-
    
    def add_row(self):
        if self.SB_colums.value() == 0:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 7, "ExcelWindow"))
            return err_msg.exec()
        self.T_preview.insertRow(self.T_preview.rowCount())
    
    # -*-* Funzione rimuovi riga *-*-
    
    def remove_row(self):
        if self.T_preview.rowCount() == 0: return # Blocco della funzione se le righe sono 0
        if self.T_preview.currentRow() == -1: # Blocco della funzione se non è stata selezionata un riga
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 8, "ExcelWindow"))
            return err_msg.exec()
        if self.T_preview.currentRow() == self.T_preview.rowCount() -1: self.B_add_row.setDisabled(False)
        self.T_preview.removeRow(self.T_preview.currentRow())
        self.T_preview.setCurrentCell(-1, -1)
    
    # -*-* Funzioni menu della tabella *-*-
    
    def T_preview_CM(self):
         if self.T_preview.currentRow() == -1: return # Blocco della funzione se non è stata selezionata un riga
         menu = QMenu(self)
         personal_insert = QAction(lang.msg(language, 9, "ExcelWindow"), self)
         tax_id_code_insert = QAction(lang.msg(language, 3, "MainWindow"), self)
         name_insert = QAction(lang.msg(language, 4, "MainWindow"), self)
         surname_insert = QAction(lang.msg(language, 5, "MainWindow"), self)
         date_of_birth_insert = QAction(lang.msg(language, 6, "MainWindow"), self)
         birth_place_insert = QAction(lang.msg(language, 7, "MainWindow"), self)
         sex_insert = QAction(lang.msg(language, 28, "MainWindow"), self)
         city_of_residence_insert = QAction(lang.msg(language, 10, "MainWindow"), self)
         residential_address_insert = QAction(lang.msg(language, 11, "MainWindow"), self)
         postal_code_insert = QAction(lang.msg(language, 12, "MainWindow"), self)
         email_insert = QAction(lang.msg(language, 13, "MainWindow"), self)
         card_number_insert = QAction(lang.msg(language, 14, "MainWindow"), self)
         date_of_membership_insert = QAction(lang.msg(language, 29, "MainWindow"), self)
         
         personal_insert.triggered.connect(self.personal_insert)
         tax_id_code_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 3, "MainWindow")))
         name_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 4, "MainWindow")))
         surname_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 5, "MainWindow")))
         date_of_birth_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 6, "MainWindow")))
         birth_place_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 7, "MainWindow")))
         sex_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 28, "MainWindow")))
         city_of_residence_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 10, "MainWindow")))
         residential_address_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 11, "MainWindow")))
         postal_code_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 12, "MainWindow")))
         email_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 13, "MainWindow")))
         card_number_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 14, "MainWindow")))
         date_of_membership_insert.triggered.connect(lambda: self.database_line_insert(lang.msg(language, 29, "MainWindow")))
         
         menu.addActions([personal_insert, tax_id_code_insert, name_insert, surname_insert, date_of_birth_insert, birth_place_insert,sex_insert,
                          city_of_residence_insert, residential_address_insert,postal_code_insert, email_insert, card_number_insert, date_of_membership_insert])
         menu.popup(QCursor.pos())
    
    # Funzione personalizza
    
    def personal_insert(self):
        msg = QInputDialog(self)
        msg.setWindowTitle(lang.msg(language, 10, "ExcelWindow"))
        msg.setLabelText(lang.msg(language, 11, "ExcelWindow"))
        msg.setOkButtonText(lang.msg(language, 15, "MainWindow"))
        msg.setCancelButtonText(lang.msg(language, 12, "ExcelWindow"))
        if msg.exec() == 1:
            text = msg.textValue()
            if "DB:" in text:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 13, "ExcelWindow"))
                return err_msg.exec()
            self.T_preview.setItem(self.T_preview.currentRow(), self.T_preview.currentColumn(), QTableWidgetItem(text))
    
    # Funzione inserimento non personalizzato
    
    def database_line_insert(self, line:str):
        if self.T_preview.rowCount() -1 != self.T_preview.currentRow(): # Errore se non è l'ultima riga
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 14, "ExcelWindow"))
            return err_msg.exec()
        self.T_preview.setItem(self.T_preview.currentRow(), self.T_preview.currentColumn(), QTableWidgetItem(f"DB:{line}"))
        self.B_add_row.setDisabled(True)
    
    # -*-* Funzione inserimento template *-*-
    
    def set_template(self):
        if self.T_preview.rowCount() != 0: # Rimozione vecchie righe
            for row in reversed(range(self.T_preview.rowCount())):
                self.T_preview.removeRow(row)
        if self.CB_template.currentText() == lang.msg(language, 2, "ExcelWindow"):
            self.SB_colums.setValue(13) # Impostazione colonne per template ASI
            self.B_add_row.setDisabled(True) # Disattivazione tasto aggiunta riga
            
            # Impostazione colonne e righe
            self.T_preview.insertRow(0)
            self.T_preview.insertRow(1)
            
            self.T_preview.setItem(0, 0, QTableWidgetItem(lang.msg(language, 15, "ExcelWindow")))
            self.T_preview.setItem(0, 1, QTableWidgetItem(lang.msg(language, 16, "ExcelWindow")))
            self.T_preview.setItem(0, 2, QTableWidgetItem(lang.msg(language, 17, "ExcelWindow")))
            self.T_preview.setItem(0, 3, QTableWidgetItem(lang.msg(language, 18, "ExcelWindow")))
            self.T_preview.setItem(0, 4, QTableWidgetItem(lang.msg(language, 4, "MainWindow").upper()))
            self.T_preview.setItem(1, 4, QTableWidgetItem(lang.msg(language, 19, "ExcelWindow")))
            self.T_preview.setItem(0, 5, QTableWidgetItem(lang.msg(language, 5, "MainWindow").upper()))
            self.T_preview.setItem(1, 5, QTableWidgetItem(lang.msg(language, 20, "ExcelWindow")))
            self.T_preview.setItem(0, 6, QTableWidgetItem(lang.msg(language, 3, "MainWindow").upper()))
            self.T_preview.setItem(1, 6, QTableWidgetItem(lang.msg(language, 21, "ExcelWindow")))
            self.T_preview.setItem(0, 7, QTableWidgetItem(lang.msg(language, 10, "MainWindow").upper()))
            self.T_preview.setItem(1, 7, QTableWidgetItem(lang.msg(language, 22, "ExcelWindow")))
            self.T_preview.setItem(0, 8, QTableWidgetItem(lang.msg(language, 11, "MainWindow").upper()))
            self.T_preview.setItem(1, 8, QTableWidgetItem(lang.msg(language, 23, "ExcelWindow")))
            self.T_preview.setItem(0, 9, QTableWidgetItem(lang.msg(language, 12, "MainWindow").upper()))
            self.T_preview.setItem(1, 9, QTableWidgetItem(lang.msg(language, 24, "ExcelWindow")))
            self.T_preview.setItem(0, 10, QTableWidgetItem(lang.msg(language, 13, "MainWindow").upper()))
            self.T_preview.setItem(1, 10, QTableWidgetItem(lang.msg(language, 25, "ExcelWindow")))
            self.T_preview.setItem(0, 11, QTableWidgetItem(lang.msg(language, 14, "MainWindow").upper()))
            self.T_preview.setItem(1, 11, QTableWidgetItem(lang.msg(language, 26, "ExcelWindow")))
            self.T_preview.setItem(0, 12, QTableWidgetItem(lang.msg(language, 27, "ExcelWindow")))
    
    # -*-* Funzione cambio Combobox chiave per export
    
    def export_key_change(self):
        self.LE_export.clear()
        if self.CB_export_key.currentText() == lang.msg(language, 3, "MainWindow") or self.CB_export_key.currentText() == lang.msg(language, 4, "MainWindow") or self.CB_export_key.currentText() == lang.msg(language, 5, "MainWindow")\
            or self.CB_export_key.currentText() == lang.msg(language, 7, "MainWindow") or self.CB_export_key.currentText() == lang.msg(language, 10, "MainWindow") or self.CB_export_key.currentText() == lang.msg(language, 11, "MainWindow")\
            or self.CB_export_key.currentText() == lang.msg(language, 12, "MainWindow") or self.CB_export_key.currentText() == lang.msg(language, 13, "MainWindow"):
            self.CB_export.clear()
            self.CB_export.addItems([lang.msg(language, 2, "DatabaseWindow"), lang.msg(language, 3, "DatabaseWindow")])
        if self.CB_export_key.currentText() == lang.msg(language, 29, "MainWindow") or self.CB_export_key.currentText() == lang.msg(language, 6, "MainWindow") or self.CB_export_key.currentText() == lang.msg(language, 14, "MainWindow"):
            self.CB_export.clear()
            self.CB_export.addItems([lang.msg(language, 2, "DatabaseWindow"), lang.msg(language, 4, "DatabaseWindow"), lang.msg(language, 5, "DatabaseWindow")])
        if self.CB_export_key.currentText() == lang.msg(language, 28, "MainWindow"):
            self.CB_export.clear()
            self.CB_export.addItems([lang.msg(language, 8, "MainWindow"), lang.msg(language, 9, "MainWindow")])
    
    # -*-* Funzione cambio Combobox export
    
    def export_change(self):
        self.LE_export.clear()
        if self.CB_export_key.currentText() == lang.msg(language, 3, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 6, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 7, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 4, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 8, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 9, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 5, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 10, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 11, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 6, "MainWindow"):
            self.LE_export.setPlaceholderText(lang.msg(language, 12, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 7, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 13, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 14, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 28, "MainWindow"):
            self.LE_export.setPlaceholderText(lang.msg(language, 15, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 10, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 16, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 17, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 11, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 18, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 19, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 12, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 20, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 21, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 13, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 22, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 23, "DatabaseWindow"))
        
        if self.CB_export_key.currentText() == lang.msg(language, 14, "MainWindow"):
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 24, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 4, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 25, "DatabaseWindow"))
            if self.CB_export.currentText() == lang.msg(language, 5, "DatabaseWindow"):
                self.LE_export.setPlaceholderText(lang.msg(language, 26, "DatabaseWindow"))
            
        if self.CB_export_key.currentText() == lang.msg(language, 29, "MainWindow"):
            date = datetime.now()
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"):
                date = date.strftime("%d/%m/%Y")
                self.LE_export.setText(date)
            if self.CB_export.currentText() == lang.msg(language, 4, "DatabaseWindow"):
                date -= timedelta(days=365)
                date = date.strftime("%d/%m/%Y")
                self.LE_export.setText(date)
            if self.CB_export.currentText() == lang.msg(language, 5, "DatabaseWindow"):
                date = date.strftime("%d/%m/%Y")
                self.LE_export.setText(date)
    
    # -*-* Funzione pulsante esportazione *-*-
    
    def export_button(self):
        excel_column_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
        # Controlli tabella
        
        if self.T_preview.rowCount() == 0:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 28, "ExcelWindow"))
            return err_msg.exec()
        
        db_check = 0
        for row in range(self.T_preview.rowCount()):
            for column in range(self.T_preview.columnCount()):
                if self.T_preview.item(row, column) != None:
                    if "DB:" in self.T_preview.item(row, column).text():
                        db_check = 1
                        break
            if db_check == 1: break            
        else:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
            err_msg.setText(lang.msg(language, 29, "ExcelWindow"))
            return err_msg.exec()
        
        # Interrogazione database
        col = self.db["cards"]
        
        if self.CB_export_key.currentText() == lang.msg(language, 3, "MainWindow"): # Tramite codice fiscale
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow") and len(self.LE_export.text()) != 16:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 22, "MainWindow"))
                return err_msg.exec()
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"tax_id_code": self.LE_export.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"tax_id_code": {"$regex": self.LE_export.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 4, "MainWindow"): # Tramite nome
            if self.has_numbers(self.LE_export.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 23, "MainWindow"))
                return err_msg.exec()
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"name": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"name": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 5, "MainWindow"): # Tramite cognome
            if self.has_numbers(self.LE_export.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 24, "MainWindow"))
                return err_msg.exec()
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"surname": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"surname": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 6, "MainWindow"): # Tramite data di nascita
            date = self.LE_export.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
            
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"date_of_birth": date}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"date_of_birth": {"$gte": date}}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"date_of_birth": {"$lte": date}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 7, "MainWindow"): # Tramite luogo di nascita
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"birth_place": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"birth_place": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 28, "MainWindow"): # Tramite sesso
            if self.CB_export.currentText() == lang.msg(language, 8, "MainWindow"): self.query_db = col.find({"sex": lang.msg(language, 8, "MainWindow")}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 9, "MainWindow"): self.query_db = col.find({"sex": lang.msg(language, 9, "MainWindow")}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 10, "MainWindow"): # Tramite città di residenza
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"city_of_residence": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"city_of_residence": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 11, "MainWindow"): # Tramite indirizzo di residenza
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"residential_address": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"residential_address": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 12, "MainWindow"): # Tramite CAP
            try: int(self.LE_export.text().replace(" ", ""))
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 28, "DatabaseWindow"))
                return err_msg.exec()
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"postal_code": self.LE_export.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"postal_code": {"$regex": self.LE_export.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 13, "MainWindow"): # Tramite e-mail
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"email": self.LE_export.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 3, "DatabaseWindow"): self.query_db = col.find({"email": {"$regex": self.LE_export.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 14, "MainWindow"): # Tramite numero tessera
            card_number = self.LE_export.text().replace(" ", "")
            numeric_card_number = True
            try: float(card_number)
            except: numeric_card_number = False
            if numeric_card_number == True:
                number_expression = re.compile(r"^\d+$")
                card_number = float(card_number)
                if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$regex": number_expression}, "$expr":{"$eq": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0})
                if self.CB_export.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$regex": number_expression}, "$expr":{"$gte": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0})
                if self.CB_export.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$regex": number_expression}, "$expr":{"$lte": [{"$toDouble":"$card_number"}, card_number]}}, {"_id": 0})
            else:
                if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"card_number": card_number}, {"_id": 0})
                if self.CB_export.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$ne": "-", "$gte": card_number}}, {"_id": 0})
                if self.CB_export.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"card_number": {"$ne": "-", "$lte": card_number}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == lang.msg(language, 29, "MainWindow"): # Tramite data tessera
            date = self.LE_export.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
            date = date.split("/")
            date = f"{date[2]}{date[1]}{date[0]}"
            try: date = int(date)
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 20, "MainWindow"))
                err_msg.setText(lang.msg(language, 27, "DatabaseWindow"))
                return err_msg.exec()
                        
            if self.CB_export.currentText() == lang.msg(language, 2, "DatabaseWindow"): self.query_db = col.find({"date_of_membership": {"$ne": "-"}, "$expr":{"$eq": [{"$toDouble": "$date_of_membership"}, date]}}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 4, "DatabaseWindow"): self.query_db = col.find({"date_of_membership": {"$ne": "-"}, "$expr":{"$gte": [{"$toDouble": "$date_of_membership"}, date]}}, {"_id": 0})
            if self.CB_export.currentText() == lang.msg(language, 5, "DatabaseWindow"): self.query_db = col.find({"date_of_membership": {"$ne": "-"}, "$expr":{"$lte": [{"$toDouble": "$date_of_membership"}, date]}}, {"_id": 0})
            
        # Esportazione in excel                
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        for row in range(self.T_preview.rowCount() - 1): # Scrittura righe senza accesso al database
            for column in range(self.T_preview.columnCount()):
                if self.T_preview.item(row, column) == None: continue
                ws[f"{excel_column_list[column]}{row+1}"] = self.T_preview.item(row, column).text()
        
        row = self.T_preview.rowCount() - 1 # Scrittura ultima riga con accesso al database
        excel_row = self.T_preview.rowCount()
        for person in self.query_db:
            for column in range(self.T_preview.columnCount()):
                if self.T_preview.item(row, column) == None: continue
                if "DB:" not in self.T_preview.item(row, column).text():
                    ws[f"{excel_column_list[column]}{excel_row}"] = self.T_preview.item(row, column).text()
                else:
                    if self.T_preview.item(row, column).text() == lang.msg(language, 21, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["tax_id_code"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 19, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["name"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 20, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["surname"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 30, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["date_of_birth"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 31, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["birth_place"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 32, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["sex"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 22, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["city_of_residence"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 23, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["residential_address"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 24, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["postal_code"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 25, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["email"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 26, "ExcelWindow"):
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["card_number"]
                    if self.T_preview.item(row, column).text() == lang.msg(language, 33, "ExcelWindow"):
                        date = person["date_of_membership"]
                        date = f"{date[6:]}/{date[4:6]}/{date[:4]}"
                        ws[f"{excel_column_list[column]}{excel_row}"] = date
            excel_row += 1
            
        # Salvataggio file excel
        date = datetime.now()
        date = date.strftime("%Y%m%d")
        wb.save(f"{os.path.expanduser('~')}/Memberships/template_{date}.xlsx")
        
        msg = QMessageBox(self)
        msg.setWindowTitle(lang.msg(language, 34, "ExcelWindow"))
        msg.setText(f"{lang.msg(language, 35, 'ExcelWindow')}:\n{os.path.expanduser('~')}/Memberships/template_{date}.xlsx")
        return msg.exec()
    
    # *-*-* Funzione controllo numeri *-*-*
    
    def has_numbers(self, st:str):
        return any(char.isdigit() for char in st)

    # -*-* Funzione di ridimensionamento finestra *-*-
    
    def resizeEvent(self, event):
        W_width = self.width()
        W_height = self.height()
        
        try:
            self.T_preview.setMinimumSize(int(W_width - 25), int(W_height - 150))
            self.LE_export.setMinimumWidth(int(W_width / 2) - 60)
        except AttributeError: pass
        
# -*-* Menu opzioni *-*-

class OptionsMenu(QWidget):
    def __init__(self):
        super().__init__()
        self.mongodb_connection = "mongodb://localhost:27017/"
        self.heading = ""
        self.interface_style = ""
        self.logo_path = ""
        self.icon_path = ""
        self.language = "ENGLISH"
        
        # Lettura file e impostazione variabili
    
        if os.path.exists(f"{os.path.expanduser('~')}/Memberships/options.txt"):
            options_file = open(f"{os.path.expanduser('~')}/Memberships/options.txt", "r")
            self.mongodb_connection =options_file.readline().replace("db_connection=", "").replace("\n", "")
            self.heading = options_file.readline().replace("heading=", "").replace("\n", "")
            self.interface_style = options_file.readline().replace("interface=", "").replace("\n", "")
            self.logo_path = options_file.readline().replace("logo=", "").replace("\n", "")
            self.icon_path = options_file.readline().replace("icon=", "").replace("\n", "")
            self.language = options_file.readline().replace("language=", "").replace("\n", "")
            options_file.close()
        
        self.setWindowIcon(QIcon(self.icon_path))
        self.setWindowTitle(f"{heading} - {lang.msg(self.language, 0, 'OptionsMenuWindow')}")
        self.lay = QVBoxLayout(self)
        self.setStyleSheet(sis.interface_style(self.interface_style))
        
        self.L_title = QLabel(self, text=lang.msg(self.language, 0, "OptionsMenuWindow"))
        self.L_title.setAccessibleName("an_title")
        self.lay.addWidget(self.L_title)
        
        self.L_language = QLabel(self, text=lang.msg(self.language, 1, "OptionsMenuWindow"))
        self.L_language.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_language)
        
        self.CB_language = QComboBox(self)
        self.CB_language.addItems(["ENGLISH", "ITALIANO"])
        self.CB_language.setCurrentText(self.language)
        self.CB_language.currentTextChanged.connect(self.language_change)
        self.lay.addWidget(self.CB_language)
        
        self.L_database_connection = QLabel(self, text=lang.msg(self.language, 2, "OptionsMenuWindow"))
        self.L_database_connection.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_database_connection)
        
        self.L_database_connection_instructions = QLabel(self)
        self.L_database_connection_instructions.setText(lang.msg(self.language, 3, "OptionsMenuWindow"))
        self.lay.addWidget(self.L_database_connection_instructions)
        
        self.LE_database_connection = QLineEdit(self)
        self.LE_database_connection.setPlaceholderText(lang.msg(self.language, 4, "OptionsMenuWindow"))
        self.LE_database_connection.setText(self.mongodb_connection)
        self.lay.addWidget(self.LE_database_connection)
        
        self.L_heading = QLabel(self, text=lang.msg(self.language, 5, "OptionsMenuWindow"))
        self.L_heading.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_heading)
        
        self.L_heading_instructions = QLabel(self)
        self.L_heading_instructions.setText(lang.msg(self.language, 6, "OptionsMenuWindow"))
        self.lay.addWidget(self.L_heading_instructions)
        
        self.LE_heading = QLineEdit(self)
        self.LE_heading.setPlaceholderText(lang.msg(self.language, 5, "OptionsMenuWindow"))
        self.LE_heading.setText(self.heading)
        self.lay.addWidget(self.LE_heading)
        
        self.L_interface_style = QLabel(self, text=lang.msg(self.language, 7, "OptionsMenuWindow"))
        self.L_interface_style.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_interface_style)
        
        self.L_interface_style_instructions = QLabel(self)
        self.L_interface_style_instructions.setText(lang.msg(self.language, 8, "OptionsMenuWindow"))
        self.lay.addWidget(self.L_interface_style_instructions)
        
        self.CB_interface_style = QComboBox(self)
        self.CB_interface_style.addItems(["98 Style", "Tech Style", "Clear Elegant Style", "Dark Elegant Style"])
        self.CB_interface_style.setCurrentText(self.interface_style)
        self.CB_interface_style.currentIndexChanged.connect(self.interface_change)
        self.lay.addWidget(self.CB_interface_style)
        
        self.L_logo = QLabel(self, text=lang.msg(self.language, 9, "OptionsMenuWindow"))
        self.L_logo.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_logo)
        
        self.L_logo_instructions = QLabel(self, text=f"{lang.msg(self.language, 10, 'OptionsMenuWindow')}: {self.logo_path}")
        self.lay.addWidget(self.L_logo_instructions)
        
        self.B_logo = QPushButton(self, text=lang.msg(self.language, 11, "OptionsMenuWindow"))
        self.B_logo.clicked.connect(self.logo_selection)
        self.lay.addWidget(self.B_logo)
        
        self.L_icon = QLabel(self, text=lang.msg(self.language, 12, "OptionsMenuWindow"))
        self.L_icon.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_icon)
        
        self.L_icon_instructions = QLabel(self, text=f"{lang.msg(self.language, 13, 'OptionsMenuWindow')}: {self.icon_path}")
        self.lay.addWidget(self.L_icon_instructions)
        
        self.B_icon = QPushButton(self, text=lang.msg(self.language, 11, "OptionsMenuWindow"))
        self.B_icon.clicked.connect(self.icon_selection)
        self.lay.addWidget(self.B_icon)
        
        self.B_close_and_save = QPushButton(self, text=lang.msg(self.language, 14, "OptionsMenuWindow"))
        self.B_close_and_save.clicked.connect(self.close_and_save)
        self.lay.addWidget(self.B_close_and_save)
        
        self.L_database_connection_st = QLabel(self)
        self.lay.addWidget(self.L_database_connection_st)
        self.L_database_connection_st.hide()
        
        global first_start_application
        if first_start_application == 0:
            first_start_application = 1
            self.B_close_and_save.animateClick()
    
    def interface_change(self):
        self.setStyleSheet(sis.interface_style(self.CB_interface_style.currentText()))
        
    def logo_selection(self):
        logo = QFileDialog()
        logo.setFileMode(QFileDialog.FileMode.AnyFile)
        logo.setNameFilter(f"{lang.msg(self.language, 15, 'OptionsMenuWindow')} (*.png)")
        logo.setViewMode(QFileDialog.ViewMode.List)
        logo_path = QFileDialog.getOpenFileName(logo)
        logo_path = Path(logo_path[0])
        self.logo_path = logo_path
        self.L_logo_instructions.setText(f"{lang.msg(self.language, 10, 'OptionsMenuWindow')}: {self.logo_path}")
    
    def icon_selection(self):
        icon = QFileDialog()
        icon.setFileMode(QFileDialog.FileMode.AnyFile)
        icon.setNameFilter(f"{lang.msg(self.language, 15, 'OptionsMenuWindow')} (*.png)")
        icon.setViewMode(QFileDialog.ViewMode.List)
        icon_path = QFileDialog.getOpenFileName(icon)
        icon_path = Path(icon_path[0])
        self.icon_path = icon_path
        self.L_icon_instructions.setText(f"{lang.msg(self.language, 13, 'OptionsMenuWindow')}: {self.icon_path}")
    
    def language_change(self):
        self.language = self.CB_language.currentText()
        self.setWindowTitle(f"{heading} - {lang.msg(self.language, 0, 'OptionsMenuWindow')}")
        self.L_title.setText(lang.msg(self.language, 0, "OptionsMenuWindow"))
        self.L_language.setText(lang.msg(self.language, 1, "OptionsMenuWindow"))
        self.L_database_connection.setText(lang.msg(self.language, 2, "OptionsMenuWindow"))
        self.L_database_connection_instructions.setText(lang.msg(self.language, 3, "OptionsMenuWindow"))
        self.LE_database_connection.setPlaceholderText(lang.msg(self.language, 4, "OptionsMenuWindow"))
        self.L_heading.setText(lang.msg(self.language, 5, "OptionsMenuWindow"))
        self.L_heading_instructions.setText(lang.msg(self.language, 6, "OptionsMenuWindow"))
        self.LE_heading.setPlaceholderText(lang.msg(self.language, 5, "OptionsMenuWindow"))
        self.L_interface_style.setText(lang.msg(self.language, 7, "OptionsMenuWindow"))
        self.L_interface_style_instructions.setText(lang.msg(self.language, 8, "OptionsMenuWindow"))
        self.L_logo.setText(lang.msg(self.language, 9, "OptionsMenuWindow"))
        self.L_logo_instructions.setText(f"{lang.msg(self.language, 10, 'OptionsMenuWindow')}: {self.logo_path}")
        self.B_logo.setText(lang.msg(self.language, 11, "OptionsMenuWindow"))
        self.L_icon.setText(lang.msg(self.language, 12, "OptionsMenuWindow"))
        self.L_icon_instructions.setText(f"{lang.msg(self.language, 13, 'OptionsMenuWindow')}: {self.icon_path}")
        self.B_icon.setText(lang.msg(self.language, 11, "OptionsMenuWindow"))
        self.B_close_and_save.setText(lang.msg(self.language, 14, "OptionsMenuWindow"))
    
    def close_and_save(self):
        global language
        if self.LE_database_connection.text() == "":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(self.language, 19, "MainWindow"))
            err_msg.setText(lang.msg(self.language, 16, "OptionsMenuWindow"))
            return err_msg.exec()
        
        # Impostazione del messaggio di connessione
        
        self.L_database_connection_st.setStyleSheet("color: #FF7800; font: 18px bold Arial;")
        self.L_database_connection_st.setText(lang.msg(self.language, 17, "OptionsMenuWindow"))
        self.L_database_connection_st.show()
        self.L_database_connection_st.repaint()
        
        # Salvataggio file
        
        options_file = open(f"{os.path.expanduser('~')}/Memberships/options.txt", "w")
        options_file.write(f"db_connection={self.LE_database_connection.text()}\nheading={self.LE_heading.text()}\ninterface={self.CB_interface_style.currentText()}\nlogo={self.logo_path}\nicon={self.icon_path}\nlanguage={self.CB_language.currentText()}")
        options_file.close()
        
        # -*-* Riavvio applicazione *-*-
        
        # Lettura file e impostazione variabili
    
        options_file = open(f"{os.path.expanduser('~')}/Memberships/options.txt", "r")
        try:
            global dbclient
            dbclient = pymongo.MongoClient(options_file.readline().replace("db_connection=", "").replace("\n", ""))
        except:
            options_file.close()
            self.L_database_connection_st.setStyleSheet("color: #8B0B0B; font: 18px bold Arial;")
            self.L_database_connection_st.setText(lang.msg(self.language, 18, "OptionsMenuWindow"))
            self.L_database_connection_st.show()
            return
        global heading
        heading = options_file.readline().replace("heading=", "").replace("\n", "")
        global interface
        interface = options_file.readline().replace("interface=", "").replace("\n", "")
        global logo_path
        logo_path = options_file.readline().replace("logo=", "").replace("\n", "")
        global icon_path
        icon_path = options_file.readline().replace("icon=", "").replace("\n", "")
        language = options_file.readline().replace("language=", "").replace("\n", "")
        options_file.close()
        
        # Test di connessione

        try:
            dbclient.server_info()
            # Avvio se la connessione al database avviene
            
            self.window = MainWindow()
            self.window.show()
            self.close()
        except:
            self.L_database_connection_st.setStyleSheet("color: #8B0B0B; font: 18px bold Arial;")
            self.L_database_connection_st.setText(lang.msg(self.language, 18, "OptionsMenuWindow"))
            self.L_database_connection_st.show()

# -*-* Avvio applicazione *-*-

if os.path.exists(f"{os.path.expanduser('~')}/Memberships") == False: os.mkdir(f"{os.path.expanduser('~')}/Memberships")

if os.path.exists(f"{os.path.expanduser('~')}/Memberships/options.txt") == False: first_start_application = 1
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OptionsMenu()
    window.show()
    app.exec()