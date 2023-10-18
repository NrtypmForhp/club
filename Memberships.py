import pymongo,sys,os,openpyxl
import SoftwareInterfaceStyle as sis
from pathlib import Path
from datetime import datetime, timedelta
from PyQt6.QtWidgets import (QWidget,QApplication,QGridLayout,QVBoxLayout,QLabel,QLineEdit,QPushButton,QComboBox,QTableWidget,QAbstractItemView,
                             QHeaderView,QMessageBox,QTextEdit,QFileDialog,QCheckBox,QSpinBox,QMenu,QInputDialog,QTableWidgetItem)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QPixmap,QAction,QCursor,QIcon

# Versione 1.0

# Variabili globali

heading = ""
dbclient = ""
interface = ""
logo_path = ""
icon_path = ""
first_start_application = 0

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.db = dbclient["98OttaniTessere"] # Apertura database
        
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
        self.lay.addWidget(L_title, 0, 1, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Checkbox per ricerca automatica (Parte superiore centrale)
        
        self.CHB_auto_search = QCheckBox(self, text="Ricerca automatica")
        self.lay.addWidget(self.CHB_auto_search, 0, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
        
        # Pulsante Opzioni (Parte superiore destra)
        
        self.B_options_menu = QPushButton(self, text="Opzioni")
        self.B_options_menu.clicked.connect(self.options_menu_open)
        self.lay.addWidget(self.B_options_menu, 0, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Casella codice fiscale (Parte centrale sinistra)
        
        self.LE_tax_id_code = QLineEdit(self)
        self.LE_tax_id_code.setPlaceholderText("Codice Fiscale")
        self.LE_tax_id_code.textChanged.connect(self.auto_search)
        self.lay.addWidget(self.LE_tax_id_code, 1, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella nome (Parte centrale sinistra)
        
        self.LE_name = QLineEdit(self)
        self.LE_name.setPlaceholderText("Nome")
        self.LE_name.textChanged.connect(self.auto_search)
        self.lay.addWidget(self.LE_name, 2, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella cognome (Parte centrale sinistra)
        
        self.LE_surname = QLineEdit(self)
        self.LE_surname.setPlaceholderText("Cognome")
        self.LE_surname.textChanged.connect(self.auto_search)
        self.lay.addWidget(self.LE_surname, 3, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella data di nascita (Parte centrale sinistra)
        
        self.LE_date_of_birth = QLineEdit(self)
        self.LE_date_of_birth.setPlaceholderText("Data di nascita")
        self.lay.addWidget(self.LE_date_of_birth, 4, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella luogo di nascita (Parte centrale sinistra)
        
        self.LE_birth_place = QLineEdit(self)
        self.LE_birth_place.setPlaceholderText("Luogo di nascita")
        self.lay.addWidget(self.LE_birth_place, 5, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Box sesso (Parte centrale sinistra)
        
        self.CB_sex = QComboBox(self)
        self.CB_sex.addItems(["MASCHIO", "FEMMINA"])
        self.lay.addWidget(self.CB_sex, 6, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella comune di residenza (Parte centrale sinistra)
        
        self.LE_city_of_residence = QLineEdit(self)
        self.LE_city_of_residence.setPlaceholderText("Comune di residenza")
        self.lay.addWidget(self.LE_city_of_residence, 7, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella indirizzo di residenza (Parte centrale sinistra)
        
        self.LE_residential_address = QLineEdit(self)
        self.LE_residential_address.setPlaceholderText("Indirizzo di residenza")
        self.lay.addWidget(self.LE_residential_address, 8, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella CAP (Parte centrale sinistra)
        
        self.LE_postal_code = QLineEdit(self)
        self.LE_postal_code.setPlaceholderText("CAP - Codice di Avviamento Postale")
        self.lay.addWidget(self.LE_postal_code, 9, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella e-mail (Parte centrale sinistra)
        
        self.LE_email = QLineEdit(self)
        self.LE_email.setPlaceholderText("e-mail")
        self.lay.addWidget(self.LE_email, 10, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella numero tessera (Parte centrale sinistra)
        
        self.LE_card_number = QLineEdit(self)
        self.LE_card_number.setPlaceholderText("Numero tessera")
        self.lay.addWidget(self.LE_card_number, 11, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante inserisci (Parte bassa sinistra)
        
        self.B_insert = QPushButton(self, text="Inserisci >>")
        self.B_insert.clicked.connect(self.insert_into_db)
        self.lay.addWidget(self.B_insert, 12, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante cerca (Parte bassa sinistra)
        
        self.B_search = QPushButton(self, text="Cerca")
        self.B_search.clicked.connect(self.search_db)
        self.lay.addWidget(self.B_search, 12, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante esporta in Excel
        
        self.B_excel_export = QPushButton(self, text="Esporta excel")
        self.B_excel_export.clicked.connect(self.excel_export)
        self.lay.addWidget(self.B_excel_export, 12, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Combobox strumenti di controllo e esportazione (Parte bassa sinistra)
        
        self.CB_control_export = QComboBox(self)
        self.CB_control_export.addItems(["Tessere ancora valide", "Tessere non valide", "Non tesserati"])
        self.lay.addWidget(self.CB_control_export, 13, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante interrogazione database (Parte bassa sinistra)
        
        self.B_query_db = QPushButton(self, text="Interroga")
        self.B_query_db.clicked.connect(self.query_db)
        self.lay.addWidget(self.B_query_db, 13, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Box interrogazione database (Parte centrale destra)
        
        self.TE_db_response = QTextEdit(self)
        self.TE_db_response.setReadOnly(True)
        self.TE_db_response.setPlainText(f"""Gestione tesseramenti {heading}\n\n
** Funzioni rapide **
Pulsante F5: Pulisce tutti i campi a sinistra""")
        self.lay.addWidget(self.TE_db_response, 1, 3, 13, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
    
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
            err_msg.setWindowTitle("Errore")
            err_msg.setText("I dati che devono obbligatoriamente essere inseriti sono:\n- Codice fiscale\n- Nome\n- Cognome")
            return err_msg.exec()
        
        # Controllo codice fiscale
        
        if len(self.tax_id_code) != 16:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Lunghezza codice fiscale non corretta!")
            return err_msg.exec()
        if self.has_numbers(self.tax_id_code[0:6]) == True or self.has_numbers(self.tax_id_code[8]) == True or self.has_numbers(self.tax_id_code[11]) == True or self.has_numbers(self.tax_id_code[15]) == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Codice fiscale non corretto!")
            return err_msg.exec()
        if self.tax_id_code[6:8].isdigit() == False or self.tax_id_code[9:11].isdigit() == False or self.tax_id_code[12:15].isdigit() == False:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Codice fiscale non corretto!")
            return err_msg.exec()
        tax_id_code_char_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","0","1","2","3","4","5","6","7","8","9"]
        for char in self.tax_id_code:
            if char not in tax_id_code_char_list:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Codice fiscale non corretto!")
                return err_msg.exec()
        
        # Controllo nome
        
        if self.has_numbers(self.name) == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Nome non corretto!")
            return err_msg.exec()
        
        # Controllo cognome
        
        if self.has_numbers(self.surname) == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Cognome non corretto!")
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
            msg.setWindowTitle("Attenzione")
            msg.setText("Persona già presente nel database!\nVuoi modificare i dati con quelli inseriti?")
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
        self.TE_db_response.insertPlainText(f"""-*-* Inserimento effettuato *-*-
\nCodice fiscale: {self.tax_id_code}
Nome: {self.name}
Cognome: {self.surname}
Data di nascita: {self.date_of_birth}
Luogo di nascita: {self.birth_place}
Sesso: {self.sex}
Comune di residenza: {self.city_of_residence}
Indirizzo di residenza: {self.residential_address}
CAP: {self.postal_code}
email: {self.email}
Numero tessera: {self.card_number}
Data tesseramento: {date_of_membership}""")
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
        self.TE_db_response.insertPlainText(f"""-*-* Modifica effettuata *-*-
\nCodice fiscale: {self.tax_id_code}
Nome: {self.name}
Cognome: {self.surname}
Data di nascita: {self.date_of_birth}
Luogo di nascita: {self.birth_place}
Sesso: {self.sex}
Comune di residenza: {self.city_of_residence}
Indirizzo di residenza: {self.residential_address}
CAP: {self.postal_code}
email: {self.email}
Numero tessera: {self.card_number}
Data tesseramento: {date_of_membership}""")
        self.clear_box()
    
    # *-*-* Funzione ricerca automatica nel database *-*-*
    
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
            self.TE_db_response.insertPlainText("-*-* Persone trovate *-*-\n\n")
            for person in people_founded:
                date_of_membership = person["date_of_membership"] # Trasformazione data in formato leggibile
                if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
                self.TE_db_response.append(f"""Codice fiscale: {person['tax_id_code']}
Nome: {person['name']}
Cognome: {person['surname']}
Data di nascita: {person['date_of_birth']}
Luogo di nascita: {person['birth_place']}
Sesso: {person['sex']}
Comune di residenza: {person['city_of_residence']}
Indirizzo di residenza: {person['residential_address']}
CAP: {person['postal_code']}
email: {person['email']}
Numero tessera: {person['card_number']}
Data tesseramento: {date_of_membership}\n\n------------------------------\n\n""")
    
    # *-*-* Funzione ricerca manuale nel database *-*-*
    
    def search_db(self):
        # Controllo se la ricerca automatica è attiva
        
        if self.CHB_auto_search.isChecked() == True:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Il pulsante non può essere usato con la ricerca automatica attiva!\nPrima disattiva la ricerca automatica.")
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
            err_msg.setWindowTitle("Errore")
            err_msg.setText("""I campi non contengono nulla!\n
Controlla di aver compilato almeno uno dei seguenti campi in questo modo:
- Codice fiscale: completo
- Nome: 3 o più lettere
- Cognome: 3 o più lettere""")
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
            err_msg.setWindowTitle("Attenzione")
            err_msg.setText("Nessun dato trovato nel database!")
            return err_msg.exec()
        else:
            self.TE_db_response.clear()
            self.TE_db_response.insertPlainText("-*-* Persone trovate *-*-\n\n")
            for person in people_founded:
                date_of_membership = person["date_of_membership"] # Trasformazione data in formato leggibile
                if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
                self.TE_db_response.append(f"""Codice fiscale: {person['tax_id_code']}
Nome: {person['name']}
Cognome: {person['surname']}
Data di nascita: {person['date_of_birth']}
Luogo di nascita: {person['birth_place']}
Sesso: {person['sex']}
Comune di residenza: {person['city_of_residence']}
Indirizzo di residenza: {person['residential_address']}
CAP: {person['postal_code']}
email: {person['email']}
Numero tessera: {person['card_number']}
Data tesseramento: {date_of_membership}\n\n------------------------------\n\n""")
    
    # *-*-* Funzione ricerca massiva nel database *-*-*
            
    def query_db(self):
        date = datetime.now()
        col = self.db["cards"]
        if self.CB_control_export.currentText() == "Tessere ancora valide":
            date -= timedelta(days=365)
            date = date.strftime("%Y%m%d")
            self.TE_db_response.clear()
            self.TE_db_response.insertPlainText("-*-* Persone trovate *-*-\n\n")
            for person in col.find({"date_of_membership": {"$gte": date, "$ne": "-"}}, {"_id": 0}):
                date_of_membership = person["date_of_membership"] # Trasformazione data in formato leggibile
                if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
                self.TE_db_response.append(f"""Codice fiscale: {person['tax_id_code']}
Nome: {person['name']}
Cognome: {person['surname']}
Data di nascita: {person['date_of_birth']}
Luogo di nascita: {person['birth_place']}
Sesso: {person['sex']}
Comune di residenza: {person['city_of_residence']}
Indirizzo di residenza: {person['residential_address']}
CAP: {person['postal_code']}
email: {person['email']}
Numero tessera: {person['card_number']}
Data tesseramento: {date_of_membership}\n\n------------------------------\n\n""")
        
        if self.CB_control_export.currentText() == "Tessere non valide":
            date -= timedelta(days=365)
            date = date.strftime("%Y%m%d")
            self.TE_db_response.clear()
            self.TE_db_response.insertPlainText("-*-* Persone trovate *-*-\n\n")
            for person in col.find({"date_of_membership": {"$lte": date, "$ne": "-"}}, {"_id": 0}):
                date_of_membership = person["date_of_membership"] # Trasformazione data in formato leggibile
                if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
                self.TE_db_response.append(f"""Codice fiscale: {person['tax_id_code']}
Nome: {person['name']}
Cognome: {person['surname']}
Data di nascita: {person['date_of_birth']}
Luogo di nascita: {person['birth_place']}
Sesso: {person['sex']}
Comune di residenza: {person['city_of_residence']}
Indirizzo di residenza: {person['residential_address']}
CAP: {person['postal_code']}
email: {person['email']}
Numero tessera: {person['card_number']}
Data tesseramento: {date_of_membership}\n\n------------------------------\n\n""")
        
        if self.CB_control_export.currentText() == "Non tesserati":
            self.TE_db_response.clear()
            self.TE_db_response.insertPlainText("-*-* Persone trovate *-*-\n\n")
            for person in col.find({"date_of_membership": "-", "card_number": "-"}, {"_id": 0}):
                date_of_membership = person["date_of_membership"] # Trasformazione data in formato leggibile
                if date_of_membership != "-": date_of_membership = f"{date_of_membership[6:]}/{date_of_membership[4:6]}/{date_of_membership[:4]}"
                self.TE_db_response.append(f"""Codice fiscale: {person['tax_id_code']}
Nome: {person['name']}
Cognome: {person['surname']}
Data di nascita: {person['date_of_birth']}
Luogo di nascita: {person['birth_place']}
Sesso: {person['sex']}
Comune di residenza: {person['city_of_residence']}
Indirizzo di residenza: {person['residential_address']}
CAP: {person['postal_code']}
email: {person['email']}
Numero tessera: {person['card_number']}
Data tesseramento: {date_of_membership}\n\n------------------------------\n\n""")
        
    
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

# -*-* Finestra Excel *-*-

class ExcelWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.db = dbclient["98OttaniTessere"] # Apertura database
        
        # *-*-* Impostazioni iniziali *-*-*
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"Esportazione excel {heading}")
        self.setMinimumSize(640, 540) # Risoluzione minima per schermi piccoli
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        
        # *-*-* Grafica dei widgets *-*-*
        
        self.setStyleSheet(sis.interface_style(interface))
        
        # *-*-* Widgets *-*-*
        
        # Label numero colonne
        
        L_columns = QLabel(self, text="Colonne:")
        self.lay.addWidget(L_columns, 0, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Spinbox numero colonne
        
        self.SB_colums = QSpinBox(self)
        self.SB_colums.setMinimum(0)
        self.SB_colums.valueChanged.connect(self.set_table_columns)
        self.lay.addWidget(self.SB_colums, 0, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Combobox template
        
        self.CB_template = QComboBox(self)
        self.CB_template.addItems(["Formato ASI"])
        self.lay.addWidget(self.CB_template, 0, 2, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante scelta template
        
        self.B_template = QPushButton(self, text="Imposta")
        self.B_template.clicked.connect(self.set_template)
        self.lay.addWidget(self.B_template, 0, 3, Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante aggiungi riga
        
        self.B_add_row = QPushButton(self, text="Aggiungi riga")
        self.B_add_row.clicked.connect(self.add_row)
        self.lay.addWidget(self.B_add_row, 1, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Pulsante rimuovi riga
        
        self.B_remove_row = QPushButton(self, text="Rimuovi riga")
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
        self.CB_export_key.addItems(["Codice fiscale", "Nome", "Cognome", "Data di nascita", "Luogo di nascita", "Sesso", "Città di residenza",
                                     "Indirizzo di residenza", "CAP", "e-mail", "Numero tessera", "Data tessera"])
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
        
        self.B_export = QPushButton(self, text="Esporta in excel")
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
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Prima di aggiungere una riga deve esserci almeno una colonna!")
            return err_msg.exec()
        self.T_preview.insertRow(self.T_preview.rowCount())
    
    # -*-* Funzione rimuovi riga *-*-
    
    def remove_row(self):
        if self.T_preview.rowCount() == 0: return # Blocco della funzione se le righe sono 0
        if self.T_preview.currentRow() == -1: # Blocco della funzione se non è stata selezionata un riga
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Seleziona una riga cliccandoci sopra!")
            return err_msg.exec()
        if self.T_preview.currentRow() == self.T_preview.rowCount() -1: self.B_add_row.setDisabled(False)
        self.T_preview.removeRow(self.T_preview.currentRow())
        self.T_preview.setCurrentCell(-1, -1)
    
    # -*-* Funzioni menu della tabella *-*-
    
    def T_preview_CM(self):
         if self.T_preview.currentRow() == -1: return # Blocco della funzione se non è stata selezionata un riga
         menu = QMenu(self)
         personal_insert = QAction("Personalizza", self)
         tax_id_code_insert = QAction("Codice fiscale", self)
         name_insert = QAction("Nome", self)
         surname_insert = QAction("Cognome", self)
         date_of_birth_insert = QAction("Data di nascita", self)
         birth_place_insert = QAction("Luogo di nascita", self)
         sex_insert = QAction("Sesso", self)
         city_of_residence_insert = QAction("Città di residenza", self)
         residential_address_insert = QAction("Indirizzo di residenza", self)
         postal_code_insert = QAction("CAP", self)
         email_insert = QAction("e-mail", self)
         card_number_insert = QAction("Numero tessera", self)
         date_of_membership_insert = QAction("Data tesseramento", self)
         
         personal_insert.triggered.connect(self.personal_insert)
         tax_id_code_insert.triggered.connect(lambda: self.database_line_insert("Codice fiscale"))
         name_insert.triggered.connect(lambda: self.database_line_insert("Nome"))
         surname_insert.triggered.connect(lambda: self.database_line_insert("Cognome"))
         date_of_birth_insert.triggered.connect(lambda: self.database_line_insert("Data di nascita"))
         birth_place_insert.triggered.connect(lambda: self.database_line_insert("Luogo di nascita"))
         sex_insert.triggered.connect(lambda: self.database_line_insert("Sesso"))
         city_of_residence_insert.triggered.connect(lambda: self.database_line_insert("Città di residenza"))
         residential_address_insert.triggered.connect(lambda: self.database_line_insert("Indirizzo di residenza"))
         postal_code_insert.triggered.connect(lambda: self.database_line_insert("CAP"))
         email_insert.triggered.connect(lambda: self.database_line_insert("e-mail"))
         card_number_insert.triggered.connect(lambda: self.database_line_insert("Numero tessera"))
         date_of_membership_insert.triggered.connect(lambda: self.database_line_insert("Data di tesseramento"))
         
         menu.addActions([personal_insert, tax_id_code_insert, name_insert, surname_insert, date_of_birth_insert, birth_place_insert,sex_insert,
                          city_of_residence_insert, residential_address_insert,postal_code_insert, email_insert, card_number_insert, date_of_membership_insert])
         menu.popup(QCursor.pos())
    
    # Funzione personalizza
    
    def personal_insert(self):
        msg = QInputDialog(self)
        msg.setWindowTitle("Inserimento personalizzato")
        msg.setLabelText("L'inserimento personalizzato ti permette di\ninserire un valore extra alla cella.\nQuesto valore verrà solamente copiato durante l'esportazione.")
        msg.setOkButtonText("Inserisci")
        msg.setCancelButtonText("Annulla")
        if msg.exec() == 1:
            text = msg.textValue()
            if "DB:" in text:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("La dicitura 'DB:' non può essere inserita manualmente!")
                return err_msg.exec()
            self.T_preview.setItem(self.T_preview.currentRow(), self.T_preview.currentColumn(), QTableWidgetItem(text))
    
    # Funzione inserimento non personalizzato
    
    def database_line_insert(self, line:str):
        if self.T_preview.rowCount() -1 != self.T_preview.currentRow(): # Errore se non è l'ultima riga
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Queste righe a differenza di quella personalizzata iniziano la ricerca nel database!\nDevono necessariamente essere inserite sull'ultima riga")
            return err_msg.exec()
        self.T_preview.setItem(self.T_preview.currentRow(), self.T_preview.currentColumn(), QTableWidgetItem(f"DB:{line}"))
        self.B_add_row.setDisabled(True)
    
    # -*-* Funzione inserimento template *-*-
    
    def set_template(self):
        if self.T_preview.rowCount() != 0: # Rimozione vecchie righe
            for row in reversed(range(self.T_preview.rowCount())):
                self.T_preview.removeRow(row)
        if self.CB_template.currentText() == "Formato ASI":
            self.SB_colums.setValue(13) # Impostazione colonne per template ASI
            self.B_add_row.setDisabled(True) # Disattivazione tasto aggiunta riga
            
            # Impostazione colonne e righe
            self.T_preview.insertRow(0)
            self.T_preview.insertRow(1)
            
            self.T_preview.setItem(0, 0, QTableWidgetItem("STAGIONE"))
            self.T_preview.setItem(0, 1, QTableWidgetItem("DISCIPLINA"))
            self.T_preview.setItem(0, 2, QTableWidgetItem("QUALIFICA"))
            self.T_preview.setItem(0, 3, QTableWidgetItem("TIPO TESSERA"))
            self.T_preview.setItem(0, 4, QTableWidgetItem("NOME"))
            self.T_preview.setItem(1, 4, QTableWidgetItem("DB:Nome"))
            self.T_preview.setItem(0, 5, QTableWidgetItem("COGNOME"))
            self.T_preview.setItem(1, 5, QTableWidgetItem("DB:Cognome"))
            self.T_preview.setItem(0, 6, QTableWidgetItem("CODICE FISCALE"))
            self.T_preview.setItem(1, 6, QTableWidgetItem("DB:Codice fiscale"))
            self.T_preview.setItem(0, 7, QTableWidgetItem("COMUNE RESIDENZA"))
            self.T_preview.setItem(1, 7, QTableWidgetItem("DB:Città di residenza"))
            self.T_preview.setItem(0, 8, QTableWidgetItem("INDIRIZZO RESIDENZA"))
            self.T_preview.setItem(1, 8, QTableWidgetItem("DB:Indirizzo di residenza"))
            self.T_preview.setItem(0, 9, QTableWidgetItem("CAP"))
            self.T_preview.setItem(1, 9, QTableWidgetItem("DB:CAP"))
            self.T_preview.setItem(0, 10, QTableWidgetItem("EMAIL"))
            self.T_preview.setItem(1, 10, QTableWidgetItem("DB:email"))
            self.T_preview.setItem(0, 11, QTableWidgetItem("CODICE TESSERA"))
            self.T_preview.setItem(1, 11, QTableWidgetItem("DB:Numero tessera"))
            self.T_preview.setItem(0, 12, QTableWidgetItem("CODICE AFFILIAZIONE"))
    
    # -*-* Funzione cambio Combobox chiave per export
    
    def export_key_change(self):
        self.LE_export.clear()
        if self.CB_export_key.currentText() == "Codice fiscale" or self.CB_export_key.currentText() == "Nome" or self.CB_export_key.currentText() == "Cognome"\
            or self.CB_export_key.currentText() == "Luogo di nascita" or self.CB_export_key.currentText() == "Città di residenza" or self.CB_export_key.currentText() == "Indirizzo di residenza"\
            or self.CB_export_key.currentText() == "CAP" or self.CB_export_key.currentText() == "e-mail":
            self.CB_export.clear()
            self.CB_export.addItems(["Uguale a", "Contiene"])
        if self.CB_export_key.currentText() == "Data tessera" or self.CB_export_key.currentText() == "Data di nascita" or self.CB_export_key.currentText() == "Numero tessera":
            self.CB_export.clear()
            self.CB_export.addItems(["Uguale a", "Maggiore di", "Minore di"])
        if self.CB_export_key.currentText() == "Sesso":
            self.CB_export.clear()
            self.CB_export.addItems(["MASCHIO", "FEMMINA"])
    
    # -*-* Funzione cambio Combobox export
    
    def export_change(self):
        self.LE_export.clear()
        if self.CB_export_key.currentText() == "Codice fiscale":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("Codice fiscale completo")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte di codice fiscale")
        
        if self.CB_export_key.currentText() == "Nome":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("Nome completo")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte del nome")
        
        if self.CB_export_key.currentText() == "Cognome":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("Cognome completo")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte del cognome")
        
        if self.CB_export_key.currentText() == "Data di nascita":
            self.LE_export.setPlaceholderText("Esempio: 18/09/1980")
        
        if self.CB_export_key.currentText() == "Luogo di nascita":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("Luogo di nascita completo")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte del luogo di nascita")
        
        if self.CB_export_key.currentText() == "Sesso":
            self.LE_export.setPlaceholderText("Casella non necessaria")
        
        if self.CB_export_key.currentText() == "Città di residenza":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("Città di residenza completa")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte della città di residenza")
        
        if self.CB_export_key.currentText() == "Indirizzo di residenza":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("Indirizzo di residenza completo")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte dell'indirizzo di residenza")
        
        if self.CB_export_key.currentText() == "CAP":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("CAP completo")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte del CAP")
        
        if self.CB_export_key.currentText() == "e-mail":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("e-mail completa")
            if self.CB_export.currentText() == "Contiene":
                self.LE_export.setPlaceholderText("Parte della e-mail")
        
        if self.CB_export_key.currentText() == "Numero tessera":
            if self.CB_export.currentText() == "Uguale a":
                self.LE_export.setPlaceholderText("Numero tessera specifico")
            if self.CB_export.currentText() == "Maggiore di":
                self.LE_export.setPlaceholderText("Maggiore del numero inserito")
            if self.CB_export.currentText() == "Minore di":
                self.LE_export.setPlaceholderText("Minore del numero inserito")
            
        if self.CB_export_key.currentText() == "Data tessera":
            date = datetime.now()
            if self.CB_export.currentText() == "Uguale a":
                date = date.strftime("%d/%m/%Y")
                self.LE_export.setText(date)
            if self.CB_export.currentText() == "Maggiore di":
                date -= timedelta(days=365)
                date = date.strftime("%d/%m/%Y")
                self.LE_export.setText(date)
            if self.CB_export.currentText() == "Minore di":
                date = date.strftime("%d/%m/%Y")
                self.LE_export.setText(date)
    
    # -*-* Funzione pulsante esportazione *-*-
    
    def export_button(self):
        excel_column_list = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
        # Controlli tabella
        
        if self.T_preview.rowCount() == 0:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("La tabella non può essere vuota!")
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
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Nella tabella ci deve essere almeno un valore da controllare nel database!")
            return err_msg.exec()
        
        # Interrogazione database
        col = self.db["cards"]
        
        if self.CB_export_key.currentText() == "Codice fiscale": # Tramite codice fiscale
            if self.CB_export.currentText() == "Uguale a" and len(self.LE_export.text()) != 16:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Codice fiscale non corretto")
                return err_msg.exec()
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"tax_id_code": self.LE_export.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"tax_id_code": {"$regex": self.LE_export.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Nome": # Tramite nome
            if self.has_numbers(self.LE_export.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Inserimento non corretto")
                return err_msg.exec()
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"name": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"name": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Cognome": # Tramite cognome
            if self.has_numbers(self.LE_export.text().upper().strip()) == True:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Inserimento non corretto")
                return err_msg.exec()
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"surname": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"surname": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Data di nascita": # Tramite data di nascita
            date = self.LE_export.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Data non corretta")
                return err_msg.exec()
            date = date.split("/")
            date = f"{date[2]}{date[1]}{date[0]}"
            
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"date_of_birth": date}, {"_id": 0})
            if self.CB_export.currentText() == "Maggiore di": self.query_db = col.find({"date_of_birth": {"$gte": date}}, {"_id": 0})
            if self.CB_export.currentText() == "Minore di": self.query_db = col.find({"date_of_birth": {"$lte": date}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Luogo di nascita": # Tramite luogo di nascita
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"birth_place": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"birth_place": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Sesso": # Tramite sesso
            if self.CB_export.currentText() == "MASCHIO": self.query_db = col.find({"sex": "MASCHIO"}, {"_id": 0})
            if self.CB_export.currentText() == "FEMMINA": self.query_db = col.find({"sex": "FEMMINA"}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Città di residenza": # Tramite città di residenza
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"city_of_residence": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"city_of_residence": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Indirizzo di residenza": # Tramite indirizzo di residenza
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"residential_address": self.LE_export.text().upper().strip()}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"residential_address": {"$regex": self.LE_export.text().upper().strip()}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "CAP": # Tramite CAP
            try: int(self.LE_export.text().replace(" ", ""))
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Inserimento non corretto")
                return err_msg.exec()
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"postal_code": self.LE_export.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"postal_code": {"$regex": self.LE_export.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "e-mail": # Tramite e-mail
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"email": self.LE_export.text().upper().replace(" ", "")}, {"_id": 0})
            if self.CB_export.currentText() == "Contiene": self.query_db = col.find({"email": {"$regex": self.LE_export.text().upper().replace(" ", "")}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Numero tessera": # Tramite numero tessera
            try: int(self.LE_export.text().replace(" ", ""))
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Inserimento non corretto")
                return err_msg.exec()
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"card_number": self.LE_export.text().replace(" ", "")}, {"_id": 0})
            if self.CB_export.currentText() == "Maggiore di": self.query_db = col.find({"card_number": {"$gte": self.LE_export.text().replace(" ", ""), "$ne": "-"}}, {"_id": 0})
            if self.CB_export.currentText() == "Minore di": self.query_db = col.find({"card_number": {"$lte": self.LE_export.text().replace(" ", ""), "$ne": "-"}}, {"_id": 0})
        
        if self.CB_export_key.currentText() == "Data tessera": # Tramite data tessera
            date = self.LE_export.text().replace(" ", "")
            if date.count("/") != 2 or len(date) != 10: # Controllo data inserita
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle("Errore")
                err_msg.setText("Data non corretta")
                return err_msg.exec()
            date = date.split("/")
            date = f"{date[2]}{date[1]}{date[0]}"
            
            if self.CB_export.currentText() == "Uguale a": self.query_db = col.find({"date_of_membership": date}, {"_id": 0})
            if self.CB_export.currentText() == "Maggiore di": self.query_db = col.find({"date_of_membership": {"$gte": date, "$ne": "-"}}, {"_id": 0})
            if self.CB_export.currentText() == "Minore di": self.query_db = col.find({"date_of_membership": {"$lte": date, "$ne": "-"}}, {"_id": 0})
            
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
                    if self.T_preview.item(row, column).text() == "DB:Codice fiscale":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["tax_id_code"]
                    if self.T_preview.item(row, column).text() == "DB:Nome":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["name"]
                    if self.T_preview.item(row, column).text() == "DB:Cognome":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["surname"]
                    if self.T_preview.item(row, column).text() == "DB:Data di nascita":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["date_of_birth"]
                    if self.T_preview.item(row, column).text() == "DB:Luogo di nascita":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["birth_place"]
                    if self.T_preview.item(row, column).text() == "DB:Sesso":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["sex"]
                    if self.T_preview.item(row, column).text() == "DB:Città di residenza":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["city_of_residence"]
                    if self.T_preview.item(row, column).text() == "DB:Indirizzo di residenza":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["residential_address"]
                    if self.T_preview.item(row, column).text() == "DB:CAP":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["postal_code"]
                    if self.T_preview.item(row, column).text() == "DB:e-mail":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["email"]
                    if self.T_preview.item(row, column).text() == "DB:Numero tessera":
                        ws[f"{excel_column_list[column]}{excel_row}"] = person["card_number"]
                    if self.T_preview.item(row, column).text() == "DB:Data di tesseramento":
                        date = person["date_of_membership"]
                        date = f"{date[6:]}/{date[4:6]}/{date[:4]}"
                        ws[f"{excel_column_list[column]}{excel_row}"] = date
            excel_row += 1
            
        # Salvataggio file excel
        date = datetime.now()
        date = date.strftime("%Y%m%d")
        wb.save(f"{os.environ['HOME']}/Memberships/template_{date}.xlsx")
        
        msg = QMessageBox(self)
        msg.setWindowTitle("File salvato")
        msg.setText(f"File salvato correttamente:\n{os.environ['HOME']}/Memberships/template_{date}.xlsx")
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
        
        # Lettura file e impostazione variabili
    
        if os.path.exists(f"{os.environ['HOME']}/Memberships/options.txt"):
            options_file = open(f"{os.environ['HOME']}/Memberships/options.txt", "r")
            self.mongodb_connection =options_file.readline().replace("db_connection=", "").replace("\n", "")
            self.heading = options_file.readline().replace("heading=", "").replace("\n", "")
            self.interface_style = options_file.readline().replace("interface=", "").replace("\n", "")
            self.logo_path = options_file.readline().replace("logo=", "").replace("\n", "")
            self.icon_path = options_file.readline().replace("icon=", "").replace("\n", "")
            options_file.close()
        
        self.setWindowIcon(QIcon(self.icon_path))
        self.setWindowTitle(f"{heading} - opzioni")
        self.lay = QVBoxLayout(self)
        self.setStyleSheet(sis.interface_style(self.interface_style))
        
        L_title = QLabel(self, text="Menu Opzioni")
        L_title.setAccessibleName("an_title")
        self.lay.addWidget(L_title)
        
        L_database_connection = QLabel(self, text="Connessione al database")
        L_database_connection.setAccessibleName("an_section_title")
        self.lay.addWidget(L_database_connection)
        
        L_database_connection_instructions = QLabel(self)
        L_database_connection_instructions.setText("""Il programma usa MongoDB come database
Inserisci il link nella casella qui sotto.
Se hai un database locale il link sarà: mongodb://localhost:27017/""")
        self.lay.addWidget(L_database_connection_instructions)
        
        self.LE_database_connection = QLineEdit(self)
        self.LE_database_connection.setPlaceholderText("Link al database")
        self.LE_database_connection.setText(self.mongodb_connection)
        self.lay.addWidget(self.LE_database_connection)
        
        L_heading = QLabel(self, text="Intestazione")
        L_heading.setAccessibleName("an_section_title")
        self.lay.addWidget(L_heading)
        
        L_heading_instructions = QLabel(self)
        L_heading_instructions.setText("""Inserisci un intestazione, verrà usata sia sulla testa
del programma che ad ogni inizio scontrino""")
        self.lay.addWidget(L_heading_instructions)
        
        self.LE_heading = QLineEdit(self)
        self.LE_heading.setPlaceholderText("Intestazione")
        self.LE_heading.setText(self.heading)
        self.lay.addWidget(self.LE_heading)
        
        L_interface_style = QLabel(self, text="Intefaccia grafica")
        L_interface_style.setAccessibleName("an_section_title")
        self.lay.addWidget(L_interface_style)
        
        L_interface_style_instructions = QLabel(self)
        L_interface_style_instructions.setText("""Stile interfaccia grafica
Seleziona uno stile grafico per il programma""")
        self.lay.addWidget(L_interface_style_instructions)
        
        self.CB_interface_style = QComboBox(self)
        self.CB_interface_style.addItems(["Stile 98", "Stile Tech", "Stile Elegante Chiaro", "Stile Elegante Scuro"])
        self.CB_interface_style.setCurrentText(self.interface_style)
        self.CB_interface_style.currentIndexChanged.connect(self.interface_change)
        self.lay.addWidget(self.CB_interface_style)
        
        L_logo = QLabel(self, text="Logo")
        L_logo.setAccessibleName("an_section_title")
        self.lay.addWidget(L_logo)
        
        self.L_logo_instructions = QLabel(self)
        self.L_logo_instructions.setText(f"""Seleziona un immagine png per il logo
Il logo verrà posizionato in alto a sinistra nell'interfaccia
Le dimensioni ideali sono: 190 x 85 pixel
Attualmente stai usando il file: {self.logo_path}""")
        self.lay.addWidget(self.L_logo_instructions)
        
        self.B_logo = QPushButton(self, text="Seleziona")
        self.B_logo.clicked.connect(self.logo_selection)
        self.lay.addWidget(self.B_logo)
        
        L_icon = QLabel(self, text="Icona")
        L_icon.setAccessibleName("an_section_title")
        self.lay.addWidget(L_icon)
        
        self.L_icon_instructions = QLabel(self)
        self.L_icon_instructions.setText(f"""Seleziona un immagine png per l'icona
L'icona la troverai su ogni finestra
Le dimensioni ideali sono: 51 x 21 pixel
Attualmente stai usando il file: {self.icon_path}""")
        self.lay.addWidget(self.L_icon_instructions)
        
        self.B_icon = QPushButton(self, text="Seleziona")
        self.B_icon.clicked.connect(self.icon_selection)
        self.lay.addWidget(self.B_icon)
        
        self.B_close_and_save = QPushButton(self, text="Chiudi e salva")
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
        logo.setNameFilter("Immagini (*.png)")
        logo.setViewMode(QFileDialog.ViewMode.List)
        logo_path = QFileDialog.getOpenFileName(logo)
        logo_path = Path(logo_path[0])
        self.logo_path = logo_path
        self.L_logo_instructions.setText(f"""Seleziona un immagine png per il logo
Il logo verrà posizionato in alto a sinistra nell'interfaccia
Le dimensioni ideali sono: 190 x 85 pixel
Attualmente stai usando il file: {self.logo_path}""")
    
    def icon_selection(self):
        icon = QFileDialog()
        icon.setFileMode(QFileDialog.FileMode.AnyFile)
        icon.setNameFilter("Immagini (*.png)")
        icon.setViewMode(QFileDialog.ViewMode.List)
        icon_path = QFileDialog.getOpenFileName(icon)
        icon_path = Path(icon_path[0])
        self.icon_path = icon_path
        self.L_icon_instructions.setText(f"""Seleziona un immagine png per l'icona
L'icona la troverai su ogni finestra
Le dimensioni ideali sono: 51 x 21 pixel
Attualmente stai usando il file: {self.icon_path}""")
    
    def close_and_save(self):
        if self.LE_database_connection.text() == "":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("La casella per la connessione al database non può essere vuota")
            return err_msg.exec()
        
        # Impostazione del messaggio di connessione
        
        self.L_database_connection_st.setStyleSheet("color: #FF7800; font: 18px bold Arial;")
        self.L_database_connection_st.setText("Connessione al database in corso....")
        self.L_database_connection_st.show()
        self.L_database_connection_st.repaint()
        
        # Salvataggio file
        
        options_file = open(f"{os.environ['HOME']}/Memberships/options.txt", "w")
        options_file.write(f"db_connection={self.LE_database_connection.text()}\nheading={self.LE_heading.text()}\ninterface={self.CB_interface_style.currentText()}\nlogo={self.logo_path}\nicon={self.icon_path}")
        options_file.close()
        
        # -*-* Riavvio applicazione *-*-
        
        # Lettura file e impostazione variabili
    
        options_file = open(f"{os.environ['HOME']}/Memberships/options.txt", "r")
        try:
            global dbclient
            dbclient = pymongo.MongoClient(options_file.readline().replace("db_connection=", "").replace("\n", ""))
        except:
            options_file.close()
            self.L_database_connection_st.setStyleSheet("color: #8B0B0B; font: 18px bold Arial;")
            self.L_database_connection_st.setText("Connessione al database fallita!")
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
            self.L_database_connection_st.setText("Connessione al database fallita!")
            self.L_database_connection_st.show()

# -*-* Avvio applicazione *-*-

if os.path.exists(f"{os.environ['HOME']}/Memberships") == False: os.mkdir(f"{os.environ['HOME']}/Memberships")

if os.path.exists(f"{os.environ['HOME']}/Memberships/options.txt") == False: first_start_application = 1
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OptionsMenu()
    window.show()
    app.exec()