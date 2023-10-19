import pymongo,sys,subprocess,os,time
import SoftwareInterfaceStyle as sis
from sys import platform
from pathlib import Path
from datetime import datetime
from PyQt6.QtWidgets import (QWidget,QApplication,QGridLayout,QVBoxLayout,QLabel,QLineEdit,QPushButton,QComboBox,QTableWidget,QAbstractItemView,
                             QHeaderView,QMessageBox,QTableWidgetItem,QMenu,QSpinBox,QTextEdit,QCalendarWidget,QFileDialog)
from PyQt6.QtCore import Qt,QDate
from PyQt6.QtGui import QPixmap,QAction,QCursor,QTextCharFormat,QColor,QTextCursor,QIcon
if platform == "win32": import win32print # Importazione del modulo stampa per sistemi operativi Windows

# Versione 1.0

# Variabili globali

heading = ""
dbclient = ""
interface = ""
logo_path = ""
icon_path = ""
first_start_application = 0
total_rows = 11

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.db = dbclient["Bar"] # Apertura database
        
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
        self.lay.addWidget(L_title, 0, 1, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Pulsante Opzioni (Parte superiore destra)
        
        self.B_options_menu = QPushButton(self, text="Opzioni")
        self.B_options_menu.clicked.connect(self.options_menu_open)
        self.lay.addWidget(self.B_options_menu, 0, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Creazione categorie (Parte centrale sinistra)
        
        L_category_creation = QLabel(self, text="Creazione Categorie")
        L_category_creation.setAccessibleName("an_section_title")
        self.lay.addWidget(L_category_creation, 1, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_category_creation_description = QLineEdit(self)
        self.LE_category_creation_description.setPlaceholderText("Descrizione")
        self.lay.addWidget(self.LE_category_creation_description, 2, 0, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_category_creation = QPushButton(self, text="Crea >>")
        self.B_category_creation.clicked.connect(self.category_creation)
        self.lay.addWidget(self.B_category_creation, 2, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Inserimento prodotti (Parte sinistra centrale)
        
        L_products_insert = QLabel(self, text="Inserimento Prodotti")
        L_products_insert.setAccessibleName("an_section_title")
        self.lay.addWidget(L_products_insert, 3, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_products_insert_description = QLineEdit(self)
        self.LE_products_insert_description.setPlaceholderText("Descrizione")
        self.lay.addWidget(self.LE_products_insert_description, 4, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_products_insert_price = QLineEdit(self)
        self.LE_products_insert_price.setPlaceholderText("Prezzo")
        self.lay.addWidget(self.LE_products_insert_price, 4, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_products_insert = QPushButton(self, text="Inserisci >>")
        self.B_products_insert.clicked.connect(self.products_insert)
        self.lay.addWidget(self.B_products_insert, 4, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Impostazioni festa (Parte sinistra centrale)
        
        L_party_settings = QLabel(self, text="Impostazioni festa")
        L_party_settings.setAccessibleName("an_section_title")
        self.lay.addWidget(L_party_settings, 5, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_party_name = QLineEdit(self)
        self.LE_party_name.setPlaceholderText("Intestazione festa")
        self.lay.addWidget(self.LE_party_name, 6, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Impostazione cliente e tavolo (Parte sinistra bassa)
        
        L_customer_table_select = QLabel(self, text="Selezione cliente e numero tavolo")
        L_customer_table_select.setAccessibleName("an_section_title")
        self.lay.addWidget(L_customer_table_select, 7, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_customer_name = QLineEdit(self)
        self.LE_customer_name.setPlaceholderText("Nome cliente")
        self.lay.addWidget(self.LE_customer_name, 8, 0, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.SB_table_select = QSpinBox(self)
        self.SB_table_select.setMinimum(-1)
        self.SB_table_select.setValue(-1)
        self.lay.addWidget(self.SB_table_select, 8, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella note aggiuntive (Parte sinistra bassa)
        
        self.TE_additional_note = QTextEdit(self)
        self.TE_additional_note.setPlaceholderText("Note aggiuntive che appariranno alla fine dello scontrino")
        self.lay.addWidget(self.TE_additional_note, 9, 0, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Funzioni per stampa e salvataggio (Parte sinistra bassa)
        
        L_print_save = QLabel(self, text="Stampa o salva scontrino")
        L_print_save.setAccessibleName("an_section_title")
        self.lay.addWidget(L_print_save, 10, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.CB_printer_list = QComboBox(self) # ComboBox lista stampanti
        printers_list = []        
        if platform == "linux" or platform == "linux2": # Lista stampanti disponibili per sistemi operativi Linux
            printers_list = subprocess.check_output(["lpstat", "-e"], encoding="utf-8")
            printers_list = printers_list.split("\n")[:-1]
        if platform == "win32": # Lista stampanti disponibili per sistemi operativi Windows
            list_count = 0
            printers = subprocess.check_output(["wmic", "printer", "get", "name"], encoding="utf-8")
            for line in printers.splitlines():
                line = line.strip()
                if len(line) > 0 and list_count !=0: printers_list.append(line)
                list_count += 1
        self.CB_printer_list.addItems(printers_list)
        self.lay.addWidget(self.CB_printer_list, 11, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_print = QPushButton(self, text="Stampa") # Pulsante stampa
        self.B_print.clicked.connect(self.print_receipt)
        self.lay.addWidget(self.B_print, 11, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_save = QPushButton(self, text="Salva") # Pulsante salva
        self.B_save.clicked.connect(self.save_receipt)
        self.lay.addWidget(self.B_save, 11, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Selezione Categoria (Parte centrale alta)
        
        self.CB_category_selection = QComboBox(self)
        self.CB_category_selection.currentIndexChanged.connect(self.category_change)
        self.CB_category_selection.setEditable(True)
        self.CB_category_selection.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.lay.addWidget(self.CB_category_selection, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Accesso al database (Parte destra alta)
        
        self.B_database_open = QPushButton(self, text="Apri il Database")
        self.B_database_open.clicked.connect(self.database_open)
        self.lay.addWidget(self.B_database_open, 1, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Tabella categorie (Parte centrale)
        
        self.T_products = QTableWidget(self)
        self.T_products.setColumnCount(2)
        self.T_products.setHorizontalHeaderLabels(["Descrizione", "Prezzo"])
        self.T_products.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_products.doubleClicked.connect(self.add_product)
        self.T_products.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_products.customContextMenuRequested.connect(self.T_products_CM)
        self.T_products_headers = self.T_products.horizontalHeader()
        self.lay.addWidget(self.T_products, 2, 3, total_rows, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
        
        # Tabella scontrino (Parte destra)
        
        self.T_receipt = QTableWidget(self)
        self.T_receipt.setColumnCount(3)
        self.T_receipt.setHorizontalHeaderLabels(["Descrizione", "Quantità", "Totale"])
        self.T_receipt.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_receipt.clicked.connect(self.remove_one_from_receipt)
        self.T_receipt.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_receipt.customContextMenuRequested.connect(self.remove_row_from_receipt)
        self.T_receipt_headers = self.T_receipt.horizontalHeader()
        self.lay.addWidget(self.T_receipt, 2, 4, total_rows-1, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Label per il totale (Parte bassa destra)
        
        self.L_total_receipt = QLabel(self, text="Totale: 0.00 €")
        self.L_total_receipt.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_total_receipt, total_rows, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Funzione Per riempimento iniziale combobox (lasciare sempre per ultima!)
        col = self.db["maincategory"]
        for category in col.find({}, {"_id": 0, "category": 1}):
            if self.CB_category_selection.findText(category["category"]) == -1:
                self.CB_category_selection.addItem(category["category"])
    
    # *-*-* Funzioni per il ridimensionamento della finestra *-*-*

    def resizeEvent(self, event):
        W_width = self.width() / 3
        W_height = self.height() / 3
        
        try:
            # Creazione categorie
            self.LE_category_creation_description.setMinimumWidth(int(W_width / 2))
            # Inserimento dati
            self.LE_products_insert_description.setMinimumWidth(int((W_width / 2) - 40))
            self.LE_products_insert_price.setMinimumWidth(int((W_width / 3) - 40))
            # Combobox categorie
            self.CB_category_selection.setMinimumWidth(int(W_width - 150))
            # Impostazioni festa
            self.LE_party_name.setMinimumWidth(int(W_width))
            # Selezione cliente e tavolo
            self.LE_customer_name.setMinimumWidth(int(W_width - 150))
            # Note aggiuntive scontrino
            self.TE_additional_note.setMinimumSize(int(W_width - 150), int(W_height / 5))
            # Funzioni di stampa e salvataggio
            self.CB_printer_list.setMinimumWidth(int(W_width / 2))
            # Tabella categorie
            self.T_products.setMinimumSize(int(W_width - 150), int((W_height * 2) + 50))
            self.T_products_headers.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            self.T_products_headers.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            # Tabella scontrino
            self.T_receipt.setMinimumSize(int(W_width - 150), int((W_height * 2)))
            self.T_receipt_headers.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            self.T_receipt_headers.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            self.T_receipt_headers.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        except AttributeError:
            pass
    
    # *-*-* Funzioni dei bottoni *-*-*
    
    # Creazione categoria
    
    def category_creation(self):
        # Controllo campo descrizione
        if self.LE_category_creation_description.text() == "":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Il campo descrizione non può essere vuoto!")
            return err_msg.exec()
        
        # Trasformazione campo inserito
        
        self.inserted_category = self.LE_category_creation_description.text().strip().upper()
        
        # Controllo presenza nella ComboBox
        
        if self.CB_category_selection.findText(self.inserted_category) != -1:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Attenzione")
            err_msg.setText("La categoria iserita esiste già \nVuoi eliminarla?")
            err_msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.No)
            err_msg.buttonClicked.connect(self.delete_category)
            self.index_delete_category = self.CB_category_selection.findText(self.inserted_category)
            return err_msg.exec()
        
        self.CB_category_selection.addItem(self.inserted_category)
        
        # Pulizia campo descrizione
        
        self.LE_category_creation_description.setText("")
    
    # Inserimento prodotti
    
    def products_insert(self):
        inserted_price = ""
        # Controllo categoria selezionata
        if self.CB_category_selection.currentText() == "":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Categoria non selezionata!")
            return err_msg.exec()
        
        # Controllo campi inseriti
        
        if self.LE_products_insert_description.text() == "" or self.LE_products_insert_price.text() == "":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Compilare i campi descrizione e prezzo!")
            return err_msg.exec()
        try:
            inserted_price = f"{float(self.LE_products_insert_price.text()):.2f}"
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Il prezzo deve contenere solo numeri e punti! \nEsempio: 5.60")
            return err_msg.exec()
        
        # Trasformazione descrizione
        
        self.inserted_category = self.LE_products_insert_description.text().strip().upper()
        
        # Connessione alla collezione del database
        
        col = self.db["maincategory"]
        
        # Controllo se un articolo è già presente nella tabella
        
        rows = self.T_products.rowCount()
        
        for row in range(rows):
            if self.inserted_category == self.T_products.item(row, 0).text():
                self.T_products.setItem(row, 1, QTableWidgetItem(inserted_price))
                col.update_one({"description": self.inserted_category, "category": self.CB_category_selection.currentText()}, {"$set": {"price": inserted_price}})
                break
        else:
            self.T_products.insertRow(rows)
            self.T_products.setItem(rows, 0, QTableWidgetItem(self.inserted_category))
            self.T_products.setItem(rows, 1, QTableWidgetItem(inserted_price))
            col.insert_one({"description": self.inserted_category, "price": inserted_price, "row": rows, "category": self.CB_category_selection.currentText()})
        
        # Pulizia campi
        
        self.LE_products_insert_description.setText("")
        self.LE_products_insert_price.setText("")
    
    # Eliminazione categoria
    
    def delete_category(self, button):
        if button.text() == "&OK" or button.text() == "OK":
            # Connessione alla collezione del database
            
            col = self.db["maincategory"]
            
            # Controllo presenza nel database ed eventuale eliminazione

            col.delete_many({"category": self.inserted_category})
            
            # Rimozione dalla ComboBox
            
            self.CB_category_selection.removeItem(self.index_delete_category)
            
            # Pulizia campo descrizione
            self.LE_category_creation_description.setText("")
    
    # Cambio categoria
    
    def category_change(self):
        # Rimozione vecchie righe
        if self.T_products.rowCount() != 0:
            for row in reversed(range(self.T_products.rowCount())):
                self.T_products.removeRow(row)
        
        # Connessione alla collezione del database
        
        col = self.db["maincategory"]
        
        # Aggiunta articoli se esistono nel database
        row_count = 0
        for product in col.find({"category": self.CB_category_selection.currentText()}, {"_id": 0}).sort("row"):
            self.T_products.insertRow(row_count)
            self.T_products.setItem(row_count, 0, QTableWidgetItem(product["description"]))
            self.T_products.setItem(row_count, 1, QTableWidgetItem(product["price"]))
            row_count += 1
    
    # Funzione per il set del totale
    
    def set_total_price(self):
        total_price = 0.0
        if self.T_receipt.rowCount() > 0:
            for product in range(self.T_receipt.rowCount()):
                total_price += float(self.T_receipt.item(product, 2).text())
        self.L_total_receipt.setText(f"Totale: {total_price:.2f} €")
    
    # *-*-* Funzioni del menu tabella prodotti *-*-*
    
    # Menu con tasto destro del mouse
    
    def T_products_CM(self):
        if self.T_products.currentRow() == -1: return
        menu = QMenu(self)
        delete_action = QAction("Elimina", self)
        move_up = QAction("Sposta su", self)
        move_down = QAction("Sposta giù", self)
        add = QAction("Aggiungi ++", self)
        remove = QAction("Togli --", self)
        delete_action.triggered.connect(self.delete_product)
        move_up.triggered.connect(self.move_up_product)
        move_down.triggered.connect(self.move_down_product)
        add.triggered.connect(self.add_product)
        remove.triggered.connect(self.remove_product)
        menu.addAction(delete_action)
        menu.addAction(move_up)
        menu.addAction(move_down)
        menu.addAction(add)
        menu.addAction(remove)
        menu.popup(QCursor.pos())
    
    # Funzione elimina
    
    def delete_product(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_products.currentRow()       
        product = self.T_products.item(row, 0).text()
        
        # Eliminazione prodotto
        
        self.T_products.removeRow(row)
        
        # Eliminazione dal Database
        
        col = self.db["maincategory"]
        dbrow = col.find_one({"description": product, "category": self.CB_category_selection.currentText()})
        dbrow = dbrow["row"]
        col.delete_one({"description": product, "category": self.CB_category_selection.currentText()})
        
        # Update della riga degli altri
        
        col.update_many({"category": self.CB_category_selection.currentText(), "row": {"$gt": dbrow}}, {"$inc": {"row": -1}})
    
    # Funzione sposta su
    
    def move_up_product(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_products.currentRow()
        
        # Messaggio di errore se il prodotto si trova nella colonna 0
        
        if row == 0:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Il prodotto selezionato si trova già nella prima casella")
            return err_msg.exec()
        
        # Variabili per descrizione e prezzo
        
        product_desc = self.T_products.item(row, 0).text()
        product_price = self.T_products.item(row, 1).text()
        product_desc_m = self.T_products.item(row -1, 0).text()
        product_price_m = self.T_products.item(row -1, 1).text()
        
        # Modifica nel database
        
        col = self.db["maincategory"]
        col.update_one({"description": product_desc, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": -1}})
        col.update_one({"description": product_desc_m, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": 1}})
        
        # Spostamento del prodotto
        
        self.T_products.setItem(row, 0, QTableWidgetItem(product_desc_m))
        self.T_products.setItem(row, 1, QTableWidgetItem(product_price_m))
        self.T_products.setItem(row -1, 0, QTableWidgetItem(product_desc))
        self.T_products.setItem(row -1, 1, QTableWidgetItem(product_price))
    
    # Funzione sposta giù
    
    def move_down_product(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_products.currentRow()
        
        # Messaggio di errore se il prodotto si trova nel'ultima colonna
        
        if row == self.T_products.rowCount() -1:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle("Errore")
            err_msg.setText("Il prodotto selezionato si trova già nell'ultima casella")
            return err_msg.exec()
        
        # Variabili per descrizione e prezzo
        
        product_desc = self.T_products.item(row, 0).text()
        product_price = self.T_products.item(row, 1).text()
        product_desc_m = self.T_products.item(row +1, 0).text()
        product_price_m = self.T_products.item(row +1, 1).text()
        
        # Modifica nel database
        
        col = self.db["maincategory"]
        col.update_one({"description": product_desc, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": 1}})
        col.update_one({"description": product_desc_m, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": -1}})
        
        # Spostamento del prodotto
        
        self.T_products.setItem(row, 0, QTableWidgetItem(product_desc_m))
        self.T_products.setItem(row, 1, QTableWidgetItem(product_price_m))
        self.T_products.setItem(row +1, 0, QTableWidgetItem(product_desc))
        self.T_products.setItem(row +1, 1, QTableWidgetItem(product_price))
        
    # Funzione aggiungi
    
    def add_product(self):
        # Ricerca del prodotto alla pressione del tasto (o con doppio click sinistro del mouse)
        
        row = self.T_products.currentRow()
        receipt_row = self.T_receipt.rowCount()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        # Inserimento nella tabella a destra
        
        if receipt_row == 0:
            self.T_receipt.insertRow(receipt_row)
            self.T_receipt.setItem(receipt_row, 0, QTableWidgetItem(self.T_products.item(row, 0).text()))
            self.T_receipt.setItem(receipt_row, 1, QTableWidgetItem("1"))
            self.T_receipt.setItem(receipt_row, 2, QTableWidgetItem(self.T_products.item(row, 1).text()))
        else:
            for product in range(receipt_row): # Controllo se l'articolo esiste già
                if self.T_products.item(row, 0).text() == self.T_receipt.item(product, 0).text():
                    quantity = int(self.T_receipt.item(product, 1).text())
                    quantity += 1
                    total_price = f"{float(self.T_products.item(row, 1).text()) * quantity:.2f}"
                    self.T_receipt.setItem(product, 1, QTableWidgetItem(str(quantity)))
                    self.T_receipt.setItem(product, 2, QTableWidgetItem(str(total_price)))
                    break
            else: # Se non esiste
                self.T_receipt.insertRow(receipt_row)
                self.T_receipt.setItem(receipt_row, 0, QTableWidgetItem(self.T_products.item(row, 0).text()))
                self.T_receipt.setItem(receipt_row, 1, QTableWidgetItem("1"))
                self.T_receipt.setItem(receipt_row, 2, QTableWidgetItem(self.T_products.item(row, 1).text()))
        self.set_total_price() # Aggiornamento prezzo totale
    
    # Funzione togli
    
    def remove_product(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_products.currentRow()
        receipt_row = self.T_receipt.rowCount()
        
        # Rimozione dalla tabella di destra
        
        if receipt_row == 0: return
        
        for product in range(self.T_receipt.rowCount()): # Controllo se l'articolo esiste già
            if self.T_products.item(row, 0).text() == self.T_receipt.item(product, 0).text(): # Controllo se l'articolo è a quantità 1 o diversa
                if self.T_receipt.item(product, 1).text() == "1":
                    self.T_receipt.removeRow(product)
                    break
                else:
                    quantity = int(self.T_receipt.item(product, 1).text())
                    quantity -= 1
                    total_price = f"{float(self.T_products.item(row, 1).text()) * quantity:.2f}"
                    self.T_receipt.setItem(product, 1, QTableWidgetItem(str(quantity)))
                    self.T_receipt.setItem(product, 2, QTableWidgetItem(total_price))
                    break 
        
        self.set_total_price() # Aggiornamento prezzo totale
    
    # *-*-* Funzioni della tabella ricevute *-*-*
    
    # Funzione togli una quantità dalla tabella scontrino (click con sinistro del mouse)
    
    def remove_one_from_receipt(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_receipt.currentRow()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        if self.T_receipt.item(row, 1).text() == "1": # Rimozione dalla tabella se l'articolo è a quantità 1
            self.T_receipt.removeRow(row)
        else:
            quantity = int(self.T_receipt.item(row, 1).text())
            total_price = float(self.T_receipt.item(row, 2).text())
            unit_price = total_price / quantity
            quantity -= 1
            total_price = unit_price * quantity
            self.T_receipt.setItem(row, 1, QTableWidgetItem(str(quantity)))
            self.T_receipt.setItem(row, 2, QTableWidgetItem(f"{total_price:.2f}"))
        self.set_total_price()
    
    # Funzione togli la riga dalla tabella scontrino (click con destro del mouse)
    
    def remove_row_from_receipt(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_receipt.currentRow()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        self.T_receipt.removeRow(row)
        self.set_total_price()
    
    # *-*-* Funzioni di stampa e salvataggio *-*-*
    
    # Funzione stampa e salvataggio su DB
    
    def print_receipt(self):
        if self.T_receipt.rowCount() == 0: return # Se la tabella scontrino è vuota
        # Salvataggio su DB
        date_time = datetime.now()
        date = date_time.strftime("%Y%m%d")
        time = date_time.strftime("%H%M%S")
        products = ""
        for row in range(self.T_receipt.rowCount()):
            description = self.T_receipt.item(row, 0).text()
            quantity = self.T_receipt.item(row, 1).text()
            total_price = self.T_receipt.item(row, 2).text()
            products = products + f"{description}@space@{quantity}@space@{total_price}@newline@"
        products = products[:-9]
        col = self.db["receipts"]
        col.insert_one({"receipt_date": date, "receipt_time": time, "receipt_products": products})
        # Stampa scontrino
        printer = self.CB_printer_list.currentText()
        
        printer_string = "" # Stringa da inviare alla stampante
        printer_total = 0.0 # Totale scontrini da inviare alla stampante
        printer_date_time = date_time.strftime("%d/%m/%Y - %H:%M:%S") # Data e ora da inviare alla stampante
        
        printer_string = printer_string + f"\n-*-* {heading} *-*-\n" # Intestazione
        if len(self.LE_party_name.text()) > 0: printer_string = printer_string + f"-*-* {self.LE_party_name.text()} *-*-\n" # Se il nome festa è compilato
        if len(self.LE_customer_name.text()) > 0: printer_string = printer_string + f"\nCliente: {self.LE_customer_name.text()}" # Se il nome cliente è compilato
        if self.SB_table_select.value() != -1: printer_string = printer_string + f"\nNumero tavolo: {self.SB_table_select.value()}" # Se il numero tavolo è compilato
        
        printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        
        # Inserimento prodotti da inviare alla stampante
        for row in range(self.T_receipt.rowCount()):
            description = self.T_receipt.item(row, 0).text()
            quantity = self.T_receipt.item(row, 1).text()
            total_price = self.T_receipt.item(row, 2).text()
            printer_total += float(total_price) # Aggiornamento totale scontrino
            printer_string = printer_string + f"\n{description} - qta {quantity} - Euro {total_price}" # Inserimento prodotto nello scontrino
        printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        printer_string = printer_string + f"\nTotale {printer_total:.2f} Euro"
        printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        
        if len(self.TE_additional_note.toPlainText()) > 0: # Se le note sono compilate
            printer_string = printer_string + f"\n{self.TE_additional_note.toPlainText()}"
            printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        printer_string = printer_string + f"\n\n-*-* {printer_date_time} *-*-" # Data e ora scontrino
        printer_string = printer_string + "\n\n\n\n\n\n\n\n\n\n\n" # Fine scontrino
        
        # Pulizia caselle e tabella
        
        self.LE_customer_name.clear()
        self.SB_table_select.setValue(-1)
        self.TE_additional_note.clear()
        for row in reversed(range(self.T_receipt.rowCount())):
            self.T_receipt.removeRow(row)
        self.set_total_price()
        
        # Stampa
        
        printer_string = bytes(printer_string, encoding="utf-8") # Preparazione della stringa per essere inviata alla stampante
        
        if platform == "linux" or platform == "linux2": # Funzioni di stampa per sistemi operativi Linux
            print_command = subprocess.Popen(["lpr", f"-P{printer}"], stdin=subprocess.PIPE)
            print_command.stdin.write(b"\x1d\x21\x09" + printer_string)
            print_command.stdin.close()
        if platform == "win32": # Funzioni di stampa per sistemi operativi Windows
            windows_printer = win32print.OpenPrinter(printer)
            try:
                printer_job = win32print.StartDocPrinter(windows_printer, (f"job {date} {time}", None, "RAW"))
                try:
                    win32print.StartPagePrinter(windows_printer)
                    win32print.WritePrinter(windows_printer, b"\x1d\x21\x09" + printer_string)
                    win32print.EndPagePrinter(windows_printer)
                finally: win32print.EndDocPrinter(windows_printer)
            finally: win32print.ClosePrinter(windows_printer)
    
    # Funzione salvataggio su DB
    def save_receipt(self):
        if self.T_receipt.rowCount() == 0: return # Se la tabella scontrino è vuota
        date_time = datetime.now()
        date = date_time.strftime("%Y%m%d")
        time = date_time.strftime("%H%M%S")
        products = ""
        for row in range(self.T_receipt.rowCount()):
            description = self.T_receipt.item(row, 0).text()
            quantity = self.T_receipt.item(row, 1).text()
            total_price = self.T_receipt.item(row, 2).text()
            products = products + f"{description}@space@{quantity}@space@{total_price}@newline@"
        products = products[:-9]
        col = self.db["receipts"]
        col.insert_one({"receipt_date": date, "receipt_time": time, "receipt_products": products})
        for row in reversed(range(self.T_receipt.rowCount())):
            self.T_receipt.removeRow(row)
        self.set_total_price()
    
    # *-*-* Funzione apertura finestra database *-*-*
    
    def database_open(self):
        self.database_window = DatabaseWindow()
        self.database_window.show()
    
    # *-*-* Funzione apertura finestra opzioni *-*-*
    
    def options_menu_open(self):
        self.options_window = OptionsMenu()
        self.options_window.show()
        self.close()

# *-*-* Finestra Database *-*-*

class DatabaseWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.db = dbclient["Bar"] # Apertura database
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"Database {heading}")
        self.setFixedSize(480, 640)
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        self.setStyleSheet(sis.interface_style(interface))
        
        # Calendario
        
        self.date_selection = 0
        self.format = QTextCharFormat()
        if interface == "Stile 98":
            self.format.setBackground(QColor("yellow"))
            self.format.setForeground(QColor("black"))
        if interface == "Stile Tech":
            self.format.setBackground(QColor("#2DB00D"))
            self.format.setForeground(QColor("black"))
        if interface == "Stile Elegante Chiaro":
            self.format.setBackground(QColor("#044B11"))
            self.format.setForeground(QColor("black"))
        if interface == "Stile Elegante Scuro":
            self.format.setBackground(QColor("#444342"))
            self.format.setForeground(QColor("white"))
        
        self.CW_calendar = QCalendarWidget(self)
        self.CW_calendar.setGridVisible(True)
        self.first_date = self.CW_calendar.selectedDate().toString("yyyy-MM-dd")
        self.first_date = self.first_date.replace("-", "")
        self.second_date = ""
        self.CW_calendar.clicked.connect(self.date_select)
        self.lay.addWidget(self.CW_calendar, 0, 0, 4, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Label istruzioni orari
        
        L_time_range = QLabel(self, text="Selezione orari\nNella casella qui sotto\npuoi selezionare un\nrange di orari.\nEsempio 17:30-20:00")
        self.lay.addWidget(L_time_range, 0, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella selezione orari
        
        self.LE_time_range = QLineEdit(self)
        self.LE_time_range.setFixedWidth(150)
        self.LE_time_range.setPlaceholderText("Orari")
        self.lay.addWidget(self.LE_time_range, 1, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Bottone interroga
        
        self.B_query_database = QPushButton(self, text="Interroga")
        self.B_query_database.clicked.connect(self.query_database)
        self.lay.addWidget(self.B_query_database, 2, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Bottone elimina dal database
        
        self.B_delete_database = QPushButton(self, text="Elimina")
        self.B_delete_database.clicked.connect(self.delete_database)
        self.lay.addWidget(self.B_delete_database, 3, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Text Edit per risposta Database
        
        self.TE_database_response = QTextEdit(self)
        self.TE_database_response.setReadOnly(True)
        self.TE_database_response.setFixedHeight(450)
        self.lay.addWidget(self.TE_database_response, 4, 0, 1, 2, Qt.AlignmentFlag.AlignTop)
    
    def date_select(self):
        if self.date_selection == 0: # Selezione della seconda data
            self.second_date = self.CW_calendar.selectedDate().toString("yyyy-MM-dd")
            self.second_date = self.second_date.replace("-", "")
            if int(self.first_date) > int(self.second_date): # Inversione date se la prima è maggiore della seconda
                date_inv = self.first_date
                self.first_date = self.second_date
                self.second_date = date_inv
            
            if self.second_date == self.first_date: return # Blocco della funzione se si seleziona la stessa data
            date = int(self.first_date)
            while date <= int(self.second_date):
                year = int(str(date)[:4])
                month = int(str(date)[4:6])
                day = int(str(date)[6:])
                self.CW_calendar.setDateTextFormat(QDate(year, month, day), self.format)
                date += 1
            self.date_selection = 1
            return
        if self.date_selection == 1: # Selezione della prima data
            self.first_date = self.CW_calendar.selectedDate().toString("yyyy-MM-dd")
            self.first_date = self.first_date.replace("-", "")
            self.second_date = ""
            self.CW_calendar.setDateTextFormat(QDate(), QTextCharFormat())
            self.date_selection = 0
            return
    
    # Interrogazione database
    
    def query_database(self):
        col = self.db["receipts"]
        detail_tot = {}
        total_receipts = 0.0
        result = 0 # Controllo se c'è qualcosa nel database
        self.TE_database_response.clear()
        self.TE_database_response.insertPlainText("-*-* Dettaglio Vendite *-*-\n\n")
        
        if self.second_date == "": # Se è selezionata una sola data
            firstdate = f"{self.first_date[6:]}/{self.first_date[4:6]}/{self.first_date[:4]}"
            if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                for products in col.find({"receipt_date": self.first_date}, {"_id": 0}):
                    result = 1
                    self.TE_database_response.append(f"\nScontrino del {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]} - orario {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
                    for product in str(products["receipt_products"]).split("@newline@"):
                        detail = product.split("@space@")
                        self.TE_database_response.append(f"\n{detail[0]} - qta {detail[1]} - € {detail[2]}")
                        total_receipts += float(detail[2])
                        if detail[0] not in detail_tot:
                            detail_tot[detail[0]] = {}
                            detail_tot[detail[0]]["quantity"] = int(detail[1])
                            detail_tot[detail[0]]["total"] = float(detail[2])
                        else:
                            detail_tot[detail[0]]["quantity"] += int(detail[1])
                            detail_tot[detail[0]]["total"] += float(detail[2])
                    self.TE_database_response.append("\n --------------------------\n")
                    
                if result == 0:
                    self.TE_database_response.clear()
                    self.TE_database_response.setPlainText("Nessun dato per la data selezionata!")
                    return
            
            else: # Se è stato selezionato un orario
                time_string = self.LE_time_range.text().replace(" ", "")
                if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                    self.TE_database_response.clear()
                    self.TE_database_response.setPlainText("Formato ora inserito non corretto!")
                    return
                time_string = time_string.replace(":", "-")
                time_string = time_string.split("-")
                first_time = f"{time_string[0]}{time_string[1]}00"
                second_time = f"{time_string[2]}{time_string[3]}00"
                if int(first_time) > int(second_time):
                    self.TE_database_response.clear()
                    self.TE_database_response.setPlainText("Formato ora inserito non corretto!")
                    return
                for products in col.find({"receipt_date": self.first_date, "receipt_time": {"$gte": first_time, "$lte": second_time}}, {"_id": 0}):
                    result = 1
                    self.TE_database_response.append(f"\nScontrino del {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]} - orario {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
                    for product in str(products["receipt_products"]).split("@newline@"):
                        detail = product.split("@space@")
                        self.TE_database_response.append(f"\n{detail[0]} - qta {detail[1]} - € {detail[2]}")
                        total_receipts += float(detail[2])
                        if detail[0] not in detail_tot:
                            detail_tot[detail[0]] = {}
                            detail_tot[detail[0]]["quantity"] = int(detail[1])
                            detail_tot[detail[0]]["total"] = float(detail[2])
                        else:
                            detail_tot[detail[0]]["quantity"] += int(detail[1])
                            detail_tot[detail[0]]["total"] += float(detail[2])
                    self.TE_database_response.append("\n --------------------------\n")
                    
                if result == 0:
                    self.TE_database_response.clear()
                    self.TE_database_response.setPlainText("Nessun dato per la data e l'orario selezionato!")
                    return
                       
            text_cursor = QTextCursor(self.TE_database_response.document()) # Spostamento del cursore all'inizio
            text_cursor.setPosition(0)
            self.TE_database_response.setTextCursor(text_cursor)
            total_string = f"-*-* Totale vendite del {firstdate} *-*-\n"
            for detail in detail_tot: # Loop del dizionario con il totale vendite
                total_string = total_string + f"\n{detail} - qta {detail_tot[detail]['quantity']} - € {detail_tot[detail]['total']:.2f}"
            total_string = total_string + f"\n\n-*-* Totale complessivo € {total_receipts:.2f} *-*-\n --------------------------\n\n\n"
            self.TE_database_response.insertPlainText(total_string)
        
        else: # Se sono state selezionate 2 date
            firstdate = f"{self.first_date[6:]}/{self.first_date[4:6]}/{self.first_date[:4]}"
            seconddate = f"{self.second_date[6:]}/{self.second_date[4:6]}/{self.second_date[:4]}"
            if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                for products in col.find({"receipt_date": {"$gte": self.first_date, "$lte": self.second_date}}, {"_id": 0}):
                    result = 1
                    self.TE_database_response.append(f"\nScontrino del {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]} - orario {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
                    for product in str(products["receipt_products"]).split("@newline@"):
                        detail = product.split("@space@")
                        self.TE_database_response.append(f"\n{detail[0]} - qta {detail[1]} - € {detail[2]}")
                        total_receipts += float(detail[2])
                        if detail[0] not in detail_tot:
                            detail_tot[detail[0]] = {}
                            detail_tot[detail[0]]["quantity"] = int(detail[1])
                            detail_tot[detail[0]]["total"] = float(detail[2])
                        else:
                            detail_tot[detail[0]]["quantity"] += int(detail[1])
                            detail_tot[detail[0]]["total"] += float(detail[2])
                    self.TE_database_response.append("\n --------------------------\n")
                    
                if result == 0:
                    self.TE_database_response.clear()
                    self.TE_database_response.setPlainText("Nessun dato per le date selezionate!")
                    return
            
            else: # Se è stato selezionato un orario
                time_string = self.LE_time_range.text().replace(" ", "")
                if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                    self.TE_database_response.clear()
                    self.TE_database_response.setPlainText("Formato ora inserito non corretto!")
                    return
                time_string = time_string.replace(":", "-")
                time_string = time_string.split("-")
                first_time = f"{time_string[0]}{time_string[1]}00"
                second_time = f"{time_string[2]}{time_string[3]}00"
                for products in col.find({"receipt_date": {"$gte": self.first_date, "$lte": self.second_date}}, {"_id": 0}):
                    if str(products["receipt_date"]) == self.first_date: # Controllo dell'orario nella prima data
                        if int(products["receipt_time"]) < int(first_time): continue
                    if str(products["receipt_date"]) == self.second_date: # Controllo dell'orario nella seconda data
                        if int(products["receipt_time"]) > int(second_time): continue
                        
                    self.TE_database_response.append(f"\nScontrino del {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]} - orario {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
                    for product in str(products["receipt_products"]).split("@newline@"):
                        detail = product.split("@space@")
                        self.TE_database_response.append(f"\n{detail[0]} - qta {detail[1]} - € {detail[2]}")
                        total_receipts += float(detail[2])
                        if detail[0] not in detail_tot:
                            detail_tot[detail[0]] = {}
                            detail_tot[detail[0]]["quantity"] = int(detail[1])
                            detail_tot[detail[0]]["total"] = float(detail[2])
                        else:
                            detail_tot[detail[0]]["quantity"] += int(detail[1])
                            detail_tot[detail[0]]["total"] += float(detail[2])
                        result = 1
                    self.TE_database_response.append("\n --------------------------\n")
                    
                if result == 0:
                    self.TE_database_response.clear()
                    self.TE_database_response.setPlainText("Nessun dato per le date e l'orario selezionati!")
                    return
            
            text_cursor = QTextCursor(self.TE_database_response.document()) # Spostamento del cursore all'inizio
            text_cursor.setPosition(0)
            self.TE_database_response.setTextCursor(text_cursor)
            total_string = f"-*-* Totale vendite dal {firstdate} al {seconddate} *-*-\n"
            for detail in detail_tot: # Loop del dizionario con il totale vendite
                total_string = total_string + f"\n{detail} - qta {detail_tot[detail]['quantity']} - € {detail_tot[detail]['total']:.2f}"
            total_string = total_string + f"\n\n-*-* Totale complessivo € {total_receipts:.2f} *-*-\n --------------------------\n\n\n"
            self.TE_database_response.insertPlainText(total_string)
    
    # Eliminazione dal database
    
    def delete_database(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Attenzione")
        msg.setText("Stai per eliminare i dati dal database\nL'operazione non è annullabile!\nVuoi continuare?")
        msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.No)
        msg.buttonClicked.connect(self.delete_database_confirm)
        return msg.exec()
    
    def delete_database_confirm(self, button):
        if button.text() == "&OK" or button.text() == "OK":
            col = self.db["receipts"]
            if self.second_date == "": # Se è selezionata una sola data
                if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                    col.delete_many({"receipt_date": self.first_date})
                else: # Se è stato selezionato un orario
                    time_string = self.LE_time_range.text().replace(" ", "")
                    if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                        self.TE_database_response.clear()
                        self.TE_database_response.setPlainText("Formato ora inserito non corretto!")
                        return
                    time_string = time_string.replace(":", "-")
                    time_string = time_string.split("-")
                    first_time = f"{time_string[0]}{time_string[1]}00"
                    second_time = f"{time_string[2]}{time_string[3]}00"
                    if int(first_time) > int(second_time):
                        self.TE_database_response.clear()
                        self.TE_database_response.setPlainText("Formato ora inserito non corretto!")
                        return
                    col.delete_many({"receipt_date": self.first_date, "receipt_time": {"$gte": first_time, "$lte": second_time}})
            else: # Se sono state selezionate 2 date
                if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                    col.delete_many({"receipt_date": {"$gte": self.first_date, "$lte": self.second_date}})
                else: # Se è stato selezionato un orario
                    time_string = self.LE_time_range.text().replace(" ", "")
                    if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                        self.TE_database_response.clear()
                        self.TE_database_response.setPlainText("Formato ora inserito non corretto!")
                        return
                    time_string = time_string.replace(":", "-")
                    time_string = time_string.split("-")
                    first_time = f"{time_string[0]}{time_string[1]}00"
                    second_time = f"{time_string[2]}{time_string[3]}00"
                    for products in col.find({"receipt_date": {"$gte": self.first_date, "$lte": self.second_date}}, {"_id": 0}):
                        if str(products["receipt_date"]) == self.first_date: # Controllo dell'orario nella prima data
                            if int(products["receipt_time"]) < int(first_time): continue
                        if str(products["receipt_date"]) == self.second_date: # Controllo dell'orario nella seconda data
                            if int(products["receipt_time"]) > int(second_time): continue
                        col.delete_one({"receipt_date": products["receipt_date"], "receipt_time": products["receipt_time"]})
            self.TE_database_response.clear()
            self.TE_database_response.insertPlainText("-*-* Eliminazione effettuata *-*-\n\n")

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
    
        if os.path.exists(f"{os.environ['HOME']}/Orders/options.txt"):
            options_file = open(f"{os.environ['HOME']}/Orders/options.txt", "r")
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
        
        options_file = open(f"{os.environ['HOME']}/Orders/options.txt", "w")
        options_file.write(f"db_connection={self.LE_database_connection.text()}\nheading={self.LE_heading.text()}\ninterface={self.CB_interface_style.currentText()}\nlogo={self.logo_path}\nicon={self.icon_path}")
        options_file.close()
        
        # -*-* Riavvio applicazione *-*-
        
        # Lettura file e impostazione variabili
    
        options_file = open(f"{os.environ['HOME']}/Orders/options.txt", "r")
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

if os.path.exists(f"{os.environ['HOME']}/Orders") == False: os.mkdir(f"{os.environ['HOME']}/Orders")

if os.path.exists(f"{os.environ['HOME']}/Orders/options.txt") == False: first_start_application = 1
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OptionsMenu()
    window.show()
    app.exec()