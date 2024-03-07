import pymongo,sys,subprocess,os,time
from bson.objectid import ObjectId
import SoftwareInterfaceStyle as sis
from sys import platform
from pathlib import Path
from datetime import datetime
from PyQt6.QtWidgets import (QWidget,QApplication,QGridLayout,QVBoxLayout,QLabel,QLineEdit,QPushButton,QComboBox,QTableWidget,QAbstractItemView,
                             QHeaderView,QMessageBox,QTableWidgetItem,QMenu,QSpinBox,QTextEdit,QCalendarWidget,QFileDialog,QInputDialog,QCheckBox)
from PyQt6.QtCore import Qt,QDate
from PyQt6.QtGui import QPixmap,QAction,QCursor,QTextCharFormat,QColor,QTextCursor,QIcon
import Orders_Language as lang
if platform == "win32": import win32print # Importazione del modulo stampa per sistemi operativi Windows

# Versione 1.0.2-r8

# Debug mode

debug_mode = False

# Variabili globali

heading = ""
dbclient = ""
interface = ""
language = ""
logo_path = ""
icon_path = ""
menu_dict = {}
first_start_application = 0
total_rows = 13

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
        
        self.B_options_menu = QPushButton(self, text=lang.msg(language, 0, "MainWindow"))
        self.B_options_menu.clicked.connect(self.options_menu_open)
        self.lay.addWidget(self.B_options_menu, 0, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Creazione ed eliminazione categorie (Parte centrale sinistra)
        
        L_category_creation = QLabel(self, text=lang.msg(language, 1, "MainWindow"))
        L_category_creation.setAccessibleName("an_section_title")
        self.lay.addWidget(L_category_creation, 1, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_category_creation_description = QLineEdit(self)
        self.LE_category_creation_description.setPlaceholderText(lang.msg(language, 2, "MainWindow"))
        self.lay.addWidget(self.LE_category_creation_description, 2, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_category_creation = QPushButton(self, text=lang.msg(language, 3, "MainWindow"))
        self.B_category_creation.clicked.connect(self.category_creation)
        self.lay.addWidget(self.B_category_creation, 2, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_category_delete = QPushButton(self, text=lang.msg(language, 26, "MainWindow"))
        self.B_category_delete.clicked.connect(self.delete_category)
        self.lay.addWidget(self.B_category_delete, 2, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Inserimento prodotti (Parte sinistra centrale)
        
        L_products_insert = QLabel(self, text=lang.msg(language, 4, "MainWindow"))
        L_products_insert.setAccessibleName("an_section_title")
        self.lay.addWidget(L_products_insert, 3, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_products_insert_description = QLineEdit(self)
        self.LE_products_insert_description.setPlaceholderText(lang.msg(language, 2, "MainWindow"))
        self.lay.addWidget(self.LE_products_insert_description, 4, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_products_insert_price = QLineEdit(self)
        self.LE_products_insert_price.setPlaceholderText(lang.msg(language, 5, "MainWindow"))
        self.lay.addWidget(self.LE_products_insert_price, 4, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_products_insert = QPushButton(self, text=lang.msg(language, 6, "MainWindow"))
        self.B_products_insert.clicked.connect(self.products_insert)
        self.lay.addWidget(self.B_products_insert, 4, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Impostazioni festa (Parte sinistra centrale)
        
        L_party_settings = QLabel(self, text=lang.msg(language, 7, "MainWindow"))
        L_party_settings.setAccessibleName("an_section_title")
        self.lay.addWidget(L_party_settings, 5, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_party_name = QLineEdit(self)
        self.LE_party_name.setPlaceholderText(lang.msg(language, 8, "MainWindow"))
        self.lay.addWidget(self.LE_party_name, 6, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Impostazione cliente e tavolo (Parte sinistra bassa)
        
        L_customer_table_select = QLabel(self, text=lang.msg(language, 9, "MainWindow"))
        L_customer_table_select.setAccessibleName("an_section_title")
        self.lay.addWidget(L_customer_table_select, 7, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.LE_customer_name = QLineEdit(self)
        self.LE_customer_name.setPlaceholderText(lang.msg(language, 10, "MainWindow"))
        self.lay.addWidget(self.LE_customer_name, 8, 0, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.SB_table_select = QSpinBox(self)
        self.SB_table_select.setMinimum(-1)
        self.SB_table_select.setValue(-1)
        self.lay.addWidget(self.SB_table_select, 8, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella note aggiuntive (Parte sinistra bassa)
        
        self.TE_additional_note = QTextEdit(self)
        self.TE_additional_note.setPlaceholderText(lang.msg(language, 11, "MainWindow"))
        self.lay.addWidget(self.TE_additional_note, 9, 0, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Ordini da inviare (Parte sinistra bassa)
        
        L_order_to_send = QLabel(self, text=lang.msg(language, 46, "MainWindow"))
        L_order_to_send.setAccessibleName("an_section_title")
        self.lay.addWidget(L_order_to_send, 10, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.list_category_to_send = []
        self.ChB_category_to_send = QCheckBox(self, text=lang.msg(language, 44, "MainWindow"))
        self.ChB_category_to_send.stateChanged.connect(self.category_to_send_changed)
        self.lay.addWidget(self.ChB_category_to_send, 11, 0, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_see_orders = QPushButton(self, text=lang.msg(language, 47, "MainWindow"))
        self.B_see_orders.clicked.connect(self.open_orders_table)
        self.lay.addWidget(self.B_see_orders, 11, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Funzioni per stampa e salvataggio (Parte sinistra bassa)
        
        L_print_save = QLabel(self, text=lang.msg(language, 12, "MainWindow"))
        L_print_save.setAccessibleName("an_section_title")
        self.lay.addWidget(L_print_save, 12, 0, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.CB_printer_list = QComboBox(self) # ComboBox lista stampanti
        printers_list = []
        if debug_mode == True: self.CB_printer_list.addItem("DEBUG")
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
        self.lay.addWidget(self.CB_printer_list, 13, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_print = QPushButton(self, text=lang.msg(language, 13, "MainWindow")) # Pulsante stampa
        self.B_print.clicked.connect(self.print_receipt)
        self.lay.addWidget(self.B_print, 13, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        self.B_save = QPushButton(self, text=lang.msg(language, 14, "MainWindow")) # Pulsante salva
        self.B_save.clicked.connect(self.save_receipt)
        self.lay.addWidget(self.B_save, 13, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Selezione Categoria (Parte centrale alta)
        
        self.CB_category_selection = QComboBox(self)
        self.CB_category_selection.currentIndexChanged.connect(self.category_change)
        self.CB_category_selection.setEditable(True)
        self.CB_category_selection.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.lay.addWidget(self.CB_category_selection, 1, 3, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Accesso al database (Parte destra alta)
        
        self.B_database_open = QPushButton(self, text=lang.msg(language, 15, "MainWindow"))
        self.B_database_open.clicked.connect(self.database_open)
        self.lay.addWidget(self.B_database_open, 1, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Tabella categorie (Parte centrale)
        
        self.T_products = QTableWidget(self)
        self.T_products.setColumnCount(2)
        self.T_products.setHorizontalHeaderLabels([lang.msg(language, 2, "MainWindow"), lang.msg(language, 5, "MainWindow")])
        self.T_products.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_products.doubleClicked.connect(self.add_product)
        self.T_products.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_products.customContextMenuRequested.connect(self.T_products_CM)
        self.T_products_headers = self.T_products.horizontalHeader()
        self.lay.addWidget(self.T_products, 2, 3, total_rows, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
        
        # Tabella scontrino (Parte destra)
        
        self.T_receipt = QTableWidget(self)
        self.T_receipt.setColumnCount(3)
        self.T_receipt.setHorizontalHeaderLabels([lang.msg(language, 2, "MainWindow"), lang.msg(language, 16, "MainWindow"), lang.msg(language, 17, "MainWindow")])
        self.T_receipt.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_receipt.doubleClicked.connect(self.remove_one_from_receipt)
        self.T_receipt.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_receipt.customContextMenuRequested.connect(self.remove_row_from_receipt)
        self.T_receipt_headers = self.T_receipt.horizontalHeader()
        self.lay.addWidget(self.T_receipt, 2, 4, total_rows-1, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        self.category_in_receipt = [] # Salvataggio delle categorie quando si inserisce un articolo allo scontrino
        
        # Label per il totale (Parte bassa destra)
        
        self.L_total_receipt = QLabel(self, text=f"{lang.msg(language, 17, 'MainWindow')}: 0.00 {lang.msg(language, 18, 'MainWindow')}")
        self.L_total_receipt.setAccessibleName("an_section_title")
        self.lay.addWidget(self.L_total_receipt, total_rows, 4, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignRight)
        
        # Funzione Per riempimento iniziale combobox (lasciare sempre per ultima!)
        try: # Controllo della connessione al database
            col = self.db["maincategory"]
            for category in col.find({}, {"_id": 0, "category": 1}):
                if self.CB_category_selection.findText(category["category"]) == -1:
                    self.CB_category_selection.addItem(category["category"])
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
    
    # *-*-* Funzioni per il ridimensionamento della finestra *-*-*

    def resizeEvent(self, event):
        W_width = self.width() / 3
        W_height = self.height() / 3
        
        try:
            # Creazione categorie
            self.LE_category_creation_description.setMinimumWidth(int(W_width / 2) - 40)
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
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 20, "MainWindow"))
            return err_msg.exec()
        
        # Trasformazione campo inserito
        
        self.inserted_category = self.LE_category_creation_description.text().strip().upper()
        
        # Controllo presenza nella ComboBox
        
        if self.CB_category_selection.findText(self.inserted_category) != -1:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 21, "MainWindow"))
            err_msg.setText(lang.msg(language, 22, "MainWindow"))
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
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 23, "MainWindow"))
            return err_msg.exec()
        
        # Controllo campi inseriti
        
        if self.LE_products_insert_description.text() == "" or self.LE_products_insert_price.text() == "":
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 24, "MainWindow"))
            return err_msg.exec()
        try:
            inserted_price = f"{float(self.LE_products_insert_price.text()):.2f}"
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 25, "MainWindow"))
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
                try: # Controllo della connessione al database
                    col.update_one({"description": self.inserted_category, "category": self.CB_category_selection.currentText()}, {"$set": {"price": inserted_price}})
                except:
                    err_msg = QMessageBox(self)
                    err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
                    err_msg.setText(lang.msg(language, 48, "MainWindow"))
                    return err_msg.exec()
                break
        else:
            self.T_products.insertRow(rows)
            self.T_products.setItem(rows, 0, QTableWidgetItem(self.inserted_category))
            self.T_products.setItem(rows, 1, QTableWidgetItem(inserted_price))
            try: # Controllo della connessione al database
                col.insert_one({"description": self.inserted_category, "price": inserted_price, "row": rows, "category": self.CB_category_selection.currentText()})
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
                err_msg.setText(lang.msg(language, 48, "MainWindow"))
                return err_msg.exec()
        
        # Pulizia campi
        
        self.LE_products_insert_description.setText("")
        self.LE_products_insert_price.setText("")
    
    # Eliminazione categoria
    
    def delete_category(self):
        col = self.db["maincategory"]
        
        # Controllo presenza nel database ed eventuale eliminazione

        try: # Controllo della connessione al database
            col.delete_many({"category": self.inserted_category})
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
        
        # Rimozione dalla lista
        
        if self.CB_category_selection.currentText() in self.list_category_to_send:
            self.list_category_to_send.pop(self.list_category_to_send.index(self.CB_category_selection.currentText()))
        
        # Rimozione dalla ComboBox
        
        self.CB_category_selection.removeItem(self.CB_category_selection.currentIndex())
    
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
        try: # Controllo della connessione al database
            for product in col.find({"category": self.CB_category_selection.currentText()}, {"_id": 0}).sort("row"):
                self.T_products.insertRow(row_count)
                self.T_products.setItem(row_count, 0, QTableWidgetItem(product["description"]))
                self.T_products.setItem(row_count, 1, QTableWidgetItem(product["price"]))
                row_count += 1
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
        # Modifica della checkbox se la categoria è già stata segnata
        if self.CB_category_selection.currentText() in self.list_category_to_send: self.ChB_category_to_send.setChecked(True)
        else: self.ChB_category_to_send.setChecked(False)
    
    # Funzione invio ordine automatico
    
    def category_to_send_changed(self):
        if self.ChB_category_to_send.isChecked() == True:
            if self.CB_category_selection.currentText() not in self.list_category_to_send:
                self.list_category_to_send.append(self.CB_category_selection.currentText())
        else:
            try: self.list_category_to_send.pop(self.list_category_to_send.index(self.CB_category_selection.currentText()))
            except: pass
    
    # Funzione per il set del totale
    
    def set_total_price(self):
        total_price = 0.0
        if self.T_receipt.rowCount() > 0:
            for product in range(self.T_receipt.rowCount()):
                total_price += float(self.T_receipt.item(product, 2).text())
        self.L_total_receipt.setText(f"{lang.msg(language, 17, 'MainWindow')}: {total_price:.2f} {lang.msg(language, 18, 'MainWindow')}")
    
    # *-*-* Funzioni del menu tabella prodotti *-*-*
    
    # Menu con tasto destro del mouse
    
    def T_products_CM(self):
        if self.T_products.currentRow() == -1: return
        menu = QMenu(self)
        delete_action = QAction(lang.msg(language, 26, "MainWindow"), self)
        move_up = QAction(lang.msg(language, 27, "MainWindow"), self)
        move_down = QAction(lang.msg(language, 28, "MainWindow"), self)
        add_specific_quantity = QAction(lang.msg(language, 29, "MainWindow"), self)
        add = QAction(lang.msg(language, 30, "MainWindow"), self)
        remove = QAction(lang.msg(language, 31, "MainWindow"), self)
        create_menu = QAction(lang.msg(language, 32, "MainWindow"), self)
        delete_action.triggered.connect(self.delete_product)
        move_up.triggered.connect(self.move_up_product)
        move_down.triggered.connect(self.move_down_product)
        add_specific_quantity.triggered.connect(self.add_product_specific_quantity)
        add.triggered.connect(self.add_product)
        remove.triggered.connect(self.remove_product)
        create_menu.triggered.connect(self.create_menu)
        menu.addActions([delete_action, move_up, move_down, add_specific_quantity, add, remove, create_menu])
        menu.popup(QCursor.pos())
    
    # Funzione elimina
    
    def delete_product(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_products.currentRow()       
        product = self.T_products.item(row, 0).text()
        
        # Eliminazione prodotto
        
        self.T_products.removeRow(row)
        
        # Eliminazione dal Database
        
        try: # Controllo della connessione al database
            col = self.db["maincategory"]
            dbrow = col.find_one({"description": product, "category": self.CB_category_selection.currentText()})
            dbrow = dbrow["row"]
            col.delete_one({"description": product, "category": self.CB_category_selection.currentText()})
            
            # Update della riga degli altri
            
            col.update_many({"category": self.CB_category_selection.currentText(), "row": {"$gt": dbrow}}, {"$inc": {"row": -1}})
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
    
    # Funzione sposta su
    
    def move_up_product(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_products.currentRow()
        
        # Messaggio di errore se il prodotto si trova nella colonna 0
        
        if row == 0:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 33, "MainWindow"))
            return err_msg.exec()
        
        # Variabili per descrizione e prezzo
        
        product_desc = self.T_products.item(row, 0).text()
        product_price = self.T_products.item(row, 1).text()
        product_desc_m = self.T_products.item(row -1, 0).text()
        product_price_m = self.T_products.item(row -1, 1).text()
        
        # Modifica nel database
        
        try: # Controllo della connessione al database
            col = self.db["maincategory"]
            col.update_one({"description": product_desc, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": -1}})
            col.update_one({"description": product_desc_m, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": 1}})
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
        
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
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 34, "MainWindow"))
            return err_msg.exec()
        
        # Variabili per descrizione e prezzo
        
        product_desc = self.T_products.item(row, 0).text()
        product_price = self.T_products.item(row, 1).text()
        product_desc_m = self.T_products.item(row +1, 0).text()
        product_price_m = self.T_products.item(row +1, 1).text()
        
        # Modifica nel database
        
        try: # Controllo della connessione al database
            col = self.db["maincategory"]
            col.update_one({"description": product_desc, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": 1}})
            col.update_one({"description": product_desc_m, "category": self.CB_category_selection.currentText()}, {"$inc": {"row": -1}})
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
        
        # Spostamento del prodotto
        
        self.T_products.setItem(row, 0, QTableWidgetItem(product_desc_m))
        self.T_products.setItem(row, 1, QTableWidgetItem(product_price_m))
        self.T_products.setItem(row +1, 0, QTableWidgetItem(product_desc))
        self.T_products.setItem(row +1, 1, QTableWidgetItem(product_price))
    
    # Funzione aggiungi quantità specifica
    
    def add_product_specific_quantity(self):
        msg = QInputDialog(self)
        msg.setWindowTitle(lang.msg(language, 35, "MainWindow"))
        msg.setLabelText(lang.msg(language, 36, "MainWindow"))
        msg.setOkButtonText(lang.msg(language, 37, "MainWindow"))
        msg.setCancelButtonText(lang.msg(language, 38, "MainWindow"))
        if msg.exec() == 1:
            number = msg.textValue()
            try: number = int(number)
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
                err_msg.setText(lang.msg(language, 39, "MainWindow"))
                return err_msg.exec()
            
            # Inserimento nella tabella a destra
            row = self.T_products.currentRow()
            receipt_row = self.T_receipt.rowCount()
            
            if receipt_row == 0:
                self.T_receipt.insertRow(receipt_row)
                total_price = f"{float(self.T_products.item(row, 1).text()) * number:.2f}"
                chkBoxItem = QTableWidgetItem(self.T_products.item(row, 0).text())
                chkBoxItem.setText(self.T_products.item(row, 0).text())
                chkBoxItem.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                if self.CB_category_selection.currentText() in self.list_category_to_send:
                    chkBoxItem.setCheckState(Qt.CheckState.Checked)
                else: chkBoxItem.setCheckState(Qt.CheckState.Unchecked)
                self.T_receipt.setItem(receipt_row, 0, chkBoxItem)
                self.T_receipt.setItem(receipt_row, 1, QTableWidgetItem(f"{number}"))
                self.T_receipt.setItem(receipt_row, 2, QTableWidgetItem(total_price))
                self.category_in_receipt.append(self.CB_category_selection.currentText()) # Aggiunta categoria in lista
            else:
                for product in range(receipt_row): # Controllo se l'articolo esiste già
                    if self.T_products.item(row, 0).text() == self.T_receipt.item(product, 0).text():
                        quantity = int(self.T_receipt.item(product, 1).text())
                        quantity += number
                        total_price = f"{float(self.T_products.item(row, 1).text()) * quantity:.2f}"
                        self.T_receipt.setItem(product, 1, QTableWidgetItem(str(quantity)))
                        self.T_receipt.setItem(product, 2, QTableWidgetItem(str(total_price)))
                        break
                else: # Se non esiste
                    self.T_receipt.insertRow(receipt_row)
                    total_price = f"{float(self.T_products.item(row, 1).text()) * number:.2f}"
                    chkBoxItem = QTableWidgetItem(self.T_products.item(row, 0).text())
                    chkBoxItem.setText(self.T_products.item(row, 0).text())
                    chkBoxItem.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                    if self.CB_category_selection.currentText() in self.list_category_to_send:
                        chkBoxItem.setCheckState(Qt.CheckState.Checked)
                    else: chkBoxItem.setCheckState(Qt.CheckState.Unchecked)
                    self.T_receipt.setItem(receipt_row, 0, chkBoxItem)
                    self.T_receipt.setItem(receipt_row, 1, QTableWidgetItem(f"{number}"))
                    self.T_receipt.setItem(receipt_row, 2, QTableWidgetItem(total_price))
                    self.category_in_receipt.append(self.CB_category_selection.currentText()) # Aggiunta categoria in lista
            self.set_total_price() # Aggiornamento prezzo totale
            
        
    # Funzione aggiungi
    
    def add_product(self):
        # Ricerca del prodotto alla pressione del tasto (o con doppio click sinistro del mouse)
        
        row = self.T_products.currentRow()
        receipt_row = self.T_receipt.rowCount()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        # Inserimento nella tabella a destra
        
        if receipt_row == 0:
            self.T_receipt.insertRow(receipt_row)
            chkBoxItem = QTableWidgetItem(self.T_products.item(row, 0).text())
            chkBoxItem.setText(self.T_products.item(row, 0).text())
            chkBoxItem.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            if self.CB_category_selection.currentText() in self.list_category_to_send:
                chkBoxItem.setCheckState(Qt.CheckState.Checked)
            else: chkBoxItem.setCheckState(Qt.CheckState.Unchecked)
            self.T_receipt.setItem(receipt_row, 0, chkBoxItem)
            self.T_receipt.setItem(receipt_row, 1, QTableWidgetItem("1"))
            self.T_receipt.setItem(receipt_row, 2, QTableWidgetItem(self.T_products.item(row, 1).text()))
            self.category_in_receipt.append(self.CB_category_selection.currentText()) # Aggiunta categoria in lista
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
                chkBoxItem = QTableWidgetItem(self.T_products.item(row, 0).text())
                chkBoxItem.setText(self.T_products.item(row, 0).text())
                chkBoxItem.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                if self.CB_category_selection.currentText() in self.list_category_to_send:
                    chkBoxItem.setCheckState(Qt.CheckState.Checked)
                else: chkBoxItem.setCheckState(Qt.CheckState.Unchecked)
                self.T_receipt.setItem(receipt_row, 0, chkBoxItem)
                self.T_receipt.setItem(receipt_row, 1, QTableWidgetItem("1"))
                self.T_receipt.setItem(receipt_row, 2, QTableWidgetItem(self.T_products.item(row, 1).text()))
                self.category_in_receipt.append(self.CB_category_selection.currentText()) # Aggiunta categoria in lista
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
                    self.category_in_receipt.pop(product) # Rimozione dalla lista categorie
                    break
                else:
                    quantity = int(self.T_receipt.item(product, 1).text())
                    quantity -= 1
                    total_price = f"{float(self.T_products.item(row, 1).text()) * quantity:.2f}"
                    self.T_receipt.setItem(product, 1, QTableWidgetItem(str(quantity)))
                    self.T_receipt.setItem(product, 2, QTableWidgetItem(total_price))
                    break 
        
        self.set_total_price() # Aggiornamento prezzo totale
    
    # Funzione crea menu
    
    def create_menu(self):
        self.create_menu_window = CreateMenuWindow(self.T_products.item(self.T_products.currentRow(), 0).text(), self.CB_category_selection.currentText())
        self.create_menu_window.show()
    
    # *-*-* Funzioni della tabella ricevute *-*-*
    
    # Funzione togli una quantità dalla tabella scontrino (click con sinistro del mouse)
    
    def remove_one_from_receipt(self):
        # Ricerca del prodotto alla pressione del tasto
        
        row = self.T_receipt.currentRow()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        if self.T_receipt.item(row, 1).text() == "1": # Rimozione dalla tabella se l'articolo è a quantità 1
            self.category_in_receipt.pop(row) # Rimozione dalla lista categorie
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
        self.category_in_receipt.pop(row) # Rimozione dalla lista categorie
        self.set_total_price()
    
    # *-*-* Funzioni di stampa e salvataggio *-*-*
    
    # Funzione stampa e salvataggio su DB
    
    def print_receipt(self):
        if self.T_receipt.rowCount() == 0: return # Se la tabella scontrino è vuota
        # Invio Ordine
        lines_to_send = ""
        for line in range(self.T_receipt.rowCount()):
            if self.T_receipt.item(line, 0).checkState() == Qt.CheckState.Checked:
                lines_to_send = lines_to_send + f"{line}-"
        if lines_to_send != "":
            lines_to_send = lines_to_send[:-1]
            # Invio ordine al database
            date_time = datetime.now()
            date = date_time.strftime("%Y%m%d")
            time = date_time.strftime("%H%M%S")
            customer = ""
            table = ""
            if self.LE_customer_name.text() != "": customer = self.LE_customer_name.text()
            else: customer = "-"
            if self.SB_table_select.value() != -1: table = f"{self.SB_table_select.value()}"
            else: table = "-"
            order = ""
            lines_list = lines_to_send.split("-")
            for line_list in lines_list:
                description = self.T_receipt.item(int(line_list), 0).text()
                quantity = self.T_receipt.item(int(line_list), 1).text()
                total_price = self.T_receipt.item(int(line_list), 2).text()
                order = order + f"{description} - {lang.msg(language, 42, 'MainWindow')}: {quantity} - {total_price} {lang.msg(language, 18, 'MainWindow')}\n"
                # Se è presente un menu per il prodotto
                dic_key = description + self.category_in_receipt[int(line_list)]
                if dic_key in menu_dict:
                    line = menu_dict[dic_key].split("@nl@")
                    for detailed_product in line:
                        detail = detailed_product.split("@sp@")
                        detailed_description = detail[0]
                        detailed_quantity = detail[1]
                        detailed_price = detail[2]
                        if detailed_quantity != "-": detailed_quantity = f"{float(detailed_quantity) * float(quantity):.2f}"
                        if detailed_price != "-" and detailed_quantity != "-": detailed_price = f"{float(detailed_price) * float(detailed_quantity):.2f}"
                        order = order + f"*-*-* {detailed_description} - {lang.msg(language, 42, 'MainWindow')} {detailed_quantity} - {lang.msg(language, 43, 'MainWindow')} {detailed_price}\n"
            order = order[:-1]
            order_note = ""
            if len(self.TE_additional_note.toPlainText()) > 0: order_note = self.TE_additional_note.toPlainText()
            else: order_note = "-"
            
            try: # Controllo della connessione al database
                col = self.db["orders"]
                col.insert_one({"status":"now_ordered", "date": date, "time": time, "customer": customer, "table": table, "order": order, "order_note": order_note})
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
                err_msg.setText(lang.msg(language, 48, "MainWindow"))
                return err_msg.exec()
            
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
        try: # Controllo della connessione al database
            col = self.db["receipts"]
            col.insert_one({"receipt_date": date, "receipt_time": time, "receipt_products": products})
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
            
        # Stampa scontrino
        printer = self.CB_printer_list.currentText()
        
        printer_string = "" # Stringa da inviare alla stampante
        printer_total = 0.0 # Totale scontrini da inviare alla stampante
        printer_date_time = date_time.strftime("%d/%m/%Y - %H:%M:%S") # Data e ora da inviare alla stampante
        
        printer_string = printer_string + f"\n-*-* {heading} *-*-\n" # Intestazione
        if len(self.LE_party_name.text()) > 0: printer_string = printer_string + f"-*-* {self.LE_party_name.text()} *-*-\n" # Se il nome festa è compilato
        if len(self.LE_customer_name.text()) > 0: printer_string = printer_string + f"\n{lang.msg(language, 40, 'MainWindow')}: {self.LE_customer_name.text()}" # Se il nome cliente è compilato
        if self.SB_table_select.value() != -1: printer_string = printer_string + f"\n{lang.msg(language, 41, 'MainWindow')}: {self.SB_table_select.value()}" # Se il numero tavolo è compilato
        
        printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        
        # Inserimento prodotti da inviare alla stampante
        for row in range(self.T_receipt.rowCount()):
            description = self.T_receipt.item(row, 0).text()
            quantity = self.T_receipt.item(row, 1).text()
            total_price = self.T_receipt.item(row, 2).text()
            printer_total += float(total_price) # Aggiornamento totale scontrino
            printer_string = printer_string + f"\n{description} - {lang.msg(language, 42, 'MainWindow')} {quantity} - {lang.msg(language, 43, 'MainWindow')} {total_price}" # Inserimento prodotto nello scontrino
            # Se è presente un menu per il prodotto
            dic_key = description + self.category_in_receipt[row]
            if dic_key in menu_dict:
                line = menu_dict[dic_key].split("@nl@")
                for detailed_product in line:
                    detail = detailed_product.split("@sp@")
                    detailed_description = detail[0]
                    detailed_quantity = detail[1]
                    detailed_price = detail[2]
                    if detailed_quantity != "-": detailed_quantity = f"{float(detailed_quantity) * float(quantity):.2f}"
                    if detailed_price != "-" and detailed_quantity != "-": detailed_price = f"{float(detailed_price) * float(detailed_quantity):.2f}"
                    printer_string = printer_string + f"\n *-*-* {detailed_description} - {lang.msg(language, 42, 'MainWindow')} {detailed_quantity} - {lang.msg(language, 43, 'MainWindow')} {detailed_price}"
        printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        printer_string = printer_string + f"\n{lang.msg(language, 17, 'MainWindow')} {printer_total:.2f} {lang.msg(language, 43, 'MainWindow')}"
        printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        
        if len(self.TE_additional_note.toPlainText()) > 0: # Se le note sono compilate
            printer_string = printer_string + f"\n{self.TE_additional_note.toPlainText()}"
            printer_string = printer_string + "\n\n---------------------------------" # Divisorio
        printer_string = printer_string + f"\n\n-*-* {printer_date_time} *-*-" # Data e ora scontrino
        printer_string = printer_string + "\n\n\n\n\n\n\n\n\n\n\n" # Fine scontrino
        
        # Pulizia caselle, tabella e lista
        
        self.LE_customer_name.clear()
        self.SB_table_select.setValue(-1)
        self.TE_additional_note.clear()
        self.category_in_receipt.clear()
        for row in reversed(range(self.T_receipt.rowCount())):
            self.T_receipt.removeRow(row)
        self.set_total_price()
        
        # Debug mode
        
        if debug_mode == True and printer == "DEBUG":
            print(printer_string)
            return
        
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
        # Invio Ordine
        lines_to_send = ""
        for line in range(self.T_receipt.rowCount()):
            if self.T_receipt.item(line, 0).checkState() == Qt.CheckState.Checked:
                lines_to_send = lines_to_send + f"{line}-"
        if lines_to_send != "":
            lines_to_send = lines_to_send[:-1]
            # Invio ordine al database
            date_time = datetime.now()
            date = date_time.strftime("%Y%m%d")
            time = date_time.strftime("%H%M%S")
            customer = ""
            table = ""
            if self.LE_customer_name.text() != "": customer = self.LE_customer_name.text()
            else: customer = "-"
            if self.SB_table_select.value() != -1: table = f"{self.SB_table_select.value()}"
            else: table = "-"
            order = ""
            lines_list = lines_to_send.split("-")
            for line_list in lines_list:
                description = self.T_receipt.item(int(line_list), 0).text()
                quantity = self.T_receipt.item(int(line_list), 1).text()
                total_price = self.T_receipt.item(int(line_list), 2).text()
                order = order + f"{description} - {lang.msg(language, 42, 'MainWindow')}: {quantity} - {total_price} {lang.msg(language, 18, 'MainWindow')}\n"
                # Se è presente un menu per il prodotto
                dic_key = description + self.category_in_receipt[int(line_list)]
                if dic_key in menu_dict:
                    line = menu_dict[dic_key].split("@nl@")
                    for detailed_product in line:
                        detail = detailed_product.split("@sp@")
                        detailed_description = detail[0]
                        detailed_quantity = detail[1]
                        detailed_price = detail[2]
                        if detailed_quantity != "-": detailed_quantity = f"{float(detailed_quantity) * float(quantity):.2f}"
                        if detailed_price != "-" and detailed_quantity != "-": detailed_price = f"{float(detailed_price) * float(detailed_quantity):.2f}"
                        order = order + f"*-*-* {detailed_description} - {lang.msg(language, 42, 'MainWindow')} {detailed_quantity} - {lang.msg(language, 43, 'MainWindow')} {detailed_price}\n"
            order = order[:-1]
            order_note = ""
            if len(self.TE_additional_note.toPlainText()) > 0: order_note = self.TE_additional_note.toPlainText()
            else: order_note = "-"
            
            try: # Controllo della connessione al database
                col = self.db["orders"]
                col.insert_one({"status":"now_ordered", "date": date, "time": time, "customer": customer, "table": table, "order": order, "order_note": order_note})
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
                err_msg.setText(lang.msg(language, 48, "MainWindow"))
                return err_msg.exec()
            
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
        try: # Controllo della connessione al database
            col = self.db["receipts"]
            col.insert_one({"receipt_date": date, "receipt_time": time, "receipt_products": products})
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
        
        # Pulizia caselle, tabella e lista
        
        self.LE_customer_name.clear()
        self.SB_table_select.setValue(-1)
        self.TE_additional_note.clear()
        self.category_in_receipt.clear()
        for row in reversed(range(self.T_receipt.rowCount())):
            self.T_receipt.removeRow(row)
        self.set_total_price()
    
    # *-*-* Funzione pressione tasti *-*-*
    
    def keyPressEvent(self, event):
        #print(event.key())
        pass
    
    # *-*-* Funzione apertura finestra database *-*-*
    
    def database_open(self):
        self.database_window = DatabaseWindow()
        self.database_window.show()
    
    # *-*-* Funzione apertura finestra opzioni *-*-*
    
    def options_menu_open(self):
        self.options_window = OptionsMenu()
        self.options_window.show()
        self.close()
    
    # *-*-* Funzione apertura finestra ordini *-*-*
    
    def open_orders_table(self):
        self.orders_window = OrdersWindow()
        self.orders_window.show()

# *-*-* Finestra Ordini *-*-*

class OrdersWindow(QWidget):
    def __init__(self):
        super().__init__()
        
        self.db = dbclient["Bar"] # Apertura database
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"{lang.msg(language, 0, 'OrdersWindow')} {heading}")
        self.setMinimumSize(640, 480)
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        self.setStyleSheet(sis.interface_style(interface))
        
        # Tabella ordini
        
        self.T_orders = QTableWidget(self)
        self.T_orders.setColumnCount(4)
        self.T_orders.setHorizontalHeaderLabels(["ID", lang.msg(language, 10, "MainWindow"), lang.msg(language, 1, "OrdersWindow"), lang.msg(language, 2, "OrdersWindow")])
        self.T_orders.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_orders.doubleClicked.connect(self.show_detailed_order)
        self.T_orders.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_orders.customContextMenuRequested.connect(self.T_orders_CM)
        self.T_orders_headers = self.T_orders.horizontalHeader()
        self.lay.addWidget(self.T_orders, 0, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
        
        # Pulsante aggiornamento tabella
        
        self.B_update = QPushButton(self, text=lang.msg(language, 10, "OrdersWindow"))
        self.B_update.clicked.connect(self.search_orders)
        self.lay.addWidget(self.B_update, 1, 0, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Avvio ricerca ordini
        self.search_orders()
        
    def resizeEvent(self, event):
        W_width = self.width()
        W_height = self.height()
        
        try:
            # Tabella ordini
            self.T_orders.setMinimumSize(int(W_width - 15), int((W_height - 65)))
            self.T_orders_headers.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            self.T_orders_headers.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            self.T_orders_headers.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            self.T_orders_headers.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        except AttributeError:
            pass
    
    # *-*-* Ricerca ordini *-*-*
    
    def search_orders(self):
        # Rimozione righe iniziale
        if self.T_orders.rowCount() != 0:
            for row in reversed(range(self.T_orders.rowCount())):
                self.T_orders.removeRow(row)
        
        try: # Controllo della connessione al database
            col = self.db["orders"]
            for line in col.find().sort("_id", pymongo.DESCENDING):
                date_time = f"{line['date'][6:]}/{line['date'][4:6]}/{line['date'][:4]} - {line['time'][:2]}:{line['time'][2:4]}:{line['time'][4:]}"
                status_st = ""
                if line["status"] == "ordered" or line["status"] == "now_ordered": status_st = lang.msg(language, 3, "OrdersWindow")
                if line["status"] == "in_progress": status_st = lang.msg(language, 4, "OrdersWindow")
                if line["status"] == "done": status_st = lang.msg(language, 5, "OrdersWindow")
                row = self.T_orders.rowCount()
                self.T_orders.insertRow(row)
                self.T_orders.setItem(row, 0, QTableWidgetItem(str(line["_id"])))
                self.T_orders.setItem(row, 1, QTableWidgetItem(f"{lang.msg(language, 10, 'MainWindow')}: {line['customer']}"))
                self.T_orders.setItem(row, 2, QTableWidgetItem(date_time))
                self.T_orders.setItem(row, 3, QTableWidgetItem(status_st))
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
    
    # *-*-* Visualizzazione dettagliata degli ordini *-*-*
    
    def show_detailed_order(self):
        row = self.T_orders.currentRow()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        try: # Controllo della connessione al database
            col = self.db["orders"]
            obj_instance = ObjectId(self.T_orders.item(row, 0).text())
            order = col.find_one({"_id": obj_instance})
            order_string = f"""ID: {order['_id']}
{lang.msg(language, 10, "MainWindow")}: {order["customer"]} - {lang.msg(language, 41, "MainWindow")}: {order["table"]}
{lang.msg(language, 7, "OrdersWindow")}: {order['date'][6:]}/{order['date'][4:6]}/{order['date'][:4]} - {lang.msg(language, 8, "OrdersWindow")}: {order['time'][:2]}:{order['time'][2:4]}:{order['time'][4:]}
-----------------------------------
{order["order"]}
-----------------------------------
{order["order_note"]}"""
            msg = QMessageBox(self)
            msg.setWindowTitle(lang.msg(language, 6, "OrdersWindow"))
            msg.setText(order_string)
            return msg.exec()
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
    
    # *-*-* Cambio stato ordine *-*-*
    
    def change_status(self, status:str):
        row = self.T_orders.currentRow()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        status_st = ""
        if status == "ordered": status_st = lang.msg(language, 3, "OrdersWindow")
        if status == "in_progress": status_st = lang.msg(language, 4, "OrdersWindow")
        if status == "done": status_st = lang.msg(language, 5, "OrdersWindow")
        
        try: # Controllo della connessione al database
            col = self.db["orders"]
            obj_instance = ObjectId(self.T_orders.item(row, 0).text())
            col.update_one({"_id": obj_instance},{"$set": {"status": status}})
            self.T_orders.setItem(row, 3, QTableWidgetItem(status_st))
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
    
    # *-*-* Eliminazione ordine *-*-*
    
    def delete_order(self):
        row = self.T_orders.currentRow()
        
        if row == -1: return # Blocco delle funzioni se nulla è selezionato
        
        try: # Controllo della connessione al database
            col = self.db["orders"]
            obj_instance = ObjectId(self.T_orders.item(row, 0).text())
            
            col.delete_one({"_id": obj_instance})
            self.T_orders.removeRow(row)
            self.T_orders.setCurrentCell(-1, -1)
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
    
    # *-*-* Custom Menu della tabella *-*-*
    
    def T_orders_CM(self):
        if self.T_orders.currentRow() == -1: return
        menu = QMenu(self)
        sdo_action = QAction(lang.msg(language, 9, "OrdersWindow"), self)
        cs_ordered_action = QAction(lang.msg(language, 11, "OrdersWindow"), self)
        cs_inprogress_action = QAction(lang.msg(language, 12, "OrdersWindow"), self)
        cs_done_action = QAction(lang.msg(language, 13, "OrdersWindow"), self)
        delete_action = QAction(lang.msg(language, 26, "MainWindow"), self)
        sdo_action.triggered.connect(self.show_detailed_order)
        cs_ordered_action.triggered.connect(lambda x: self.change_status("ordered"))
        cs_inprogress_action.triggered.connect(lambda x: self.change_status("in_progress"))
        cs_done_action.triggered.connect(lambda x: self.change_status("done"))
        delete_action.triggered.connect(self.delete_order)
        menu.addActions([sdo_action,cs_ordered_action,cs_inprogress_action,cs_done_action,delete_action])
        menu.popup(QCursor.pos())

# *-*-* Finestra Database *-*-*

class DatabaseWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.db = dbclient["Bar"] # Apertura database
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"Database {heading}")
        self.setFixedSize(540, 640)
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        self.setStyleSheet(sis.interface_style(interface))
        
        # Calendario
        
        self.date_selection = 0
        self.format = QTextCharFormat()
        if interface == "98 Style":
            self.format.setBackground(QColor("yellow"))
            self.format.setForeground(QColor("black"))
        if interface == "Tech Style":
            self.format.setBackground(QColor("#2DB00D"))
            self.format.setForeground(QColor("black"))
        if interface == "Clear Elegant Style":
            self.format.setBackground(QColor("#044B11"))
            self.format.setForeground(QColor("black"))
        if interface == "Dark Elegant Style":
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
        
        L_time_range = QLabel(self, text=lang.msg(language, 0, "DatabaseWindow"))
        self.lay.addWidget(L_time_range, 0, 1, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Casella selezione orari
        
        self.LE_time_range = QLineEdit(self)
        self.LE_time_range.setFixedWidth(250)
        self.LE_time_range.setPlaceholderText(lang.msg(language, 1, "DatabaseWindow"))
        self.lay.addWidget(self.LE_time_range, 1, 1, 1, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Bottone interroga
        
        self.B_query_database = QPushButton(self, text=lang.msg(language, 2, "DatabaseWindow"))
        self.B_query_database.clicked.connect(self.query_database)
        self.lay.addWidget(self.B_query_database, 2, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Label ordinamento dizionario
        
        L_sort_orders = QLabel(self, text=lang.msg(language, 17, "DatabaseWindow"))
        self.lay.addWidget(L_sort_orders, 2, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Bottone elimina dal database
        
        self.B_delete_database = QPushButton(self, text=lang.msg(language, 26, "MainWindow"))
        self.B_delete_database.clicked.connect(self.delete_database)
        self.lay.addWidget(self.B_delete_database, 3, 1, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Combobox ordinamento dizionario
        
        self.CB_sort_orders = QComboBox(self)
        self.CB_sort_orders.addItems([lang.msg(language, 18, "DatabaseWindow"),lang.msg(language, 19, "DatabaseWindow"),lang.msg(language, 20, "DatabaseWindow"),
                                      lang.msg(language, 21, "DatabaseWindow"),lang.msg(language, 22, "DatabaseWindow"),lang.msg(language, 23, "DatabaseWindow")])
        self.lay.addWidget(self.CB_sort_orders, 3, 2, Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        
        # Text Edit per risposta Database
        
        self.TE_database_response = QTextEdit(self)
        self.TE_database_response.setReadOnly(True)
        self.TE_database_response.setFixedHeight(450)
        self.lay.addWidget(self.TE_database_response, 4, 0, 1, 3, Qt.AlignmentFlag.AlignTop)
    
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
        self.TE_database_response.insertPlainText(f"-*-* {lang.msg(language, 3, 'DatabaseWindow')} *-*-\n\n")
        
        try: # Controllo della connessione al database
            if self.second_date == "": # Se è selezionata una sola data
                firstdate = f"{self.first_date[6:]}/{self.first_date[4:6]}/{self.first_date[:4]}"
                if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                    for products in col.find({"receipt_date": self.first_date}, {"_id": 0}):
                        result = 1
                        self.TE_database_response.append(f"\n{lang.msg(language, 4, 'DatabaseWindow')}: {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]}\n{lang.msg(language, 5, 'DatabaseWindow')}: {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
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
                        self.TE_database_response.setPlainText(lang.msg(language, 6, "DatabaseWindow"))
                        return
                
                else: # Se è stato selezionato un orario
                    time_string = self.LE_time_range.text().replace(" ", "")
                    if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                        self.TE_database_response.clear()
                        self.TE_database_response.setPlainText(lang.msg(language, 7, "DatabaseWindow"))
                        return
                    time_string = time_string.replace(":", "-")
                    time_string = time_string.split("-")
                    first_time = f"{time_string[0]}{time_string[1]}00"
                    second_time = f"{time_string[2]}{time_string[3]}00"
                    if int(first_time) > int(second_time):
                        self.TE_database_response.clear()
                        self.TE_database_response.setPlainText(lang.msg(language, 7, "DatabaseWindow"))
                        return
                    for products in col.find({"receipt_date": self.first_date, "receipt_time": {"$gte": first_time, "$lte": second_time}}, {"_id": 0}):
                        result = 1
                        self.TE_database_response.append(f"\n{lang.msg(language, 4, 'DatabaseWindow')}: {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]}\n{lang.msg(language, 5, 'DatabaseWindow')}: {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
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
                        self.TE_database_response.setPlainText(lang.msg(language, 8, "DatabaseWindow"))
                        return
                        
                text_cursor = QTextCursor(self.TE_database_response.document()) # Spostamento del cursore all'inizio
                text_cursor.setPosition(0)
                self.TE_database_response.setTextCursor(text_cursor)
                total_string = f"-*-* {lang.msg(language, 9, 'DatabaseWindow')} {firstdate} *-*-\n"
                if self.CB_sort_orders.currentIndex() == 0: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["total"])) # Ordinato per prezzo
                if self.CB_sort_orders.currentIndex() == 1: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["quantity"])) # Ordinato per quantità
                if self.CB_sort_orders.currentIndex() == 2: detail_tot = dict(sorted(detail_tot.items())) # Ordinato per descrizione
                if self.CB_sort_orders.currentIndex() == 3: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["total"], reverse=True)) # Ordinato per prezzo (contrario)
                if self.CB_sort_orders.currentIndex() == 4: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["quantity"], reverse=True)) # Ordinato per quantità (contrario)
                if self.CB_sort_orders.currentIndex() == 5: detail_tot = dict(sorted(detail_tot.items(), reverse=True)) # Ordinato per descrizione (contrario)
                for detail in detail_tot: # Loop del dizionario con il totale vendite
                    total_string = total_string + f"\n{detail} - {lang.msg(language, 42, 'MainWindow')} {detail_tot[detail]['quantity']} - {lang.msg(language, 18, 'MainWindow')} {detail_tot[detail]['total']:.2f}"
                total_string = total_string + f"\n\n-*-* {lang.msg(language, 10, 'DatabaseWindow')} {lang.msg(language, 18, 'MainWindow')} {total_receipts:.2f} *-*-\n --------------------------\n\n\n"
                self.TE_database_response.insertPlainText(total_string)
            
            else: # Se sono state selezionate 2 date
                firstdate = f"{self.first_date[6:]}/{self.first_date[4:6]}/{self.first_date[:4]}"
                seconddate = f"{self.second_date[6:]}/{self.second_date[4:6]}/{self.second_date[:4]}"
                if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                    for products in col.find({"receipt_date": {"$gte": self.first_date, "$lte": self.second_date}}, {"_id": 0}):
                        result = 1
                        self.TE_database_response.append(f"\n{lang.msg(language, 4, 'DatabaseWindow')}: {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]}\n{lang.msg(language, 5, 'DatabaseWindow')}: {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
                        for product in str(products["receipt_products"]).split("@newline@"):
                            detail = product.split("@space@")
                            self.TE_database_response.append(f"\n{detail[0]} - {lang.msg(language, 42, 'MainWindow')} {detail[1]} - {lang.msg(language, 18, 'MainWindow')} {detail[2]}")
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
                        self.TE_database_response.setPlainText(lang.msg(language, 11, "DatabaseWindow"))
                        return
                
                else: # Se è stato selezionato un orario
                    time_string = self.LE_time_range.text().replace(" ", "")
                    if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                        self.TE_database_response.clear()
                        self.TE_database_response.setPlainText(lang.msg(language, 7, "DatabaseWindow"))
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
                            
                        self.TE_database_response.append(f"\n{lang.msg(language, 4, 'DatabaseWindow')}: {str(products['receipt_date'])[6:]}/{str(products['receipt_date'])[4:6]}/{str(products['receipt_date'])[:4]}\n{lang.msg(language, 5, 'DatabaseWindow')} {str(products['receipt_time'])[:2]}:{str(products['receipt_time'])[2:4]}:{str(products['receipt_time'])[4:]}\n")
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
                        self.TE_database_response.setPlainText(lang.msg(language, 12, "DatabaseWindow"))
                        return
                
                text_cursor = QTextCursor(self.TE_database_response.document()) # Spostamento del cursore all'inizio
                text_cursor.setPosition(0)
                self.TE_database_response.setTextCursor(text_cursor)
                total_string = f"-*-* {lang.msg(language, 13, 'DatabaseWindow')} {firstdate} {lang.msg(language, 14, 'DatabaseWindow')} {seconddate} *-*-\n"
                if self.CB_sort_orders.currentIndex() == 0: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["total"])) # Ordinato per prezzo
                if self.CB_sort_orders.currentIndex() == 1: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["quantity"])) # Ordinato per quantità
                if self.CB_sort_orders.currentIndex() == 2: detail_tot = dict(sorted(detail_tot.items())) # Ordinato per descrizione
                if self.CB_sort_orders.currentIndex() == 3: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["total"], reverse=True)) # Ordinato per prezzo (contrario)
                if self.CB_sort_orders.currentIndex() == 4: detail_tot = dict(sorted(detail_tot.items(), key=lambda item: item[1]["quantity"], reverse=True)) # Ordinato per quantità (contrario)
                if self.CB_sort_orders.currentIndex() == 5: detail_tot = dict(sorted(detail_tot.items(), reverse=True)) # Ordinato per descrizione (contrario)
                for detail in detail_tot: # Loop del dizionario con il totale vendite
                    total_string = total_string + f"\n{detail} - {lang.msg(language, 42, 'MainWindow')} {detail_tot[detail]['quantity']} - {lang.msg(language, 18, 'MainWindow')} {detail_tot[detail]['total']:.2f}"
                total_string = total_string + f"\n\n-*-* {lang.msg(language, 10, 'DatabaseWindow')} {lang.msg(language, 18, 'MainWindow')} {total_receipts:.2f} *-*-\n --------------------------\n\n\n"
                self.TE_database_response.insertPlainText(total_string)
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
    
    # Eliminazione dal database
    
    def delete_database(self):
        msg = QMessageBox(self)
        msg.setWindowTitle(lang.msg(language, 21, "MainWindow"))
        msg.setText(lang.msg(language, 15, "DatabaseWindow"))
        msg.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.No)
        msg.buttonClicked.connect(self.delete_database_confirm)
        return msg.exec()
    
    def delete_database_confirm(self, button):
        if button.text() == "&OK" or button.text() == "OK":
            try: # Controllo della connessione al database
                col = self.db["receipts"]
                if self.second_date == "": # Se è selezionata una sola data
                    if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                        col.delete_many({"receipt_date": self.first_date})
                    else: # Se è stato selezionato un orario
                        time_string = self.LE_time_range.text().replace(" ", "")
                        if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                            self.TE_database_response.clear()
                            self.TE_database_response.setPlainText(lang.msg(language, 7, "DatabaseWindow"))
                            return
                        time_string = time_string.replace(":", "-")
                        time_string = time_string.split("-")
                        first_time = f"{time_string[0]}{time_string[1]}00"
                        second_time = f"{time_string[2]}{time_string[3]}00"
                        if int(first_time) > int(second_time):
                            self.TE_database_response.clear()
                            self.TE_database_response.setPlainText(lang.msg(language, 7, "DatabaseWindow"))
                            return
                        col.delete_many({"receipt_date": self.first_date, "receipt_time": {"$gte": first_time, "$lte": second_time}})
                else: # Se sono state selezionate 2 date
                    if len(self.LE_time_range.text()) == 0: # Se non è stato selezionato un orario
                        col.delete_many({"receipt_date": {"$gte": self.first_date, "$lte": self.second_date}})
                    else: # Se è stato selezionato un orario
                        time_string = self.LE_time_range.text().replace(" ", "")
                        if len(time_string) != 11 or time_string.count("-") != 1 or time_string.count(":") != 2: # Controllo della stringa orario
                            self.TE_database_response.clear()
                            self.TE_database_response.setPlainText(lang.msg(language, 7, "DatabaseWindow"))
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
                self.TE_database_response.insertPlainText(f"-*-* {lang.msg(language, 16, 'DatabaseWindow')} *-*-\n\n")
            except:
                err_msg = QMessageBox(self)
                err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
                err_msg.setText(lang.msg(language, 48, "MainWindow"))
                return err_msg.exec()

# -*-* Creazione menu articolo *-*-

class CreateMenuWindow(QWidget):
    def __init__(self, product:str, category:str):
        super().__init__()
        self.db = dbclient["Bar"] # Apertura database
        
        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle(f"Menu {heading}")
        self.setFixedSize(480, 640)
        self.lay = QGridLayout(self)
        self.setLayout(self.lay)
        self.lay.setContentsMargins(10,10,10,10)
        self.lay.setSpacing(1)
        self.setStyleSheet(sis.interface_style(interface))
        self.product = product
        self.category = category
        self.menu_product = product + category
        
        L_product = QLabel(self, text=f"{lang.msg(language, 0, 'CreateMenuWindow')} {self.product}")
        L_product.setAccessibleName("an_section_title")
        self.lay.addWidget(L_product, 0, 0, 1, 3, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.LE_description = QLineEdit(self)
        self.LE_description.setPlaceholderText(lang.msg(language, 2, "MainWindow"))
        self.lay.addWidget(self.LE_description, 1, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.LE_quantity = QLineEdit(self)
        self.LE_quantity.setPlaceholderText(lang.msg(language, 16, "MainWindow"))
        self.lay.addWidget(self.LE_quantity, 1, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.LE_price = QLineEdit(self)
        self.LE_price.setPlaceholderText(lang.msg(language, 1, "CreateMenuWindow"))
        self.lay.addWidget(self.LE_price, 1, 2, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.B_insert = QPushButton(self, text=lang.msg(language, 6, "MainWindow"))
        self.B_insert.clicked.connect(self.insert_menu)
        self.lay.addWidget(self.B_insert, 2, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.B_create = QPushButton(self, text=lang.msg(language, 3, "MainWindow"))
        self.B_create.clicked.connect(self.create_menu)
        self.lay.addWidget(self.B_create, 2, 1, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.B_delete = QPushButton(self, text=lang.msg(language, 26, "MainWindow"))
        self.B_delete.clicked.connect(self.delete_menu)
        self.lay.addWidget(self.B_delete, 2, 2, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        self.T_menu = QTableWidget(self)
        self.T_menu.setColumnCount(3)
        self.T_menu.setFixedSize(450, 500)
        self.T_menu.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.T_menu.setHorizontalHeaderLabels([lang.msg(language, 2, "MainWindow"), lang.msg(language, 16, "MainWindow"), lang.msg(language, 1, "CreateMenuWindow")])
        self.T_menu.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.T_menu.customContextMenuRequested.connect(self.T_menu_CM)
        self.T_menu_header = self.T_menu.horizontalHeader()
        self.T_menu_header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.T_menu_header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.T_menu_header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.lay.addWidget(self.T_menu, 3, 0, 1, 3, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        
        # Controllo dizionario menu e aggiunta alla tabella
        
        if self.menu_product in menu_dict:
            lines = menu_dict[self.menu_product].split("@nl@")
            for line in lines:
                detail = line.split("@sp@")
                detailed_product = detail[0]
                try: detailed_quantity = f"{float(detail[1]):.2f}"
                except: detailed_quantity = "-"
                try: detailed_price = f"{float(detail[2]):.2f}"
                except: detailed_price = "-"
                row = self.T_menu.rowCount()
                self.T_menu.insertRow(row)
                self.T_menu.setItem(row, 0, QTableWidgetItem(detailed_product))
                self.T_menu.setItem(row, 1, QTableWidgetItem(detailed_quantity))
                self.T_menu.setItem(row, 2, QTableWidgetItem(detailed_price))
                   
    # -*-* Funzione inserimento prodotto *-*-
    
    def insert_menu(self):
        if self.LE_description.text().strip().upper() == "": # Controllo casella descrizione
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 20, "MainWindow"))
            return err_msg.exec()
        
        description = self.LE_description.text().strip().upper()
        quantity = self.LE_quantity.text().strip()
        price = self.LE_price.text().strip()
        
        if quantity == "": quantity = "-" # Controllo quantità
        else:
            try: quantity = float(quantity)
            except: quantity = "-"
        if price == "": price = "-" # Controllo prezzo
        else:
            try: price = float(price)
            except: price = "-"
        
        # Inserimento nella tabella
        row = self.T_menu.rowCount()
        self.T_menu.insertRow(row)
        self.T_menu.setItem(row, 0, QTableWidgetItem(description))
        if type(quantity) == float: self.T_menu.setItem(row, 1, QTableWidgetItem(f"{quantity:.2f}"))
        else: self.T_menu.setItem(row, 1, QTableWidgetItem(f"{quantity}"))
        if type(price) == float: self.T_menu.setItem(row, 2, QTableWidgetItem(f"{price:.2f}"))
        else: self.T_menu.setItem(row, 2, QTableWidgetItem(f"{price}"))
        
        # Pulizia caselle
        
        self.LE_description.clear()
        self.LE_quantity.clear()
        self.LE_price.clear()
    
    # -*-* Funzione creazione menu *-*-
    
    def create_menu(self):
        if self.T_menu.rowCount() == 0:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 2, "CreateMenuWindow"))
            return err_msg.exec()
        
        # Preparazione stringa
        
        menu_string = ""
        for row in range(self.T_menu.rowCount()):
            menu_string = menu_string + f"{self.T_menu.item(row, 0).text()}@sp@{self.T_menu.item(row, 1).text()}@sp@{self.T_menu.item(row, 2).text()}@nl@"
        menu_string = menu_string[:-4]
        
        # Inserimento nel database
        
        global menu_dict
        try: # Controllo della connessione al database
            col = self.db["maincategory"]
            col.update_one({"description": self.product, "category": self.category}, {"$set": {"menu": menu_string}}, upsert=True)
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
        
        # Inserimento nel dizionario
        menu_dict.update({self.menu_product: menu_string})
        
        # Chiusura finestra
        
        self.close()
    
    # -*-* Funzione eliminazione menu *-*-
    
    def delete_menu(self):
        global menu_dict
        if self.menu_product not in menu_dict: # Controllo se il menu non è presente
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 3, "CreateMenuWindow"))
            return err_msg.exec()
        
        # Rimozione dal database
        
        try: # Controllo della connessione al database
            col = self.db["maincategory"]
            col.update_one({"description": self.product, "category": self.category}, {"$unset": {"menu": ""}})
        except:
            err_msg = QMessageBox(self)
            err_msg.setWindowTitle(lang.msg(language, 19, "MainWindow"))
            err_msg.setText(lang.msg(language, 48, "MainWindow"))
            return err_msg.exec()
        
        # Rimozione dal dizionario
        
        menu_dict.pop(self.menu_product)
        
        # Chiusura finestra
        
        self.close()

    # -*-* Funzioni del menu tabella prodotti *-*-
    
    def T_menu_CM(self):
        if self.T_menu.currentRow() == -1: return
        menu = QMenu(self)
        delete_action = QAction(lang.msg(language, 26, "MainWindow"), self)
        delete_action.triggered.connect(self.delete_product)
        menu.addAction(delete_action)
        menu.popup(QCursor.pos())
    
    # Eliminazione prodotto
    
    def delete_product(self):
        self.T_menu.removeRow(self.T_menu.currentRow())
        self.T_menu.setCurrentCell(-1, -1)

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
    
        if os.path.exists(f"{os.path.expanduser('~')}/Orders/options.txt"):
            options_file = open(f"{os.path.expanduser('~')}/Orders/options.txt", "r")
            self.mongodb_connection = options_file.readline().replace("db_connection=", "").replace("\n", "")
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
        
        options_file = open(f"{os.path.expanduser('~')}/Orders/options.txt", "w")
        options_file.write(f"db_connection={self.LE_database_connection.text()}\nheading={self.LE_heading.text()}\ninterface={self.CB_interface_style.currentText()}\nlogo={self.logo_path}\nicon={self.icon_path}\nlanguage={self.CB_language.currentText()}")
        options_file.close()
        
        # -*-* Riavvio applicazione *-*-
        
        # Lettura file e impostazione variabili
    
        options_file = open(f"{os.path.expanduser('~')}/Orders/options.txt", "r")
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
            
            # Inserimento prodotti nel dizionario

            db = dbclient["Bar"]
            col = db["maincategory"]
            global menu_dict
            for product in col.find():
                try:
                    st_menu = str(product["menu"])
                    dic_product = str(product["description"]) + str(product["category"])
                    menu_dict.update({dic_product: st_menu})
                except: pass
            
            self.window = MainWindow()
            self.window.show()
            self.close()
        except:
            self.L_database_connection_st.setStyleSheet("color: #8B0B0B; font: 18px bold Arial;")
            self.L_database_connection_st.setText(lang.msg(self.language, 18, "OptionsMenuWindow"))
            self.L_database_connection_st.show()

# -*-* Avvio applicazione *-*-

if os.path.exists(f"{os.path.expanduser('~')}/Orders") == False: os.mkdir(f"{os.path.expanduser('~')}/Orders")

if os.path.exists(f"{os.path.expanduser('~')}/Orders/options.txt") == False: first_start_application = 1
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = OptionsMenu()
    window.show()
    app.exec()