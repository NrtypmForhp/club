from kivymd.app import MDApp
from kivy.lang import Builder
from kivy.clock import Clock
from kivymd.uix.list import ThreeLineListItem
from kivy.core.audio import SoundLoader
from kivy.properties import StringProperty
import pymongo, certifi, os
from kivy.utils import platform
from bson.objectid import ObjectId

# Versione 1.0.0 r5

# Variabili Globali

# Dimensione testo

title_size = "[size=35dp]"
list_title_size = "[size=25dp]"
normal_size = "[size=20dp]"

mongodb_connection = "mongodb://localhost:27017/" # Stringa di connessione al database
language = "IT" # Stringa di selezione lingua
heading = "98 Ottani The Club" # Stringa di selezione titolo
update_time = 30.0 # Tempo di update del database
if platform == "android":
    from android.storage import primary_external_storage_path
    from android.permissions import request_permissions, Permission, check_permission

# Lingue

def lang(lg:str, index:int):
    it = ["0@-@INDIETRO","1@-@CAMBIA STATO IN -> ","2@-@Connessione al database fallita!","3@-@Connessione al database in corso....","4@-@ORDINATO","5@-@IN CORSO",
          "6@-@Stato","7@-@ORDINE FATTO","8@-@ELIMINA ORDINE","9@-@Nome cliente","10@-@Tavolo","11@-@Data ordine","12@-@Orario ordine"]
    en = ["0@-@BACK","1@-@CHANGE STATUS IN -> ","2@-@Connection to database failed!","3@-@Connection to database....","4@-@ORDERED","5@-@IN PROGRESS",
          "6@-@Status","7@-@ORDER DONE","8@-@DELETE ORDER","9@-@Customer name","10@-@Table","11@-@Order date","12@-@Order time"]
    if lg == "IT": return it[index][it[index].index("@-@")+3:]
    if lg == "EN": return en[index][en[index].index("@-@")+3:]

# Applicazione

class MainWindow(MDApp):
    def scheduled_function(self, dt): # Funzione che controlla il database periodicamente
        if self.root.current == "orders_table":
            self.root.ids.LS_products.clear_widgets()
            col = self.db["orders"]
            for line in col.find().sort("_id", pymongo.DESCENDING):
                status_st = ""
                color_st = ""
                if line["status"] == "now_ordered":
                    status_st = lang(language, 4)
                    color_st = "#BA681F"
                    obj_instance = ObjectId(line["_id"])
                    col.update_one({"_id": obj_instance}, {"$set": {"status": "ordered"}})
                    if self.sound != "-": self.sound.play()
                if line["status"] == "ordered":
                    status_st = lang(language, 4)
                    color_st = "#BA681F"
                if line["status"] == "in_progress":
                    status_st = lang(language, 5)
                    color_st = "#08506D"
                if line["status"] == "done":
                    status_st = lang(language, 7)
                    color_st = "#158B09"
                customer_and_table = f"{lang(language, 9)}: {line['customer']} - {lang(language, 10)}: {line['table']}"
                date_and_time = f"{lang(language, 11)}: {line['date'][6:]}/{line['date'][4:6]}/{line['date'][:4]} - {lang(language, 12)}: {line['time'][:2]}:{line['time'][2:4]}:{line['time'][4:]}"
                list_item = ThreeLineListItem(text=f"{list_title_size}ID: {line['_id']}[/size]",
                                              secondary_text=f"{normal_size}{customer_and_table} - {date_and_time}[/size]",
                                              tertiary_text=f"{normal_size}{lang(language, 6)}: {status_st}[/size]",
                                              theme_text_color="Custom", text_color="#B0B006",
                                              secondary_theme_text_color="Custom", secondary_text_color="#B0B006",
                                              tertiary_theme_text_color="Custom", tertiary_text_color=color_st)
                list_item.bind(on_release = lambda x, item=list_item: self.show_detailed_order(item))
                self.root.ids.LS_products.add_widget(list_item)
                obj_instance = ObjectId(line["_id"])
    
    def scheduled_database_connection_1(self, dt): # Impostazione schermata iniziale
        self.root.current = "database_connection"
        Clock.schedule_once(self.scheduled_database_connection_2, 1.0) # Avvio della connessione al database
    
    def scheduled_database_connection_2(self, dt): # Connessione al database e caricamento variabili
        try:
            if platform == "android": self.dbclient = pymongo.MongoClient(mongodb_connection, tlsCAFile=certifi.where()) # Solo per Android
            else: self.dbclient = pymongo.MongoClient(mongodb_connection)
            self.dbclient.server_info() # Test per la connessione al database
            self.title = heading
            self.db = self.dbclient["Bar"]
            self.root.current = "orders_table"
        except:
            self.root.current = "database_connection"
            self.root.ids["L_connection"].text = lang(language, 2)
            return
    
        
        self.root.ids["TAB_orders"].title = heading
        # Avvio degli ordini in corso
        col = self.db["orders"]
        self.root.ids.LS_products.clear_widgets()
        for line in col.find().sort("_id", pymongo.DESCENDING):
            status_st = ""
            color_st = ""
            if line["status"] == "now_ordered":
                status_st = lang(language, 4)
                color_st = "#BA681F"
                obj_instance = ObjectId(line["_id"])
                col.update_one({"_id": obj_instance}, {"$set": {"status": "ordered"}})
            if line["status"] == "ordered":
                status_st = lang(language, 4)
                color_st = "#BA681F"
            if line["status"] == "in_progress":
                status_st = lang(language, 5)
                color_st = "#08506D"
            if line["status"] == "done":
                status_st = lang(language, 7)
                color_st = "#158B09"
            customer_and_table = f"{lang(language, 9)}: {line['customer']} - {lang(language, 10)}: {line['table']}"
            date_and_time = f"{lang(language, 11)}: {line['date'][6:]}/{line['date'][4:6]}/{line['date'][:4]} - {lang(language, 12)}: {line['time'][:2]}:{line['time'][2:4]}:{line['time'][4:]}"
            list_item = ThreeLineListItem(text=f"{list_title_size}ID: {line['_id']}[/size]",
                                            secondary_text=f"{normal_size}{customer_and_table} - {date_and_time}[/size]",
                                            tertiary_text=f"{normal_size}{lang(language, 6)}: {status_st}[/size]",
                                            theme_text_color="Custom", text_color="#B0B006",
                                            secondary_theme_text_color="Custom", secondary_text_color="#B0B006",
                                            tertiary_theme_text_color="Custom", tertiary_text_color=color_st)
            list_item.bind(on_release = lambda x, item=list_item: self.show_detailed_order(item))
            self.root.ids.LS_products.add_widget(list_item)
            
        Clock.schedule_once(self.scheduled_permissions_check, 3.0) # Controllo permessi
    
    def scheduled_permissions_check(self, dt): # Controllo permessi
        if platform == "android":
            def check_permissions(perms):
                for perm in perms:
                    if check_permission(perm) != True:
                        return False
                return True
            
            perms = [Permission.WRITE_EXTERNAL_STORAGE,
                    Permission.READ_EXTERNAL_STORAGE,
                    Permission.READ_MEDIA_AUDIO,
                    Permission.INTERNET]    
            if check_permissions(perms)!= True:
                request_permissions(perms)
            Clock.schedule_once(self.scheduled_audio_load, 15.0)
        else: pass
    
    def scheduled_audio_load(self, dt): # Caricamento suono di notifica
        if platform == "android":
            dir = primary_external_storage_path()
            download_dir_path = os.path.join(dir, "Download")
            self.sound = SoundLoader.load(f"{download_dir_path}/beep.wav")
    
    def order_undo(self):
        self.root.current = "orders_table"
    
    def order_change_state(self):
        col = self.db["orders"]
        obj_instance = ObjectId(self.instance.text.replace(f"{list_title_size}","").replace("ID: ","").replace("[/size]",""))
        order = col.find_one({"_id": obj_instance})
        if order["status"] == "ordered":
            col.update_one({"_id": obj_instance}, {"$set": {"status": "in_progress"}})
            self.instance.tertiary_text = f"{normal_size}{lang(language, 6)}: {lang(language, 5)}[/size]"
            self.instance.tertiary_text_color = "#08506D"
        if order["status"] == "in_progress":
            col.update_one({"_id": obj_instance}, {"$set": {"status": "done"}})
            self.instance.tertiary_text = f"{normal_size}{lang(language, 6)}: {lang(language, 7)}[/size]"
            self.instance.tertiary_text_color = "#158B09"
        if order["status"] == "done":
            self.remove_widget_instance(self.instance, self.list_item)
        self.root.current = "orders_table"
    
    def show_detailed_order(self, instance): # Visualizzazione dettagliata della comanda
        col = self.db["orders"]
        obj_instance = ObjectId(instance.text.replace(f"{list_title_size}","").replace("ID: ","").replace("[/size]",""))
        order = col.find_one({"_id": obj_instance})
        order_string = f"""[color=1F7F06]{title_size}ID: {order['_id']}[/size]
{list_title_size}{lang(language, 9)}: {order["customer"]} - {lang(language, 10)}: {order["table"]}
{lang(language, 11)}: {order['date'][6:]}/{order['date'][4:6]}/{order['date'][:4]} - {lang(language, 12)}: {order['time'][:2]}:{order['time'][2:4]}:{order['time'][4:]}[/color]
[color=DE7A10][b]-----------------------------------[/b][/color]
[color=B0B006]{order["order"]}[/color]
[color=DE7A10][b]-----------------------------------[/b][/color]
[color=B29933]{order["order_note"]}[/size][/color]"""
        self.root.ids["L_order"].text = order_string
        self.instance = instance
        self.list_item = self.root.ids.LS_products
        if order["status"] == "ordered" or order["status"] == "seen": self.root.ids["B_order_change_state"].text = lang(language, 1) + lang(language, 5)
        if order["status"] == "in_progress": self.root.ids["B_order_change_state"].text = lang(language, 1) + lang(language, 7)
        if order["status"] == "done": self.root.ids["B_order_change_state"].text = lang(language, 8)
        self.root.current = "orders_detail"
            
    def remove_widget_instance(self, instance, parent_widget): # Eliminazione comanda alla pressione del tasto
        # Rimozione dal database
        col = self.db["orders"]
        obj_instance = ObjectId(instance.text.replace(f"{list_title_size}","").replace("ID: ","").replace("[/size]",""))
        col.delete_one({"_id": obj_instance})
        # Rimozione dall'interfaccia
        parent_widget.remove_widget(instance)
    
    def build(self):
        # Impostazioni stile
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Yellow"
        # Avvio delle funzioni del database
        Clock.schedule_once(self.scheduled_database_connection_1, 0.2) # Avvio della connessione al database
        Clock.schedule_interval(self.scheduled_function, update_time) # Controllo del database periodico
        # Caricamento suono di notifica su pc
        if platform != "android":
            self.sound = SoundLoader.load("beep.wav")
        else: self.sound = "-"
        
        self.KV = f"""
ScreenManager:
    Screen:
        name: "orders_table"
        BoxLayout:
            orientation: "vertical"
            MDTopAppBar:
                id: TAB_orders
                anchor_title: "left"
            ScrollView:
                MDList:
                    id: LS_products
    Screen:
        name: "database_connection"
        BoxLayout:
            orientation: "vertical"
            ScrollView:
                MDLabel:
                    id: L_connection
                    text: "{lang(language, 3)}"
                    font_style: "H3"
                    halign: "center"
                    theme_text_color: "Custom"
                    text_color: "#B0B006"
    Screen:
        name: "orders_detail"
        BoxLayout:
            spacing: 4
            orientation: "vertical"
            ScrollView:
                MDLabel:
                    id: L_order
                    font_style: "H6"
                    halign: "left"
                    markup: True
                    size_hint: 1, None
                    size: self.texture_size
            MDFillRoundFlatButton:
                id: B_order_undo
                text: "{lang(language, 0)}"
                size_hint_x: 1
                on_release: app.order_undo()
            MDFillRoundFlatButton:
                id: B_order_change_state
                size_hint_x: 1
                on_release: app.order_change_state()
                    """
        
        return Builder.load_string(self.KV)

if __name__ == "__main__":
    MainWindow().run()