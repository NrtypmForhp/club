from kivymd.app import MDApp
from kivy.lang import Builder
from kivy.clock import Clock
from kivymd.uix.list import ThreeLineListItem
from kivymd.uix.button import MDFlatButton
from kivymd.uix.dialog import MDDialog
from kivy.properties import StringProperty
import pymongo
from bson.objectid import ObjectId

# Versione 1.0.0 r1

# Variabili Globali

mongodb_connection = "mongodb://localhost:27017/" # Stringa di connessione al database
language = "IT" # Stringa di selezione lingua
heading = "98 Ottani The Club" # Stringa di selezione titolo

# Lingue

def lang(lg:str, index:int):
    it = ["0@-@INDIETRO","1@-@ELIMINA ORDINE","2@-@Connessione al database fallita!","3@-@Connessione al database in corso...."]
    en = ["0@-@BACK","1@-@DELETE ORDER","2@-@Connection to database failed!","3@-@Connection to database...."]
    if lg == "IT": return it[index][it[index].index("@-@")+3:]
    if lg == "EN": return en[index][en[index].index("@-@")+3:]

# Applicazione

class MainWindow(MDApp):
    def scheduled_function(self, dt): # Funzione che controlla il database periodicamente
        if self.root.current == "orders_table":
            col = self.db["orders"]
            for line in col.find({"status": "0"}):
                list_item = ThreeLineListItem(text=f"ID: {line['_id']}",secondary_text=f"{line['customer_and_table']}",tertiary_text=f"{line['date_time']}",
                                              theme_text_color="Custom", text_color="#B0B006",
                                              secondary_theme_text_color="Custom", secondary_text_color="#B0B006",
                                              tertiary_theme_text_color="Custom", tertiary_text_color="#B0B006")
                list_item.bind(on_release = lambda x, item=list_item: self.show_detailed_order(item))
                self.root.ids.LS_products.add_widget(list_item)
                obj_instance = ObjectId(line["_id"])
                col.update_one({"_id": obj_instance}, {"$set": {"status": "1"}})
    
    def scheduled_database_connection_1(self, dt): # Impostazione schermata iniziale
        self.root.current = "database_connection"
        Clock.schedule_once(self.scheduled_database_connection_2, 1.0) # Avvio della connessione al database
    
    def scheduled_database_connection_2(self, dt): # Connessione al database e caricamento variabili
        try:
            self.dbclient = pymongo.MongoClient(mongodb_connection)
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
        for line in col.find():
            list_item = ThreeLineListItem(text=f"ID: {line['_id']}",secondary_text=f"{line['customer_and_table']}",tertiary_text=f"{line['date_time']}",
                                          theme_text_color="Custom", text_color="#B0B006",
                                          secondary_theme_text_color="Custom", secondary_text_color="#B0B006",
                                          tertiary_theme_text_color="Custom", tertiary_text_color="#B0B006")
            list_item.bind(on_release = lambda x, item=list_item: self.show_detailed_order(item))
            self.root.ids.LS_products.add_widget(list_item)
        col.update_many({"status": "0"}, {"$set": {"status": "1"}})
    
    def order_undo(self):
        self.root.current = "orders_table"
    
    def order_delete(self):
        self.remove_widget_instance(self.instance, self.list_item)
        self.root.current = "orders_table"
    
    def show_detailed_order(self, instance): # Visualizzazione dettagliata della comanda
        col = self.db["orders"]
        obj_instance = ObjectId(instance.text[4:])
        order = col.find_one({"_id": obj_instance})
        order_string = f"""ID: {order['_id']}
{order["customer_and_table"]}
{order["date_time"]}
-----------------------------------
{order["order"]}
-----------------------------------
{order["order_note"]}"""
        self.root.ids["L_order"].text = order_string
        self.instance = instance
        self.list_item = self.root.ids.LS_products
        self.root.current = "orders_detail"
            
    def remove_widget_instance(self, instance, parent_widget): # Eliminazione comanda alla pressione del tasto
        # Rimozione dal database
        col = self.db["orders"]
        obj_instance = ObjectId(instance.text[4:])
        col.delete_one({"_id": obj_instance})
        # Rimozione dall'interfaccia
        parent_widget.remove_widget(instance)
    
    def build(self):
        # Impostazioni stile
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Yellow"
        # Avvio delle funzioni del database
        Clock.schedule_once(self.scheduled_database_connection_1, 0.2) # Avvio della connessione al database
        Clock.schedule_interval(self.scheduled_function, 5.0) # Controllo del database periodico
        
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
            MDLabel:
                id: L_connection
                text: "{lang(language, 3)}"
                font_style: "H4"
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
                    theme_text_color: "Custom"
                    text_color: "#B0B006"
                    size_hint: 1, None
                    size: self.texture_size
            MDFillRoundFlatButton:
                id: B_order_undo
                text: "{lang(language, 0)}"
                size_hint_x: 1
                on_release: app.order_undo()
            MDFillRoundFlatButton:
                id: B_order_delete
                text: "{lang(language, 1)}"
                size_hint_x: 1
                on_release: app.order_delete()
                    """
        
        return Builder.load_string(self.KV)

if __name__ == "__main__":
    MainWindow().run()