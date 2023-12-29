from kivymd.app import MDApp
from kivy.lang import Builder
from kivy.clock import Clock
from kivy.uix.screenmanager import ScreenManager, Screen
from kivymd.uix.menu import MDDropdownMenu
from kivymd.uix.list import OneLineListItem, ThreeLineListItem
from kivymd.uix.button import MDFlatButton
from kivymd.uix.dialog import MDDialog
from kivy.properties import StringProperty
import pymongo,os
from bson.objectid import ObjectId

# Versione 1.0.0 r1

# Lingue

def lang(lg:str, index:int):
    it = ["0@-@Comande","1@-@Seleziona la lingua","2@-@Connessione al database","3@-@Inserisci il link","4@-@SALVA","5@-@Connessione al database fallita!",
          "6@-@Connessione al database in corso....","7@-@Intestazione","8@-@Scegli un intestazione","9@-@Annulla","10@-@Elimina","11@-@OPZIONI"]
    en = ["0@-@Orders","1@-@Select language","2@-@Database connection","3@-@Enter the link","4@-@SAVE","5@-@Database connection failed",
          "6@-@Connecting to database....","7@-@Heading","8@-@Select an heading","9@-@Undo","10@-@Delete","11@-@SETTINGS"]
    if lg == "ITALIANO": return it[index][it[index].index("@-@")+3:]
    if lg == "ENGLISH": return en[index][en[index].index("@-@")+3:]

class IconListItem(OneLineListItem):
    icon = StringProperty()

class MainWindow(MDApp):
    
    def open_lang_menu(self): # Menu selezione lingua
        menu_items = [
            {
                "viewclass": "OneLineListItem",
                "text": "ITALIANO",
                "theme_text_color": "Custom",
                "text_color": "#B0B006",
                "on_release": lambda x="ITALIANO": self.lang_menu_callback(x),
            },{
                "viewclass": "OneLineListItem",
                "theme_text_color": "Custom",
                "text_color": "#B0B006",
                "text": "ENGLISH",
                "on_release": lambda x="ENGLISH": self.lang_menu_callback(x),
            }
        ]
        self.lm = MDDropdownMenu(caller=self.root.ids.B_language_select, items=menu_items, width_mult=2,max_height=100)
        self.lm.open()
    
    def open_options_menu(self, op): # Menu opzioni (3 punti in alto a dx)
        menu_items = [
            {
                "viewclass": "OneLineListItem",
                "text": lang(self.language, 11),
                "theme_text_color": "Custom",
                "text_color": "#B0B006",
                "on_release": lambda x="-": self.open_options(x),
            }
        ]
        self.om = MDDropdownMenu(caller=self.root.ids.TAB_orders.ids["right_actions"], items=menu_items, width_mult=2,max_height=100)
        self.om.open()
    
    def open_options(self, op): # Apertura scheda opzioni
        self.root.current = "options_menu"
        self.om.dismiss()
    
    def lang_menu_callback(self, text_item): # Callback quando si cambia la lingua
        self.root.ids["B_language_select"].text = text_item
        self.language = text_item
        self.root.ids["L_title"].text = lang(self.language, 0)
        self.title = lang(self.language, 0)
        self.root.ids["L_language_select"].text = lang(self.language, 1)
        self.root.ids["L_database_connection"].text = lang(self.language, 2)
        self.root.ids["TF_database_connection"].hint_text = lang(self.language, 3)
        self.root.ids["B_save"].text = lang(self.language, 4)
        self.root.ids["L_heading"].text = lang(self.language, 7)
        self.root.ids["TF_heading"].hint_text = lang(self.language, 8)
        self.lm.dismiss()
    
    def show_alert(self, instance): # Visualizzazione dettagliata della comanda
        col = self.db["orders"]
        obj_instance = ObjectId(instance.text[4:])
        order = col.find_one({"_id": obj_instance})
        order_string = f"[color=#B0B006]ID: {order['_id']}\n{order['customer_and_table']}\n{order['date_time']}\n\n{order['order']}\n\n{order['order_note']}[/color]"
        self.dialog = MDDialog(
            text=order_string,
            buttons=[
                MDFlatButton(
                    text=lang(self.language, 9),
                    theme_text_color="Custom",
                    text_color=self.theme_cls.primary_color,
                    on_release = lambda _: self.dialog.dismiss()
                ),
                MDFlatButton(
                    text=lang(self.language, 10),
                    theme_text_color="Custom",
                    text_color=self.theme_cls.primary_color,
                    on_release = lambda x: self.remove_widget_instance(instance, self.root.ids.LS_products),
                ),
            ],
        )
        self.dialog.open()
            
    def remove_widget_instance(self, instance, parent_widget): # Eliminazione comanda alla pressione del tasto
        # Rimozione dal database
        col = self.db["orders"]
        obj_instance = ObjectId(instance.text[4:])
        col.delete_one({"_id": obj_instance})
        # Rimozione dall'interfaccia
        parent_widget.remove_widget(instance)
        self.dialog.dismiss()

    def save_options(self): # Salvataggio opzioni alla pressione del pulsante
        self.mongodb_connection = self.root.ids["TF_database_connection"].text
        self.heading = self.root.ids["TF_heading"].text
        options_file = open(f"{os.environ['HOME']}/Orders/options2.txt", "w")
        options_file.write(f"db_connection={self.mongodb_connection}\nheading={self.heading}\nlanguage={self.language}")
        options_file.close()
        self.root.ids["L_connection"].text = lang(self.language, 6)
        self.root.ids["TAB_orders"].title = self.heading
        Clock.schedule_once(self.scheduled_database_connection, 0.2)
    
    def scheduled_database_connection(self, dt): # Connessione al database e caricamento variabili
        try:
            self.dbclient = pymongo.MongoClient(self.mongodb_connection)
            self.dbclient.server_info() # Test per la connessione al database
            self.title = self.heading
            self.db = self.dbclient["Bar"]
            self.root.current = "orders_table"
            self.root.ids["L_connection"].text = ""
        except:
            self.root.ids["L_connection"].text = lang(self.language, 5)
        # Avvio degli ordini in corso
        col = self.db["orders"]
        self.root.ids.LS_products.clear_widgets()
        for line in col.find():
            list_item = ThreeLineListItem(text=f"ID: {line['_id']}",secondary_text=f"{line['customer_and_table']}",tertiary_text=f"{line['date_time']}",
                                          theme_text_color="Custom", text_color="#B0B006",
                                          secondary_theme_text_color="Custom", secondary_text_color="#B0B006",
                                          tertiary_theme_text_color="Custom", tertiary_text_color="#B0B006")
            list_item.bind(on_release = lambda x, item=list_item: self.show_alert(item))
            self.root.ids.LS_products.add_widget(list_item)
        col.update_many({"status": "0"}, {"$set": {"status": "1"}})
    
    def scheduled_start(self, dt): # Impostazioni ritardate di un secondo per permettere la creazione dei widgets
        self.root.ids["L_connection"].text = lang(self.language, 6)
        self.root.ids["TAB_orders"].title = self.heading
        Clock.schedule_once(self.scheduled_database_connection, 0.2)
    
    def scheduled_function(self, dt): # Funzione che controlla il database periodicamente
        if self.root.current == "orders_table":
            col = self.db["orders"]
            for line in col.find({"status": "0"}):
                list_item = ThreeLineListItem(text=f"ID: {line['_id']}",secondary_text=f"{line['customer_and_table']}",tertiary_text=f"{line['date_time']}",
                                              theme_text_color="Custom", text_color="#B0B006",
                                              secondary_theme_text_color="Custom", secondary_text_color="#B0B006",
                                              tertiary_theme_text_color="Custom", tertiary_text_color="#B0B006")
                list_item.bind(on_release = lambda x, item=list_item: self.show_alert(item))
                self.root.ids.LS_products.add_widget(list_item)
                obj_instance = ObjectId(line["_id"])
                col.update_one({"_id": obj_instance}, {"$set": {"status": "1"}})
    
    def build(self):
        self.language = "ENGLISH"
        self.heading = lang(self.language, 0)
        self.theme_cls.theme_style = "Dark"
        self.theme_cls.primary_palette = "Yellow"
        self.mongodb_connection = "mongodb://localhost:27017/"
        
        # Avvio applicazione
        
        if os.path.exists(f"{os.environ['HOME']}/Orders") == False: os.mkdir(f"{os.environ['HOME']}/Orders")

        if os.path.exists(f"{os.environ['HOME']}/Orders/options2.txt") == True:
            options_file = open(f"{os.environ['HOME']}/Orders/options2.txt", "r")
            self.mongodb_connection = options_file.readline().replace("db_connection=", "").replace("\n", "")
            self.heading = options_file.readline().replace("heading=", "").replace("\n", "")
            self.language = options_file.readline().replace("language=", "").replace("\n", "")
            options_file.close()
            Clock.schedule_once(self.scheduled_start, 1.0)
        
        Clock.schedule_interval(self.scheduled_function, 5.0)
        
        self.KV = f"""
ScreenManager:
    Screen:
        name: "options_menu"
        BoxLayout:
            orientation: "vertical"
            MDLabel:
                id: L_title
                text: "{lang(self.language, 0)}"
                font_style: "H4"
                halign: "center"
                theme_text_color: "Custom"
                text_color: "#B0B006"
            MDLabel:
                id: L_language_select
                text: "{lang(self.language, 1)}"
                font_style: "H6"
                halign: "center"
                theme_text_color: "Custom"
                text_color: "#B0B006"
            MDRaisedButton:
                id: B_language_select
                text: "{self.language}"
                on_release: app.open_lang_menu() 
                size_hint_x: 1
            MDLabel:
                id: L_heading
                text: "{lang(self.language, 7)}"
                font_style: "H4"
                halign: "center"
                theme_text_color: "Custom"
                text_color: "#B0B006"
            MDTextField:
                id: TF_heading
                text: "{self.heading}"
                hint_text: "{lang(self.language, 8)}"
            MDLabel:
                id: L_database_connection
                text: "{lang(self.language, 2)}"
                font_style: "H6"
                halign: "center"
                theme_text_color: "Custom"
                text_color: "#B0B006"
            MDTextField:
                id: TF_database_connection
                hint_text: "{lang(self.language, 3)}"
                text: "{self.mongodb_connection}"
            MDRaisedButton:
                id: B_save
                text: "{lang(self.language, 4)}"
                size_hint_x: 1
                on_release: app.save_options()
            MDLabel:
                id: L_connection
                text: ""
                font_style: "H6"
                halign: "center"
                theme_text_color: "Custom"
                text_color: "#B0B006"
    Screen:
        name: "orders_table"
        BoxLayout:
            orientation: "vertical"
            MDTopAppBar:
                id: TAB_orders
                anchor_title: "left"
                right_action_items: [["dots-vertical", lambda x: app.open_options_menu(x)]]
                pos_hint: {{"center_x": .5, "center_y": .5}}
            ScrollView:
                MDList:
                    id: LS_products
                    """
        
        return Builder.load_string(self.KV)

if __name__ == "__main__":
    MainWindow().run()