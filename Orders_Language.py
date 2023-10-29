# MainWindow

ita_mw = ["Opzioni","Creazione Categorie","Descrizione","Crea >>","Inserimento Prodotti","Prezzo","Inserisci >>","Impostazioni Festa","Intestazione festa",
          "Selezione Cliente e Numero Tavolo","Nome cliente","Note aggiuntive che appariranno alla fine dello scontrino","Stampa o Salva Scontrino","Stampa",
          "Salva","Apri il Database","Quantità","Totale","€","Errore","Il campo descrizione non può essere vuoto","Attenzione",
          "La categoria inserita esiste già\nVuoi eliminarla?","Categoria non selezionata!","Compilare i campi descrizione e prezzo!","Prezzo inserito non corretto!",
          "Elimina","Sposta su","Sposta giù","Aggiungi quantità specifica","Aggiungi +1 (doppio click sx)","Togli -1 (click dx)","Crea un menù",
          "Il prodotto selezionato si trova già nella prima casella!","Il prodotto selezionato si trova già nell'ultima casella","Quantità specifica",
          "Inserisci un numero intero che verrà aggiunto alla quantità nella tabella dello scontrino","Inserisci","Annulla","Devi inserire un numero intero!",
          "Cliente","Numero tavolo","qta","EUR"]
eng_mw = ["Options","Category Creation","Description","Create >>","Products Insert","Price","Insert >>","Party Settings","Party name",
          "Customer and Table Number Selection","Customer name","Additional notes that will appear at the end of the receipt","Print or Save Receipt","Print",
          "Save","Open Database","Quantity","Total","£","Error","Description field can't be empty","Caution",
          "Category selected already exists\nDo you want to delete it?","Category not selected!","Fill in the description and price fields!","Incorrect price entered",
          "Delete","Move up","Move down","Add specific quantity","Add +1 (double left click)","Remove -1 (right click)","Create a menu",
          "Selected product is already in the first slot","Selected product is already in the last slot","Specific quantity",
          "Insert an integer number that will be added to the receipt table","Insert","Cancel","You must insert an integer number!",
          "Customer","Table number","qty","GBP"]

# DatabaseWindow

ita_dw = ["Selezione orari\nNella casella qui sotto\npuoi selezionare un\nrange di orari.\nEsempio 17:30-20:00","Orari","Interroga","Dettaglio vendite","Data scontrino",
          "Orario scontrino","Nessun dato per la data selezionata!","Formato ora inserito non corretto!","Nessun dato per la data e ora selezionata!",
          "Totale vendite del","Totale complessivo","Nessun dato per le date selezionate!","Nessun dato per le date e ora selezionate!","Totale vendite dal","al",
          "Stai per eliminare i dati selezionati dal database\nL'operazione non è annullabile\nVuoi continuare?","Eliminazione effettuata"]
eng_dw = ["Times selection\nIn the bottom field\nyou can select a\nrange of times\nExample 17:30-20:00","Times","Query","Sales detail","Receipt date",
          "Receipt time","No data for the selected date!","Inserted time format not correct!","No data for the selected date and time!",
          "Total sales of","Total","No data for the selected dates!","No data for the selected dates and time!","Total sales from","to",
          "You're going to delete the selected data from the database\nThe operation is not reversible\nDo you want to continue?","Deletion done"]

# CreateMenuWindow

ita_cmw = ["Creazione menù per il prodotto","Prezzo unitario","Il menù non può essere vuoto","Il menù che stai tentando di eliminare non esiste"]
eng_cmw = ["Menu creation for product","Unit price","Menu can't be empty","The menu you're trying to delete doesn't exists"]

# OptionsMenuWindow

ita_omw = ["Menù opzioni","Selezione lingua","Connessione al database","Il programma usa MongoDB come database\nInserisci il link nella casella qui sotto\nSe hai un database locale il link sarà: mongodb://localhost:27017/",
           "Link al database","Intestazione","Inserisci un intestazione, verrà usata sia sulla testa\ndel programma che ad ogni inizio scontrino","Interfaccia grafica",
           "Seleziona uno stile grafico per il programma","Logo",
           "Seleziona un immagine PNG per il logo\nVerrà posizionato in alto a sinistra nell'interfaccia\nLe dimensioni ideali sono 190x85 pixel\nAttualmente stai usando il file",
           "Seleziona","Icona",
           "Seleziona un immagine PNG per l'icona\nL'icona la troverai su ogni finestra\nLe dimensioni ideali sono 51x21 pixel\nAttualmente stai usando il file",
           "Chiudi e Salva","Immagini","La casella per la connessione al database non può essere vuota","Connessione al database in corso...","Connessione al database fallita!"]
eng_omw = ["Options menu","Language selection","Database connection","The program use MongoDB as database\nInsert database link in the bottom field\nIf you have a local database the link will be: mongodb://localhost:27017/",
           "Database link","Heading","Insert an heading, will be used both\non program head and at every receipt beginning","Graphical interface",
           "Select a graphic style for the program","Logo",
           "Select a PNG image for the logo\nIt will be positioned on top left of the interface\nIdeal dimensions are 190x85 pixels\nActually you're using the file",
           "Select","Icon",
           "Select a PNG image for the icon\nThe icon you will find in every window\nIdeal dimensions are 51x21 pixels\nActually you're using the file",
           "Close and Save","Images","The database connection field can't be empty","Connecting to database...","Connection to database failed!"]
    
def msg(lang:str, msg_index:int, msg_window:str):
    if msg_window == "MainWindow":
        if lang == "ITALIANO": return ita_mw[msg_index]
        if lang == "ENGLISH": return eng_mw[msg_index]
    if msg_window == "DatabaseWindow":
        if lang == "ITALIANO": return ita_dw[msg_index]
        if lang == "ENGLISH": return eng_dw[msg_index]
    if msg_window == "CreateMenuWindow":
        if lang == "ITALIANO": return ita_cmw[msg_index]
        if lang == "ENGLISH": return eng_cmw[msg_index]
    if msg_window == "OptionsMenuWindow":
        if lang == "ITALIANO": return ita_omw[msg_index]
        if lang == "ENGLISH": return eng_omw[msg_index]