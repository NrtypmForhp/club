# MainWindow

ita_mw = ["0@-@Opzioni","1@-@Creazione ed Eliminazione Categorie","2@-@Descrizione","3@-@Crea >>","4@-@Inserimento Prodotti","5@-@Prezzo","6@-@Inserisci >>","7@-@Impostazioni Festa",
          "8@-@Intestazione festa","9@-@Selezione Cliente e Numero Tavolo","10@-@Nome cliente","11@-@Note aggiuntive che appariranno alla fine dello scontrino",
          "12@-@Stampa o Salva Scontrino","13@-@Stampa","14@-@Salva","15@-@Apri il Database","16@-@Quantità","17@-@Totale","18@-@€","19@-@Errore",
          "20@-@Il campo descrizione non può essere vuoto","21@-@Attenzione","22@-@Categoria già esistente!","23@-@Categoria non selezionata!",
          "24@-@Compilare i campi descrizione e prezzo!","25@-@Prezzo inserito non corretto!","26@-@Elimina","27@-@Sposta su","28@-@Sposta giù",
          "29@-@Aggiungi quantità specifica","30@-@Aggiungi +1 (doppio click sx)","31@-@Togli -1 (click dx)","32@-@Crea un menù",
          "33@-@Il prodotto selezionato si trova già nella prima casella!","34@-@Il prodotto selezionato si trova già nell'ultima casella","35@-@Quantità specifica",
          "36@-@Inserisci un numero intero che verrà aggiunto alla quantità nella tabella dello scontrino","37@-@Inserisci","38@-@Annulla",
          "39@-@Devi inserire un numero intero!","40@-@Cliente","41@-@Numero tavolo","42@-@qta","43@-@EUR",
          "44@-@Invia automaticamente la categoria corrente","45@-@Nelle righe da inviare c'è qualcosa di sbagliato!\nCompilare come in questo esempio: 1-2-4 o 1-2:4-6",
          "46@-@Invio ordini","47@-@Controlla ordini","48@-@Connessione con il database fallita!\nControlla la connessione a internet!"]
eng_mw = ["0@-@Options","1@-@Creating and Deleting Categories","2@-@Description","3@-@Create >>","4@-@Products Insert","5@-@Price","6@-@Insert >>","7@-@Party Settings",
          "8@-@Party name","9@-@Customer and Table Number Selection","10@-@Customer name","11@-@Additional notes that will appear at the end of the receipt",
          "12@-@Print or Save Receipt","13@-@Print","14@-@Save","15@-@Open Database","16@-@Quantity","17@-@Total","18@-@£","19@-@Error",
          "20@-@Description field can't be empty","21@-@Caution","22@-@Already existing category","23@-@Category not selected!",
          "24@-@Fill in the description and price fields!","25@-@Incorrect price entered","26@-@Delete","27@-@Move up","28@-@Move down","29@-@Add specific quantity",
          "30@-@Add +1 (double left click)","31@-@Remove -1 (right click)","32@-@Create a menu","33@-@Selected product is already in the first slot",
          "34@-@Selected product is already in the last slot","35@-@Specific quantity","36@-@Insert an integer number that will be added to the receipt table",
          "37@-@Insert","38@-@Cancel","39@-@You must insert an integer number!","40@-@Customer","41@-@Table number","42@-@qty","43@-@GBP",
          "44@-@Automatically send the current category","45@-@There's something wrong in the lines to send!\nFill in as in this example: 1-2-4 or 1-2:4-6",
          "46@-@Sending orders","47@-@Check orders","48@-@Connection to database failed!\nCheck internet connection!"]

# DatabaseWindow

ita_dw = ["0@-@Selezione orari\nNella casella qui sotto puoi selezionare un\nrange di orari. Esempio 17:30-20:00","1@-@Orari","2@-@Interroga","3@-@Dettaglio vendite",
          "4@-@Data scontrino","5@-@Orario scontrino","6@-@Nessun dato per la data selezionata!","7@-@Formato ora inserito non corretto!",
          "8@-@Nessun dato per la data e ora selezionata!","9@-@Totale vendite del","10@-@Totale complessivo","11@-@Nessun dato per le date selezionate!",
          "12@-@Nessun dato per le date e ora selezionate!","13@-@Totale vendite dal","14@-@al",
          "15@-@Stai per eliminare i dati selezionati dal database\nL'operazione non è annullabile\nVuoi continuare?","16@-@Eliminazione effettuata",
          "17@-@Ordina risultati","18@-@Per prezzo","19@-@Per quantità","20@-@Per descrizione","21@-@Per prezzo (contrario)","22@-@Per quantità (contrario)",
          "23@-@Per descrizione (contrario)"]
eng_dw = ["0@-@Times selection\nIn the bottom field you can select a\nrange of times Example 17:30-20:00","1@-@Times","2@-@Query","3@-@Sales detail","4@-@Receipt date",
          "5@-@Receipt time","6@-@No data for the selected date!","7@-@Inserted time format not correct!","8@-@No data for the selected date and time!",
          "9@-@Total sales of","10@-@Total","11@-@No data for the selected dates!","12@-@No data for the selected dates and time!","13@-@Total sales from","14@-@to",
          "15@-@You're going to delete the selected data from the database\nThe operation is not reversible\nDo you want to continue?","16@-@Deletion done",
          "17@-@Sort results","18@-@By price","19@-@By quantity","20@-@By description","21@-@By price (reverse)","22@-@By quantity (reverse)",
          "23@-@By description (reverse)"]

# CreateMenuWindow

ita_cmw = ["0@-@Creazione menù per il prodotto","1@-@Prezzo unitario","2@-@Il menù non può essere vuoto","3@-@Il menù che stai tentando di eliminare non esiste"]
eng_cmw = ["0@-@Menu creation for product","1@-@Unit price","2@-@Menu can't be empty","3@-@The menu you're trying to delete doesn't exists"]

# OptionsMenuWindow

ita_omw = ["0@-@Menù opzioni","1@-@Selezione lingua","2@-@Connessione al database",
           "3@-@Il programma usa MongoDB come database\nInserisci il link nella casella qui sotto\nSe hai un database locale il link sarà: mongodb://localhost:27017/",
           "4@-@Link al database","5@-@Intestazione","6@-@Inserisci un intestazione, verrà usata sia sulla testa\ndel programma che ad ogni inizio scontrino",
           "7@-@Interfaccia grafica","8@-@Seleziona uno stile grafico per il programma","9@-@Logo",
           "10@-@Seleziona un immagine PNG per il logo\nVerrà posizionato in alto a sinistra nell'interfaccia\nLe dimensioni ideali sono 190x85 pixel\nAttualmente stai usando il file",
           "11@-@Seleziona","12@-@Icona",
           "13@-@Seleziona un immagine PNG per l'icona\nL'icona la troverai su ogni finestra\nLe dimensioni ideali sono 51x21 pixel\nAttualmente stai usando il file",
           "14@-@Chiudi e Salva","15@-@Immagini","16@-@La casella per la connessione al database non può essere vuota","17@-@Connessione al database in corso...",
           "18@-@Connessione al database fallita!"]
eng_omw = ["0@-@Options menu","1@-@Language selection","2@-@Database connection",
           "3@-@The program use MongoDB as database\nInsert database link in the bottom field\nIf you have a local database the link will be: mongodb://localhost:27017/",
           "4@-@Database link","5@-@Heading","6@-@Insert an heading, will be used both\non program head and at every receipt beginning","7@-@Graphical interface",
           "8@-@Select a graphic style for the program","9@-@Logo",
           "10@-@Select a PNG image for the logo\nIt will be positioned on top left of the interface\nIdeal dimensions are 190x85 pixels\nActually you're using the file",
           "11@-@Select","12@-@Icon","13@-@Select a PNG image for the icon\nThe icon you will find in every window\nIdeal dimensions are 51x21 pixels\nActually you're using the file",
           "14@-@Close and Save","15@-@Images","16@-@The database connection field can't be empty","17@-@Connecting to database...","18@-@Connection to database failed!"]
    
# OrdersWindow

ita_ow = ["0@-@Ordini","1@-@Data e Ora","2@-@Stato","3@-@ORDINATO","4@-@IN CORSO","5@-@ORDINE FATTO","6@-@Dettaglio ordine","7@-@Data ordine","8@-@Orario ordine",
          "9@-@Vedi ordine (doppo click sx)","10@-@Aggiorna","11@-@Cambia stato in: Ordinato","12@-@Cambia stato in: In Corso","13@-@Cambia stato in: Ordine Fatto"]
eng_ow = ["0@-@Orders","1@-@Date and Time","2@-@Status","3@-@ORDERED","4@-@IN PROGRESS","5@-@ORDER DONE","6@-@Order detail","7@-@Order date","8@-@Order time",
          "9@-@Show order (double left click)","10@-@Update","11@-@Change status in: Ordered","12@-@Change status in: In Progress","13@-@Change status in: Order Done"]

def msg(lang:str, msg_index:int, msg_window:str):
    if msg_window == "MainWindow":
        if lang == "ITALIANO": return ita_mw[msg_index][ita_mw[msg_index].index("@-@")+3:]
        if lang == "ENGLISH": return eng_mw[msg_index][eng_mw[msg_index].index("@-@")+3:]
    if msg_window == "DatabaseWindow":
        if lang == "ITALIANO": return ita_dw[msg_index][ita_dw[msg_index].index("@-@")+3:]
        if lang == "ENGLISH": return eng_dw[msg_index][eng_dw[msg_index].index("@-@")+3:]
    if msg_window == "CreateMenuWindow":
        if lang == "ITALIANO": return ita_cmw[msg_index][ita_cmw[msg_index].index("@-@")+3:]
        if lang == "ENGLISH": return eng_cmw[msg_index][eng_cmw[msg_index].index("@-@")+3:]
    if msg_window == "OptionsMenuWindow":
        if lang == "ITALIANO": return ita_omw[msg_index][ita_omw[msg_index].index("@-@")+3:]
        if lang == "ENGLISH": return eng_omw[msg_index][eng_omw[msg_index].index("@-@")+3:]
    if msg_window == "OrdersWindow":
        if lang == "ITALIANO": return ita_ow[msg_index][ita_ow[msg_index].index("@-@")+3:]
        if lang == "ENGLISH": return eng_ow[msg_index][eng_ow[msg_index].index("@-@")+3:]