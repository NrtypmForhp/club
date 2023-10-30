# MainWindow

ita_mw = ["Ricerca Automatica","Apri Database","Opzioni","Codice fiscale","Nome","Cognome","Data di nascita","Luogo di nascita","MASCHIO","FEMMINA","Comune di residenza",
          "Indirizzo di residenza","CAP - Codice di avviamento postale","e-mail","Numero tessera","Inserisci","Cerca","Esporta Excel","Gestione tesseramenti",
          "** Funzioni rapide **\nF5: Pulisce tutti i campi","Errore","I dati che devono obbligatoriamente essere inseriti sono:\n- Codice fiscale\n- Nome\n- Cognome",
          "Codice fiscale non corretto!","Nome non corretto!","Cognome non corretto!","Attenzione","Persona già presente nel database\nVuoi modificare i dati con quelli inseriti?",
          "Inserimento effettuato","Sesso","Data di tesseramento","Modifica effettuata","Persone trovate","Il pulsante non può essere usato con la ricerca automatica attiva!",
          "I campi non contengono nulla!\nControlla di aver compilato almeno uno dei seguenti campi:\nCodice fiscale - Completo\nNome - almeno 3 lettere\nCognome - almeno 3 lettere",
          "Nessun dato trovato nel database!"]
eng_mw = ["Automatic Search","Open Database","Options","Tax id code","Name","Surname","Date of birth","Birth place","MALE","FEMALE","City of residence",
          "Residential address","Postal code","e-mail","Card number","Insert","Search","Excel Export","Membership Management",
          "** Quick functions **\nF5: Clear all the fields","Error","Data that must be inserted is:\n-Tax id code\n- Name\n- Surname",
          "Incorrect tax code!","Incorrect name!","Incorrect surname","Caution","Person already present in the database\nDo you want to change the data with those entered?",
          "Insertion completed","Sex","Date of membership","Change made","People founded","The button can't be used with the automatic search enabled!",
          "Fields contain nothing!\nCheck that you have filled in the following fields:\nTax id code - Complete\nName - at least 3 charachters\nSurname - at least 3 charachters",
          "No data founded in the database!"]

# DatabaseWindow

ita_dw = ["Non tesserati","Elimina","Uguale a","Contiene","Maggiore di","Minore di","Codice fiscale completo","Parte di codice fiscale","Nome completo","Parte del nome",
          "Cognome completo","Parte del cognome","Esempio: 15/06/1982","Luogo di nascita completo","Parte del luogo di nascita","Casella non necessaria",
          "Città di residenza completa","Parte della città di residenza","Indirizzo di residenza completo","Parte dell'indirizzo di residenza","CAP completo",
          "Parte del CAP","e-mail completa","Parte della e-mail","Numero tessera specifico","Maggiore del numero inserito","Minore del numero inserito","Data non corretta!",
          "CAP errato!","Numero tessera non corretto!","Stai per eliminare i dati selezionati dal database!\nL'operazione non è reversibile!\nSei sicuro?"]
eng_dw = ["Not members","Delete","Equal to","Contains","Greater than","Less than","Complete tax id code","Part of tax code","Complete name","Part of name",
          "Complete surname","Part of surname","Example: 15/06/1982","Complete birth place","Part of birth place","Field not necessary",
          "Complete city of residence","Part of city of residence","Complete residential address","Part of residential address","Complete postal code",
          "Part of postal code","Complete e-mail","Part of e-mail","Specific card number","Greater than the number entered","Less than the number entered","Incorrect date!",
          "Incorrect postal code!","Incorrect card number!","You're going to delete the selected data from database!\nThe operation is not reversible!\nAre you sure?"]

# ExcelWindow

ita_ew = ["Esportazione Excel","Colonne","Formato ASI","Imposta","Aggiungi riga","Rimuovi riga","Esporta in excel","Non è presente nessuna colonna!",
          "Seleziona una riga (click sopra)!","Personalizza","Inserimento personalizzato",
          "L'inserimento personalizzato ti permette di\ninserire un valore extra alla cella.\nQuesto valore in fase di esportazione verrà solo copiato!","Annulla",
          "La dicitura DB: non può essere inserita manualmente!","Questa riga inizia la ricerca nel database\nDeve essere necessariamente inserita nell'ultima riga!",
          "STAGIONE","DISCIPLINA","QUALIFICA","TIPO TESSERA","DB:Nome","DB:Cognome","DB:Codice fiscale","DB:Città di residenza","DB:Indirizzo di residenza","DB:CAP - Codice di avviamento postale",
          "DB:e-mail","DB:Numero tessera","CODICE AFFILIAZIONE","La tabella non può essere vuota!","Nella tabella ci deve essere almeno un valore da controllare nel database!",
          "DB:Data di nascita","DB:Luogo di nascita","DB:Sesso","DB:Data di tesseramento","File salvato","Percorso"]
eng_ew = ["Excel Export","Columns","ASI Format","Set","Add row","Remove row","Export to excel","No columns present!",
          "Select a row (click on)!","Customize","Custom insertion",
          "Custom insertion allows you to\nenter an extra value to the cell.\nThis value during exportation will be only copied!","Cancel",
          "The wording DB: cannot be entered manually","This row starts the database searching\nIt must be entered in the last row!",
          "SEASON","DISCIPLINE","QUALIFICATION","CARD TYPE","DB:Name","DB:Surname","DB:Tax id code","DB:City of residence","DB:Residential address","DB:Postal code",
          "DB:e-mail","DB:Card number","AFFILIATION CODE","The table cannot be empty!","In the table there must be at least one value to control in the database!",
          "DB:Date of birth","DB:Birth place","DB:Sex","DB:Date of membership","File saved","Path"]

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
    if msg_window == "ExcelWindow":
        if lang == "ITALIANO": return ita_ew[msg_index]
        if lang == "ENGLISH": return eng_ew[msg_index]
    if msg_window == "OptionsMenuWindow":
        if lang == "ITALIANO": return ita_omw[msg_index]
        if lang == "ENGLISH": return eng_omw[msg_index]