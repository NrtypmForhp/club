# File di debug, serve a cercare l'index nelle liste dei file lingua

import Memberships_Languages as ML
import Orders_Language as OL

st = "DB:data di tesseramento"
ty = "Memberships"

if ty == "Orders":
    if st in OL.ita_mw: print("MainWindow - " + str(OL.ita_mw.index(st)))
    if st in OL.eng_mw: print("MainWindow - " + str(OL.eng_mw.index(st)))
    if st in OL.ita_dw: print("DatabaseWindow - " + str(OL.ita_dw.index(st)))
    if st in OL.eng_dw: print("DatabaseWindow - " + str(OL.eng_dw.index(st)))
    if st in OL.ita_cmw: print("CreateMenuWindow - " + str(OL.ita_cmw.index(st)))
    if st in OL.eng_cmw: print("CreateMenuWindow - " + str(OL.eng_cmw.index(st)))
    if st in OL.ita_omw: print("OptionsMenuWindow - " + str(OL.ita_omw.index(st)))
    if st in OL.eng_omw: print("OptionsMenuWindow - " + str(OL.eng_omw.index(st)))

if ty == "Memberships":
    if st in ML.ita_mw: print("MainWindow - " + str(ML.ita_mw.index(st)))
    if st in ML.eng_mw: print("MainWindow - " + str(ML.eng_mw.index(st)))
    if st in ML.ita_dw: print("DatabaseWindow - " + str(ML.ita_dw.index(st)))
    if st in ML.eng_dw: print("DatabaseWindow - " + str(ML.eng_dw.index(st)))
    if st in ML.ita_ew: print("ExcelWindow - " + str(ML.ita_ew.index(st)))
    if st in ML.eng_ew: print("ExcelWindow - " + str(ML.eng_ew.index(st)))