def interface_style(name_of_style:str):
    style_string = ""
    if name_of_style == "": name_of_style = "98 Style" # Impostazione allo stile di default
    if name_of_style == "98 Style":
        style_string = """
                    QWidget{
                        background-color: #404040;
                    }
                    QLabel{
                        color: #A6AB13;
                        font: 10pt Arial;
                    }
                    QLabel[accessibleName="an_title"]{
                        color: #A6AB13;
                        font: bold 18pt Arial;
                    }
                    QLabel[accessibleName="an_section_title"]{
                        color: #A6AB13;
                        font: bold 14pt Arial;
                    }
                    QLineEdit{
                        border: 1px solid black;
                        background-color: #3D3A3D;
                        color: #A6AB13;
                        font: 10pt Arial;
                    }
                    QPushButton{
                        background-color: black;
                        color: #A6AB13;
                        font: 10pt Arial;
                        padding: 4px, 8px;
                        outline: 0px;
                        font-weight: 600;
                        border-radius: 15px;
                        border: 2px solid #A6AB13;
                        border-style: outset;
                        width: 110px;
                    }
                    QPushButton:hover{
                        background-color: #0C3B0B;
                    }
                    QComboBox{
                        background-color: black;
                        color: #A6AB13;
                        padding: 1px 0px 1px 3px; /* Rende possibile il cambio del colore del testo nell'hover */
                        font: 10pt Arial;
                    }
                    QComboBox:hover{
                        background-color: #0C3B0B;
                    }
                    QComboBox:selected{
                        background-color: #075206;
                    }
                    QTableWidget{
                        color: #A6AB13;
                        border: 2px solid black;
                        font: 10pt Arial;
                    }
                    QTableWidget::item:selected{
                        background-color: black;
                        color: #A6AB13;
                    }
                    QTableWidget QTableCornerButton::section {
                        background-color: black;
                        border: 1px solid black;
                        border-style: outset;
                    }
                    QHeaderView{
                        background-color: black;
                        color: #A6AB13;
                        font: 10pt Arial;
                    }
                    QTextEdit{
                        color: #A6AB13;
                        background-color: #3D3A3D;
                        border: 2px solid black;
                    }
                    QMenu{
                        border: 1px solid #A6AB13;
                    }
                    QMenu::item{
                        background-color: black;
                        color: #A6AB13;
                        font: 10pt Arial;
                    }
                    QMenu::item:selected{
                        background-color: #075206;
                    }
                    QMenu::item:disabled{
                        background-color: #3D3A3D;
                        color: white;
                    }
                    QSpinBox{
                        background-color: #3D3A3D;
                        color: #A6AB13;
                        padding-right: 15px;
                        border: 3px solid black;
                        font: 10pt Arial;
                    }
                    QSpinBox::up-button{
                        width: 16px;
                    }
                    QSpinBox::down-button{
                        width: 16px;
                    }
                    QCalendarWidget QAbstractItemView:enabled{
                        background-color: black;
                        color: yellow;
                        selection-background-color: yellow; 
                        selection-color: black;
                        font: 10pt Arial;
                    }
                    QCheckBox{
                        padding: 5px;
                        color: #A6AB13;
                        font: 10pt Arial;
                    }
                    QCheckBox::indicator{
                        border: 2px solid black;
                        width: 20px;
                        height: 20px;
                        border-radius: 12px;
                    }
                    QCheckBox::indicator:checked{
                        background-color: #A6AB13;
                    }
                    QCheckBox::indicator:unchecked{
                        background-color: none;
                    }
                    """
    if name_of_style == "Tech Style":
        style_string = """
                    QWidget{
                        background-color: black;
                    }
                    QLabel{
                        color: #1CB414;
                        font: 10pt SourceCodeVF;
                    }
                    QLabel[accessibleName="an_title"]{
                        color: #1CB414;
                        font: bold 18pt SourceCodeVF;
                    }
                    QLabel[accessibleName="an_section_title"]{
                        color: #1CB414;
                        font: bold 14pt SourceCodeVF;
                    }
                    QLineEdit{
                        border: 1px solid #0F7ABF;
                        background-color: #292424;
                        color: #1CB414;
                        font: 10pt SourceCodeVF;
                    }
                    QPushButton{
                        background-color: #292424;
                        color: #1CB414;
                        font: 10pt SourceCodeVF;
                    }
                    QPushButton:hover{
                        background-color: #0C3B0B;
                        color: #D59B0C;
                    }
                    QComboBox{
                        background-color: #292424;
                        color: #1CB414;
                        padding: 1px 0px 1px 3px; /* Rende possibile il cambio del colore del testo nell'hover */
                        font: 10pt SourceCodeVF;
                    }
                    QComboBox:hover{
                        background-color: #0C3B0B;
                        color: #D59B0C;
                    }
                    QComboBox:selected{
                        background-color: #075206;
                        color: #D59B0C;
                    }
                    QTableWidget{
                        color: #1CB414;
                        border: 1px solid #0F7ABF;
                        font: 10pt SourceCodeVF;
                    }
                    QTableWidget::item:selected{
                        background-color: #0C3B0B;
                        color: #D59B0C;
                    }
                    QTableWidget QTableCornerButton::section {
                        background-color: #292424;
                        border: 1px solid #0F370D;
                        border-style: outset;
                    }
                    QHeaderView{
                        background-color: #292424;
                        color: #1CB414;
                        font: 10pt SourceCodeVF;
                    }
                    QTextEdit{
                        color: #1CB414;
                        background-color: #292424;
                        border: 2px solid #0F7ABF;
                        font: 10pt SourceCodeVF;
                    }
                    QMenu{
                        border: 1px solid #0F7ABF;
                    }
                    QMenu::item{
                        background-color: #292424;
                        color: #1CB414;
                        font: 10pt SourceCodeVF;
                    }
                    QMenu::item:selected{
                        background-color: #0C3B0B;
                        color: #D59B0C;
                    }
                    QMenu::item:disabled{
                        background-color: #3D3A3D;
                        color: white;
                    }
                    QSpinBox{
                        background-color: #292424;
                        color: #1CB414;
                        padding-right: 15px;
                        border: 1px solid #0F7ABF;
                        font: 10pt SourceCodeVF;
                    }
                    QSpinBox::up-button{
                        width: 16px;
                    }
                    QSpinBox::down-button{
                        width: 16px;
                    }
                    QCalendarWidget QAbstractItemView:enabled{
                        background-color: #292424;
                        color: #1CB414;
                        selection-background-color: #1CB414; 
                        selection-color: #292424;
                        font: 10pt SourceCodeVF;
                    }
                    QCheckBox{
                        padding: 5px;
                        color: #1CB414;
                        font: 10pt SourceCodeVF;
                    }
                    QCheckBox::indicator{
                        border: 2px solid #0F7ABF;
                        width: 20px;
                        height: 20px;
                        border-radius: 12px;
                    }
                    QCheckBox::indicator:checked{
                        background-color: #0F7ABF;
                    }
                    QCheckBox::indicator:unchecked{
                        background-color: none;
                    }
                    """
    if name_of_style == "Clear Elegant Style":
        style_string = """
                    QWidget{
                        background-color: #EEEAEA;
                    }
                    QLabel{
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QLabel[accessibleName="an_title"]{
                        color: black;
                        font: bold 18pt Adobe New Century Schoolbook;
                    }
                    QLabel[accessibleName="an_section_title"]{
                        color: black;
                        font: bold 14pt Adobe New Century Schoolbook;
                    }
                    QLineEdit{
                        border: 1px solid black;
                        background-color: white;
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QPushButton{
                        background-color: #444342;
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QPushButton:hover{
                        background-color: #092B06;
                        color: white;
                    }
                    QComboBox{
                        background-color: #444342;
                        color: black;
                        padding: 1px 0px 1px 3px; /* Rende possibile il cambio del colore del testo nell'hover */
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QComboBox:hover{
                        background-color: #092B06;
                        color: white;
                    }
                    QComboBox:selected{
                        background-color: #444342;
                        color: black;
                    }
                    QTableWidget{
                        color: black;
                        border: 1px solid black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QTableWidget::item:selected{
                        background-color: #092B06;
                        color: white;
                    }
                    QTableWidget QTableCornerButton::section {
                        background-color: white;
                        border: 1px solid white;
                        border-style: outset;
                    }
                    QHeaderView{
                        background-color: #444342;
                        color: white;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QTextEdit{
                        color: black;
                        background-color: #EEEAEA;
                        border: 2px solid black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QMenu{
                        border: 1px solid black;
                    }
                    QMenu::item{
                        background-color: #444342;
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QMenu::item:selected{
                        background-color: #092B06;
                        color: white;
                    }
                    QMenu::item:disabled{
                        background-color: white;
                        color: #444342;
                    }
                    QSpinBox{
                        background-color: white;
                        color: black;
                        padding-right: 15px;
                        border: 1px solid black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QSpinBox::up-button{
                        width: 16px;
                    }
                    QSpinBox::down-button{
                        width: 16px;
                    }
                    QCalendarWidget QAbstractItemView:enabled{
                        background-color: white;
                        color: black;
                        selection-background-color: black; 
                        selection-color: white;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QCheckBox{
                        padding: 5px;
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QCheckBox::indicator{
                        border: 2px solid black;
                        width: 20px;
                        height: 20px;
                    }
                    QCheckBox::indicator:checked{
                        background-color: black;
                    }
                    QCheckBox::indicator:unchecked{
                        background-color: none;
                    }
                    """
    if name_of_style == "Dark Elegant Style":
        style_string = """
                    QWidget{
                        background-color: #444342;
                    }
                    QLabel{
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QLabel[accessibleName="an_title"]{
                        color: black;
                        font: bold 18pt Adobe New Century Schoolbook;
                    }
                    QLabel[accessibleName="an_section_title"]{
                        color: black;
                        font: bold 14pt Adobe New Century Schoolbook;
                    }
                    QLineEdit{
                        border: 1px solid black;
                        background-color: #444342;
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QPushButton{
                        background-color: black;
                        color: white;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QPushButton:hover{
                        background-color: #092B06;
                    }
                    QComboBox{
                        background-color: black;
                        color: white;
                        padding: 1px 0px 1px 3px; /* Rende possibile il cambio del colore del testo nell'hover */
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QComboBox:hover{
                        background-color: #092B06;
                    }
                    QComboBox:selected{
                        background-color: black;
                    }
                    QTableWidget{
                        color: black;
                        border: 1px solid black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QTableWidget::item:selected{
                        background-color: #092B06;
                        color: white;
                    }
                    QTableWidget QTableCornerButton::section {
                        background-color: #444342;
                        border: 1px solid #444342;
                        border-style: outset;
                    }
                    QHeaderView{
                        background-color: black;
                        color: white;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QTextEdit{
                        color: white;
                        background-color: #444342;
                        border: 2px solid black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QMenu{
                        border: 1px solid black;
                    }
                    QMenu::item{
                        background-color: #444342;
                        color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QMenu::item:selected{
                        background-color: #092B06;
                        color: white;
                    }
                    QMenu::item:disabled{
                        background-color: white;
                        color: #444342;
                    }
                    QSpinBox{
                        background-color: #444342;
                        color: white;
                        padding-right: 15px;
                        border: 1px solid black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QSpinBox::up-button{
                        width: 16px;
                    }
                    QSpinBox::down-button{
                        width: 16px;
                    }
                    QCalendarWidget QAbstractItemView:enabled{
                        background-color: black;
                        color: white;
                        selection-background-color: white; 
                        selection-color: black;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QCheckBox{
                        padding: 5px;
                        color: white;
                        font: 10pt Adobe New Century Schoolbook;
                    }
                    QCheckBox::indicator{
                        border: 2px solid black;
                        width: 20px;
                        height: 20px;
                    }
                    QCheckBox::indicator:checked{
                        background-color: white;
                    }
                    QCheckBox::indicator:unchecked{
                        background-color: none;
                    }
                    """
    return style_string