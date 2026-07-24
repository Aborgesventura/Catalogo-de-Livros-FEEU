#cd /d "J:\Catalogo de Livros FEEU 7.27"

pyinstaller --onedir --noconsole --name Catalogo_FEEU --icon=feeu.ico --collect-submodules=reportlab --collect-submodules=openpyxl --collect-submodules=odf --hidden-import=reportlab --hidden-import=odf Catalogo_de_Livros.py


#pyinstaller --onefile --noconsole --name Catalogo_FEEU --icon=feeu.ico --collect-submodules=reportlab --collect-submodules=openpyxl --collect-submodules=odf --hidden-import=reportlab --hidden-import=odf Catalogo_de_Livros.py