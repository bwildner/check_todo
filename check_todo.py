#check_todo.py 
# Dezember 2017 Bernd Wildner
# Oeffnet die ToDo Liste und sucht nach der angegebenen MNummer

#import time
import sys
import win32com.client
import logging
from PyQt4 import uic 
from PyQt4.QtGui import QMainWindow, QMessageBox, QApplication
from ConfigParser import SafeConfigParser


parser = SafeConfigParser()
parser.read('check_todo.ini')

url = parser.get('config', 'url') # hole Wert aus Abschnitt config Feld url
loglevel = parser.get('config', 'loglevel')





if loglevel == "INFO":
    logging.basicConfig(filename="check_todo.log",level = logging.INFO,format = "%(asctime)s [%(levelname)-8s] %(message)s")
elif loglevel == "DEBUG":
    logging.basicConfig(filename="check_todo.log",level = logging.DEBUG,format = "%(asctime)s [%(levelname)-8s] %(message)s")
else:
    app = QApplication(sys.argv)
    w = QMessageBox()
    w.setIcon(QMessageBox.Warning)
    w.setText("falscher Loglevel")
    w.setWindowTitle("Fehler ini File")
    w.show()
    sys.exit(app.exec_())

logging.info("****************Start Logging****************")


form_class = uic.loadUiType("check_todo_gui.ui")[0]                 # Load the UI

        

class MyWindowClass(QMainWindow, form_class):
    def __init__(self, parent=None):
        QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.button_eingabe.clicked.connect(self.eingabe_clicked) #Button oder Enter schliesst die Eingabe ab
        self.mnummer_textfeld.returnPressed.connect(self.eingabe_clicked)  # Bind the event handlers
        
               
        
    def eingabe_clicked(self):                  #  button event handler
        eingabe = str(self.mnummer_textfeld.text())
        
        #pruefen der eingabe und Fehlermeldung wenn noetig
        if not eingabe.isdigit() or len(eingabe)<>5:
            QMessageBox.warning(self, "Eingabefehler","Bitte die MNummer 5stellig eingeben.",QMessageBox.Cancel, QMessageBox.NoButton,QMessageBox.NoButton)
            return
        logging.info("Eingabe: "+ str(eingabe))
        
        suchen(eingabe)
        if len(ergebnis)==0:
            QMessageBox.warning(self, "Ergebnis","Nichts gefunden", QMessageBox.Cancel, QMessageBox.NoButton, QMessageBox.NoButton)
        else:
            app.quit() #GUI Schleife beenden
       
        
               
def suchespalte(register): 
    #sucht die Spalte mit der Ueberschrift M-Nr in dem Sheet/Register
    
    logging.info("suche MNummern Spalte")
    suchsheet = xlfile.Worksheets(register)
        
    for zeile in range(1,5): #suchen nur in den ersten 5 Zeilen und 10 Spalten
        for spalte in range (1,10):
            zelleunicode = unicode(suchsheet.Cells(zeile,spalte)) #fals unicode zeichen enthalten sind, gibt es so keinen Fehler
            zelleunicodeentf= zelleunicode.encode('utf8', 'replace')
            #print suchsheet.Cells(i,6)
            #zelle= str(suchsheet.Cells(i,6))
            zelle=str(zelleunicodeentf)
            suche= str("M-Nr") 
            
            if zelle.find(suche)>=0:
                logging.info ("suchspalte gefunden Zeile:"+str(zeile)+" Spalte:"+str(spalte))
                return spalte
    logging.info("suchespalte: nichts gefunden")
            

def suchen(mnummer):
    print "Suchen Start"
    logging.info("Suche Start")
 
    anzsheets = xlfile.Worksheets.Count # letztes Sheet Hilfe nicht durchsuchen
    #anzsheets=2
    for f in range(1,anzsheets):
        suchspalte= suchespalte(f) #in welcher Zeile stehen die MNummern
        suchsheet = xlfile.Worksheets(f)
        print "++++++++++++++++Suche in Sheet: "+xlfile.Worksheets(f).Name
        logging.info("Suche Start in sheet: "+str(xlfile.Worksheets(f).Name))
        logging.debug ("Anzahl Zeilen: "+ str(xlfile.Worksheets(f).UsedRange.Rows.Count))
 
        
        for i in range(1,xlfile.Worksheets(f).UsedRange.Rows.Count): #von 1 bis Ende Zeilen durchsuchen
            zelleunicode = unicode(suchsheet.Cells(i,suchspalte))
            zelleunicodeentf= zelleunicode.encode('utf8', 'replace')
            #print suchsheet.Cells(i,6)
            #zelle= str(suchsheet.Cells(i,6))
            zelle=str(zelleunicodeentf)
            suche= str(mnummer)
            #print zelle.find(suche)
            logging.debug(" Sheetnr:"+str(f)+" Zeile:"+ str(i)+" Zelle:"+str(zelle))
            
            if zelle.find(suche)>=0:
            
                print "Suchtext gefunden "+zelle
                logging.info("Suchtext gefunden"+str(zelle))
                ergebnis.extend((f,i))
                
                #break
            #else:
                #print suche+" ist Nicht gleich " + zelle
              

    logging.info("Suchen Ende, Ergebnis: "+str(ergebnis))        
    

def initxl():
    global xl, xlfile, app, myWindow
    global ergebnis
    ergebnis = []
    
    xl = win32com.client.DispatchEx('Excel.Application')
    xlfile = xl.Workbooks.Open(url, ReadOnly=False, IgnoreReadOnlyRecommended=True) 
    
    #xlfile = xl.Workbooks.Open('d:/todo.xls', ReadOnly=False, IgnoreReadOnlyRecommended=True) 
    app = QApplication(sys.argv)
    myWindow = MyWindowClass(None)

def zwischenablage():
    #zwischenablage auslesen, pruefen und in das Textfeld kopieren
    zwischenablage = str(QApplication.clipboard().text())
    if len(zwischenablage) >5:
        QApplication.clipboard().setText("")
    myWindow.mnummer_textfeld.setText(QApplication.clipboard().text())

    



##### Hauptprogramm #####

print "Prg Start"



initxl() #Init 
zwischenablage() #Zwischenablage pruefen



myWindow.show()
logging.info("show gui")


app.exec_() #start GUI Schleife

if len(ergebnis) == 2: #nur ein Ergebnis, Excel anzeigen und Zeile markieren
    xl.Visible = True 
    xl.WindowState = 2
    xlfile.Worksheets(ergebnis[0]).Activate() #gesuchtes Sheet aktivieren

    #xl.Range("1").Select()    
    xl.Rows(str(ergebnis[1])).Select() #gesuchte Zeile markieren
    xl.WindowState = 3 #maximiert das Fenster
    #zelle="H"+str(i)    
    #xl.Range(zelle).Select()
    logging.debug("Nur ein Ergebnis, Excel anzeigen und Zeile markieren ")    


if len(ergebnis) >2: #mehr als ein Ergebnis, neues Excel Workbook oeffnen und Zeilen kopieren
    
    logging.debug("mehr als ein Ergebnis, neues Excel Workbook oeffnen und Zeilen kopieren ")    

    wb = xl.Workbooks.Add()
    ws = wb.Worksheets.Add()
    ws.Name = "Ergebnisse"
    #wb.Worksheet(1).Rows(1).Value = xlfile.Worksheet(1).Rows(1)
    
    for i in range(1,len(ergebnis),2):
        logging.debug("Zaehler "+str(i))
        logging.debug("Sheet: "+str(ergebnis[i-1]))#listen zaehlen ab 0
        logging.debug("Zeile: "+str(ergebnis[i]))    

        xlfile.Worksheets(ergebnis[i-1]).Rows(ergebnis[i]).Copy()
        #xlfile.Worksheets(i).Rows(i).Copy()
        
        ws.Paste(ws.Rows(i))
        logging.debug(wb.Worksheets(1).Rows(i))

    #xl2.WindowState = 1
    xlfile.Close(True)
    xl.Visible = True
    
print "Programm Ende"    






 
 
