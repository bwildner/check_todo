#check_todo.py 
# Dezember 2017 Bernd Wildner
# Oeffnet die ToDo Liste und sucht nach der angegebenen MNummer

import time
import sys
import win32com.client
import logging
from PyQt4 import uic 
from PyQt4.QtGui import QMainWindow, QMessageBox, QApplication


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
        logging.debug ("Anzahl Zeilen: "+xlfile.Worksheets(f).UsedRange.Rows.Count)
 
        
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
                #break
            #else:
                #print suche+" ist Nicht gleich " + zelle
              

    print "Suchen Ende"        
    xl.Visible = True  

    zelle="H"+str(i)    
    xl.Range(zelle).Select()    
    time.sleep(10)    


def initxl():
    global xl, xlfile, erl, sh, app, myWindow
    xl = win32com.client.DispatchEx('Excel.Application')
    xlfile = xl.Workbooks.Open('e:/todo.xls', ReadOnly=False, IgnoreReadOnlyRecommended=True) 
    erl= xlfile.Worksheets("Erledigt")
    xlfile.Worksheets("Erledigt").Activate()
    sh = xlfile.ActiveSheet
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
logging.basicConfig(filename="check_todo.log",level = logging.INFO,format = "%(asctime)s [%(levelname)-8s] %(message)s")
logging.info("****************Start Logging****************")



initxl() #Init 
zwischenablage() #Zwischenablage pruefen



myWindow.show()


app.exec_()


#xlfile.Close(False) 
xl.Application.Quit()    


print "Programm Ende"    






 
 
