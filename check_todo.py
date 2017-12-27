import time
import sys
import win32com.client
import logging
from PyQt4 import uic 
from PyQt4.QtGui import QWidget, QMainWindow, QVBoxLayout, QTextEdit, QMessageBox, QApplication




logging.basicConfig(filename="check_todo.log",level = logging.DEBUG,format = "%(asctime)s [%(levelname)-8s] %(message)s")

 
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
               
        
           
 
 
app = QApplication(sys.argv)
myWindow = MyWindowClass(None)

#zwischenablage auslesen, pruefen und in das Textfeld kopieren
zwischenablage = str(QApplication.clipboard().text())
if len(zwischenablage) >5:
    QApplication.clipboard().setText("")
myWindow.mnummer_textfeld.setText(QApplication.clipboard().text())


myWindow.show()


app.exec_()







 
#----------------------------------------------------------------------
#xl = win32.gencache.EnsureDispatch('Excel.Application')
print "Prg Start"
xl = win32com.client.DispatchEx('Excel.Application')

ss = xl.Workbooks.Open('e:/todo.xls', ReadOnly=False, IgnoreReadOnlyRecommended=True)    

erl= ss.Worksheets("Erledigt")

ss.Worksheets("Erledigt").Activate()    
sh = ss.ActiveSheet

print "Suchen Start"

print ss.Worksheets(1).Name   
print ss.Worksheets(2).Name    
print ss.Worksheets(3).Name    
print ss.Worksheets(4).Name    
print ss.Worksheets(5).Name    

 
print ss.Worksheets.Count

time.sleep(10)

for i in range(1,1500):
    
    #print sh.Cells(i,8)
    print erl.Cells(i,8)
    zelle= str(erl.Cells(i,8))
    suche= str(56040)
    print zelle.find(suche)
    if zelle.find(suche)>=0:
        
        print "Suchtext gefunden "+zelle
        break
    else:
        print suche+" ist Nicht gleich " + zelle
              

print "Suchen Ende"        
xl.Visible = True  

zelle="H"+str(i)    
xl.Range(zelle).Select()    
time.sleep(10)    
ss.Close(False) 
xl.Application.Quit()    
print "Programm Ende"    
 
