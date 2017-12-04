import time
import win32com.client
 
#----------------------------------------------------------------------
#xl = win32.gencache.EnsureDispatch('Excel.Application')
print "Prg Start"
xl = win32com.client.DispatchEx('Excel.Application')

ss = xl.Workbooks.Open('e:/todo.xls', ReadOnly=False, IgnoreReadOnlyRecommended=True)    

erl= ss.Worksheets("Erledigt")

ss.Worksheets("Erledigt").Activate()    
sh = ss.ActiveSheet

print "Suchen Start"
    

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
 
