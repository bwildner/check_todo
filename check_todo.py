import time
import win32com.client
 
#----------------------------------------------------------------------
#xl = win32.gencache.EnsureDispatch('Excel.Application')
xl = win32com.client.DispatchEx('Excel.Application')

ss = xl.Workbooks.Open('d:/todo.xls', ReadOnly=False, IgnoreReadOnlyRecommended=True)    
sh = ss.ActiveSheet
erl= ss.Worksheets("Erledigt")
    
xl.Visible = True 
time.sleep(1)    


    

for i in range(1,50):
    #print sh.Cells(i,8)
    #print erl.Cells(i,8)
    zelle= str(erl.Cells(i,8))
    if zelle.find('53000')>0:
        
        print "Suchtext gefunden"+str(erl.Cells(i,8))
        break
              
        
zelle="A"+str(100)    
xl.Range(zelle).Select()    
time.sleep(5)    
    
#ss.Close(False) 
#xl.Application.Quit()    
    
 
