import time
import win32com.client
 
#----------------------------------------------------------------------
#xl = win32.gencache.EnsureDispatch('Excel.Application')
xl = win32com.client.DispatchEx('Excel.Application')

ss = xl.Workbooks.Open('e:/todo.xls', ReadOnly=False, IgnoreReadOnlyRecommended=True)    
sh = ss.ActiveSheet    
    
xl.Visible = True 
time.sleep(1)    

    

for i in range(1,10):
    print sh.Cells(i,4)
    zelle= str(sh.Cells(i,4))
    if zelle=="53000.0":
        print "Hello"
        
zelle="A"+str(100)    
xl.Range(zelle).Select()    
time.sleep(5)    
    
#ss.Close(False) 
#xl.Application.Quit()    
    
 
