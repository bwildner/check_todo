import time
import win32com.client
 
#----------------------------------------------------------------------
#xl = win32.gencache.EnsureDispatch('Excel.Application')
xl = win32com.client.DispatchEx('Excel.Application')

ss = xl.Workbooks.Open('d:/todo.xls', ReadOnly=False, IgnoreReadOnlyRecommended=True)    
sh = ss.ActiveSheet    
    
xl.Visible = True 
time.sleep(1)    

    
sh.Cells(1,1).Value = 'Hacking Excel with Python Demo' 
    
time.sleep(1) 
for i in range(2,8):    
    sh.Cells(i,1).Value = 'Line %i' % i    
    time.sleep(0)    
        
zelle="A"+str(100)    
xl.Range(zelle).Select()    
time.sleep(5)    
    
ss.Close(False) 
xl.Application.Quit()    
    
 
