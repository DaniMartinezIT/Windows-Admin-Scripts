Set objOutlook = CreateObject("Outlook.Application")
Set objMail = objOutlook.CreateItem(0)

Dim variable1
Dim variable2
Dim variableN

variable1 = InputBox("******:")
variable2 = InputBox("******:")
variableN = InputBox("******:")


  
   objMail.Display   'Mostrar mensaje por pantalla
   objMail.Recipients.Add ("xxx@xxx.xxx")
   objMail.Recipients.Add ("aaa@aaa.aaa")
    
   objMail.Subject = "Titulo:  "+ variable1
   objMail.htmlBody = 
   Set objMail = Nothing
   Set objOutlook = Nothing
