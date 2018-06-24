'***************************************************************
'Server Check script
'Created By: Brian Bohanon
'-This script reads in a list of servers from a text file and
'-queries each server using WMI for hard disk space and service
'-status.
'***Customizations****
'strEngineer = Email address of engineer running script
'strDirectory = Directory of server list and script
'srcFileName = name of file containing list of servers
'strLocation = office location
'***************************************************************

'Global variable to determine who to mail the report to
Dim strMailFlag, strEngineer, strDirectory, strFileName

'Change to the address of the engineer running the script
strEngineer = "engineer@myjob.com"
'Change to the directory where the script and server list live
strDirectory = "D:\Scripts\"
'Change to the appropriate location
strLocation = "My Office"
'Excel file name
strFileName = strDirectory & strLocation & "_Server_Checks_" & Month(Date()) & "_" & Day(Date()) & "_" & Year(Date()) & " " & Right(Time(),2) & ".xls"
'Change this variable to another location if needed
srcFileName = "servers.txt"

CreateWorkbook()
'SendAttach()
'DeleteFile()
MsgBox "Complete"

Sub CreateWorkbook()

   Dim disk_size, disk_free
   Dim m,n
   Set objExcel = CreateObject("Excel.Application")
   Set objWorkbook = objExcel.Workbooks.Add()

   n = 1
   m = 1

   'Column headers
   objExcel.Cells(m, n) = "Server Name"
   objExcel.Cells(m, n).Font.Bold = True
   n = n + 1
   objExcel.Cells(m, n) = "Free Space"
   objExcel.Cells(m, n).Font.Bold = True
   n = n + 1
   objExcel.Cells(m, n) = "Services"
   objExcel.Cells(m, n).Font.Bold = True
   n = n + 1

   'Open File of server names -------------------------------------
   i = 0 
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   'Open the text file for reading
   Set objFile = objFSO.OpenTextFile(srcFileName, 1) 
   Do Until objFile.AtEndOfStream 
    Redim Preserve arrFileLines(i) 
    arrFileLines(i) = objFile.ReadLine 
    i = i + 1 
   Loop 

   objFile.Close 
   '---------------------------------------------------------------
   n = 1
   m = 3
   'For each server name get info and put into worksheet
   For l = Ubound(arrFileLines) to LBound(arrFileLines) Step -1 
    'set computer to the current index in the array
    strComputer = arrFileLines(l)
    'connect to the computer's WMI service
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
     If Err <> 0 Then
      DisplayErrorInfo()
      objExcel.Quit
     End If

    objExcel.Cells(m, n) = strComputer
    objExcel.Cells(m, n).Font.Bold = True
    j = m
    m = m + 1
    '-----------------------------------------------------------
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
    Set colServices = objWMIService.ExecQuery("Select * From Win32_Service")

    For each objDisk in colDisks
     objExcel.Cells(m, n) = objDisk.DeviceID
     n = n + 1
     'Convert into GB
     If objDisk.Size > 0 Then
      disk_size = objDisk.Size / 1073741824
      disk_free = objDisk.FreeSpace / 1073741824
      If disk_free > 0 Then
       strPercent = Round((int(disk_free)/int(disk_size)*100),2)
       'Check the percentage of the disk free space - if less than 10% free, only mail to engineer, not team
        If strPercent < 10 Then          
        	strMailFlag = 0         
        Else          
        	strMailFlag = 1         
        End If       
       Else         
       	strMailFlag = 0       
       End If       
     'Write to Excel       
     objExcel.Cells(m, n) = "Total: " & int(disk_size) & "GB" & " Free: " & int(disk_free) & "GB" & " (" & strPercent & "%)"      
     End If      
     'Move to next cell      
     n = 1      
     m = m + 1     
     Next     
     'Check the status of predetermined services-------------------     
     'New services can be added based on the service name     
     For Each objService in colServices      
     	If InStr(objService.Name, "MSExchangeIS") Then       
     		objExcel.Cells(j, 3) = objService.Name       
     		objExcel.Cells(j, 4) = objService.State       
     		If objService.State = "Stopped" Then        
     			strmailflag = 0       
     		End if       
     		j = j + 1      
     	ElseIf InStr(objService.Name, "MSExchangeMGMT") Then       
     		objExcel.Cells(j, 3) = objService.Name       
     		objExcel.Cells(j, 4) = objService.State       
     		If objService.State = "Stopped" Then        
     			strmailflag = 0       
     		End if       
     		j = j + 1      
     	ElseIf InStr(objService.Name, "MSExchangeMTA") Then       
     		objExcel.Cells(j, 3) = objService.Name       
     		objExcel.Cells(j, 4) = objService.State       
     		If objService.State = "Stopped" Then        
     			strmailflag = 0       
     		End if       
     		j = j + 1      
     	ElseIf InStr(objService.Name, "MSExchangeSA") Then       
     		objExcel.Cells(j, 3) = objService.Name       
     		objExcel.Cells(j, 4) = objService.State       
     		If objService.State = "Stopped" Then        
     			strmailflag = 0       
     		End if       
     		j = j + 1      
     	ElseIf InStr(objService.Name, "IISADMIN") Then       
     		objExcel.Cells(j, 3) = objService.Name       
     		objExcel.Cells(j, 4) = objService.State       
     		If objService.State = "Stopped" Then        
     			strmailflag = 0       
     		End if       
     		j = j + 1      
     	ElseIf InStr(objService.Name, "W3SVC") Then       
     		objExcel.Cells(j, 3) = objService.Name       
     		objExcel.Cells(j, 4) = objService.State       
     		If objService.State = "Stopped" Then        
     			strmailflag = 0       
     		End if       
     		j = j + 1      
     	End If     
     	Next     
     	Set objService = Nothing     
     	Set objDisk = Nothing     
     	m = m + 3    
     Next    
     '--------------------------------------------------------------    
     ' Autofit the first column to fit the longest service name    
     objExcel.Columns("A:Z").EntireColumn.AutoFit    
     'Delete remaining worksheets    
     objExcel.Worksheets("Sheet2").Delete    
     objExcel.Worksheets("Sheet3").Delete    
     'Save    
     objWorkbook.SaveAs strDirectory & strLocation & "_Server_Checks_" & Month(Date()) & "_" & Day(Date()) & "_" & Year(Date()) & " " & Right(Time(),2) & ".xls", 56    
     'Close Excel    
     objExcel.Quit    
     Set objExcel = Nothing    
     Set objFSO = Nothing    
     Set objWMIService = Nothing 
     End Sub 
     
     'Create a mail message and send it via Outlook sub SendAttach()    
     'Open mail, adress, attach report    
     Dim objOutlk    
     Dim objMail    
     Dim strMsg    
     Const olMailItem = 0    
     'Create a new message     
     Set objOutlk = createobject("Outlook.Application")     
     Set objMail = objOutlk.createitem(olMailItem)     
     If strMailFlag = 0 Then      
     	objMail.To = strEngineer      
     	objMail.Importance = 2     
     Else      
     	objMail.To = "boss@myjob.com"      
     	objMail.cc = "myteam@myjob.com" 
     	'Enter an address here To include a carbon copy; bcc is For blind carbon copy's      
     	'objMail.bcc = ""     
     End if    
     'Set up Subject Line    
     objMail.subject = "Server Check " & strLocation & " " & Month(Date()) & "_" & Day(Date()) & "_" & Year(Date()) & " " & Right(Time(),2)    
     objMail.attachments.add(strFileName)    
     objMail.Send    
     'Clean up    
     Set objMail = nothing    
     'Set objOutlk = nothing end sub 
     'Delete the file after sending 
     Sub DeleteFile()    
     	Set objFSO = CreateObject("Scripting.FileSystemObject")    
     	If objFSO.FileExists(strFileName) Then     
     		objFSO.DeleteFile strFileName    
     	End if    
     	Set objFSO = nothing 
     End Sub 
     	
     Sub DisplayErrorInfo    
     	WScript.Echo "Error: : " & Err    
     	WScript.Echo "Error (hex) : &H" & Hex(Err)    
     	WScript.Echo "Source : " & Err.Source    
     	WScript.Echo "Description : " & Err.Description    
     	Err.Clear 
     End Sub

