'===================================================
'VBA Macro to export employee details
'from Outlook to an Excel file for analysis
'Author: Rajathithan Rajasekar
'Date: 03/13/2021
'===================================================

Option Explicit

'Excel Variables
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim SFilename, c, g, cl As String
Dim Count As Integer
    
' Employee variables
Global empCName, empFName, empLName, empDepartment, empTitle, empOffice, _
 empCity, empAlias, empEmail, supFname, supLname, supAlias, supEmail, _
empBPhone, empMPhone As String
Sub extract_employees_info()

'Error Handler
 'On Error Resume Next
 
 'Specify the directory path in which you want to save the file
 SFilename = "C:\Employees-Data.xlsx"
 
 'row count
 Count = 1
 
 'Set Excel object
 Set objExcel = CreateObject("Excel.Application")
 
 'If no file available create one
 If Dir(SFilename) <> "" Then
     Set objWorkbook = objExcel.Workbooks.Open(SFilename)
 Else
     Set objWorkbook = objExcel.Workbooks.Add
     objWorkbook.SaveAs SFilename
     Set objWorkbook = objExcel.Workbooks.Open(SFilename)
 End If
 
 'Set the Column Names
 cl = clearExcel(objWorkbook)
 
 c = updateExcel(objWorkbook, Count)
 
 'Outlook's address entires
 Dim usersList As Outlook.AddressEntries
 
 'Outlook's address entry
 Dim oEntry As Outlook.AddressEntry

 'Powerful command - get all the user details
 Set usersList = Outlook.Application.Session.AddressLists.Item("All Users"). _
 AddressEntries
 
 'Using the GetNext method before the GetNext method
 Set oEntry = usersList.GetFirst
 
 'To know the item that is being processed
 Debug.Print String(65535, vbCr)
 
 g = getEmployeeInfo(oEntry, Count)
 
 'Increasing the count
 Count = Count + 1
 
 'Adding the first entry to excel
 c = UpdateExcelContent(objWorkbook, Count)
 
 'Increasing the row count
 Count = Count + 1
 
 'Getting the next employee details from the userlist
 Set oEntry = usersList.GetNext
 
 'Until the entry is set to nothing it executes
 Do While oEntry <> ""
     
     g = getEmployeeInfo(oEntry, Count)
 
     'Update Excel
     c = UpdateExcelContent(objWorkbook, Count)
     
     'Get the next employee details
     Set oEntry = usersList.GetNext
     
     'Increase the count
     Count = Count + 1
     
     'My Program is very slow, at the max it can process only 50 to 60
     'employee recordsper minute, So if want you can restrict the
     'count to 100
     If Count = 11 Then
         Exit Do
     End If
 Loop
 
 
 'Save and close everything
 objWorkbook.Save
 objWorkbook.Close
 Set objWorkbook = Nothing
 objExcel.Quit
 Set objExcel = Nothing
 'Debug.Print "Analysis Completed !!"
    
End Sub

Function updateExcel(objWorkbook As Excel.Workbook, Count)

With objWorkbook.Sheets("Sheet1")
        .Cells(Count, 1).Value = "S.NO"
        .Cells(Count, 2).Value = "Company Name"
        .Cells(Count, 3).Value = "Employee First Name"
        .Cells(Count, 4).Value = "Employee Last Name"
        .Cells(Count, 5).Value = "Employee Department"
        .Cells(Count, 6).Value = "Employee JobTitle"
        .Cells(Count, 7).Value = "Employee Office Location"
        .Cells(Count, 8).Value = "Employee City"
        .Cells(Count, 9).Value = "Employee Alias"
        .Cells(Count, 10).Value = "Employee Email Address"
        .Cells(Count, 11).Value = "Supervisor FirstName"
        .Cells(Count, 12).Value = "Supervisor LastName"
        .Cells(Count, 13).Value = "Supervisor Alias"
        .Cells(Count, 14).Value = "Supervisor Email Address"
        .Cells(Count, 15).Value = "BusinessTelephoneNumber"
        .Cells(Count, 16).Value = "HomeTelephoneNumber"
        .Columns.AutoFit
End With

End Function

Function clearExcel(objWorkbook As Excel.Workbook)

With objWorkbook.Sheets("Sheet1")
        .Cells.ClearContents
End With

End Function


Function UpdateExcelContent(objWorkbook As Excel.Workbook, Count)

With objWorkbook.Sheets("Sheet1")
        .Cells(Count, 1).Value = Count - 1
        .Cells(Count, 2).Value = empCName
        .Cells(Count, 3).Value = empFName
        .Cells(Count, 4).Value = empLName
        .Cells(Count, 5).Value = empDepartment
        .Cells(Count, 6).Value = empTitle
        .Cells(Count, 7).Value = empOffice
        .Cells(Count, 8).Value = empCity
        .Cells(Count, 9).Value = empAlias
        .Cells(Count, 10).Value = empEmail
        .Cells(Count, 11).Value = supFname
        .Cells(Count, 12).Value = supLname
        .Cells(Count, 13).Value = supAlias
        .Cells(Count, 14).Value = supEmail
        .Cells(Count, 15).Value = empBPhone
        .Cells(Count, 16).Value = empMPhone
        .Columns.AutoFit
End With

End Function

Function getEmployeeInfo(oEntry As Outlook.AddressEntry, Count As Integer)
    On Error GoTo ErrHandler
    If Count = 1 Then
        Debug.Print "Processing Employee:" & Count
    Else
        Debug.Print "Processing Employee:" & Count - 1
    End If
    'Get details
    'extracting the first employee details from the users list
    If oEntry.GetExchangeUser().CompanyName = "" Then empCName = "NA" Else: _
    empCName = oEntry.GetExchangeUser().CompanyName
    If oEntry.GetExchangeUser().FirstName = "" Then empFName = "NA" Else: _
    empFName = oEntry.GetExchangeUser().FirstName
    If oEntry.GetExchangeUser().LastName = "" Then empLName = "NA" Else: _
    empLName = oEntry.GetExchangeUser().LastName
    
    Debug.Print ("Name: " & empFName & " " & empLName)
    
    If oEntry.GetExchangeUser().Department = "" Then empDepartment = "NA" Else: _
    empDepartment = oEntry.GetExchangeUser().Department
    If oEntry.GetExchangeUser().JobTitle = "" Then empTitle = "NA" Else: _
    empTitle = oEntry.GetExchangeUser().JobTitle
    If oEntry.GetExchangeUser().OfficeLocation = "" Then empOffice = "NA" Else: _
    empOffice = oEntry.GetExchangeUser().OfficeLocation
    
    If oEntry.GetExchangeUser().City = "" Then empCity = "NA" Else: _
    empCity = oEntry.GetExchangeUser().City
    If oEntry.GetExchangeUser().Alias = "" Then empAlias = "NA" Else: _
    empAlias = oEntry.GetExchangeUser().Alias
    If oEntry.GetExchangeUser().PrimarySmtpAddress = "" Then empEmail = "NA" Else: _
    empEmail = oEntry.GetExchangeUser().PrimarySmtpAddress
    
    If oEntry.GetExchangeUser().BusinessTelephoneNumber = "" Then empBPhone = "NA" Else: _
    empBPhone = oEntry.GetExchangeUser().BusinessTelephoneNumber
    'Debug.Print (oEntry.GetExchangeUser().BusinessTelephoneNumber)
    If oEntry.GetExchangeUser().MobileTelephoneNumber = "" Then empMPhone = "NA" Else: _
    empMPhone = oEntry.GetExchangeUser().MobileTelephoneNumber
    'Debug.Print (oEntry.GetExchangeUser().MobileTelephoneNumber)
    
    If oEntry.GetExchangeUser().GetExchangeUserManager().FirstName = "" Then supFname = "NA" Else: _
    supFname = oEntry.GetExchangeUser().GetExchangeUserManager().FirstName
    If oEntry.GetExchangeUser().GetExchangeUserManager().LastName = "" Then supLname = "NA" Else: _
    supLname = oEntry.GetExchangeUser().GetExchangeUserManager().LastName
    If oEntry.GetExchangeUser().GetExchangeUserManager().Alias = "" Then supAlias = "NA" Else: _
    supAlias = oEntry.GetExchangeUser().GetExchangeUserManager().Alias
    If oEntry.GetExchangeUser().GetExchangeUserManager().PrimarySmtpAddress = "" Then supEmail = "NA" Else: _
    supEmail = oEntry.GetExchangeUser().GetExchangeUserManager().PrimarySmtpAddress
    
    
    Exit Function

    
ErrHandler:
            supFname = "NA"
            supLname = "NA"
            supAlias = "NA"
            supEmail = "NA"
    
    
End Function

