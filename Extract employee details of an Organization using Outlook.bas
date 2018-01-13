Attribute VB_Name = "Module11"
Sub extract_employees_info()

'===============================================================
'Macro to export all the employee details of an Organization
'from Outlook to an Excel file for analysis
'Coder: Rajathithan Rajasekar
'GitHub: https://github.com/rajathithan/EXCEL-VBA
'WebSite: https://www.gadoth.com
'Facebook: https://www.facebook.com/gadoth/
'===============================================================

'Error Handler
 On Error Resume Next

    'Excel Variables
    Dim objExcel As Excel.Application
    Dim objWorkbook As Excel.Workbook
    
    'Make sure you have D drive, because it creates the file in
    'D drive
    SFilename = "D:\orgUserList.xlsx"
    
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
    With objWorkbook.Sheets("Sheet1")
        .Cells.ClearContents
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
    End With

    'Outlook's address entires
    Dim usersList As Outlook.AddressEntries
    
    'Outlook's address entry
    Dim oEntry As Outlook.AddressEntry
   
    'Powerful command - get all the user details
    Set usersList = Outlook.Application.Session.AddressLists.item("All Users").AddressEntries
    
    'Using the GetNext method before the GetNext method
    Set oEntry = usersList.GetFirst
    
    'extracting the first employee details from the users list
    empCName = oEntry.GetExchangeUser().CompanyName
    empFName = oEntry.GetExchangeUser().FirstName
    empLName = oEntry.GetExchangeUser().LastName
    empDepartment = oEntry.GetExchangeUser().Department
    empTitle = oEntry.GetExchangeUser().JobTitle
    empOffice = oEntry.GetExchangeUser().OfficeLocation
    empCity = oEntry.GetExchangeUser().City
    empAlias = oEntry.GetExchangeUser().Alias
    empEmail = oEntry.GetExchangeUser().PrimarySmtpAddress
    supFname = oEntry.GetExchangeUser().GetExchangeUserManager().FirstName
    supLname = oEntry.GetExchangeUser().GetExchangeUserManager().LastName
    supAlias = oEntry.GetExchangeUser().GetExchangeUserManager().Alias
    supEmail = oEntry.GetExchangeUser().GetExchangeUserManager().PrimarySmtpAddress
    
    'Increasing the count
    Count = Count + 1
    
    'Adding the first entry to excel
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
    End With
    
    'Increasing the row count
    Count = Count + 1
    
    'Getting the next employee details from the userlist
    Set oEntry = usersList.GetNext
    
    'Until the entry is set to nothing it executes
    Do While oEntry <> ""
        'To know the item that is being processed
        Debug.Print "Processing item:" & Count
        'Get details
        empCName = oEntry.GetExchangeUser().CompanyName
        empFName = oEntry.GetExchangeUser().FirstName
        empLName = oEntry.GetExchangeUser().LastName
        empDepartment = oEntry.GetExchangeUser().Department
        empTitle = oEntry.GetExchangeUser().JobTitle
        empOffice = oEntry.GetExchangeUser().OfficeLocation
        empCity = oEntry.GetExchangeUser().City
        empAlias = oEntry.GetExchangeUser().Alias
        empEmail = oEntry.GetExchangeUser().PrimarySmtpAddress
        
        'If supervisor details are present it will extract the details
        If (IsError(supFname = oEntry.GetExchangeUser().GetExchangeUserManager().FirstName)) Then
            supFname = ""
            supLname = ""
            supAlias = ""
            supEmail = ""
        Else
            supFname = oEntry.GetExchangeUser().GetExchangeUserManager().FirstName
            supLname = oEntry.GetExchangeUser().GetExchangeUserManager().LastName
            supAlias = oEntry.GetExchangeUser().GetExchangeUserManager().Alias
            supEmail = oEntry.GetExchangeUser().GetExchangeUserManager().PrimarySmtpAddress
        End If
    
        'Update Excel
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
            .Columns.AutoFit
        End With
        
        'Get the next employee details
        Set oEntry = usersList.GetNext
        
        'Increase the count
        Count = Count + 1
        
        'My Program is very slow, at the max it can process only 50 to 60 items
        'per minute, So i have restricted the extraction count to 100
        If Count = 100 Then
            Exit Do
        End If
    Loop
    
    
    'Save and close everything
    objWorkbook.Save
    objWorkbook.Close
    Set objWorkbook = Nothing
    objExcel.Quit
    Set objExcel = Nothing
    Debug.Print "Analysis Completed !!"
    
End Sub
