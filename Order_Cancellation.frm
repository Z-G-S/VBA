VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Order_Cancellation 
   Caption         =   "Order Cancellation"
   ClientHeight    =   3843
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   10020
   OleObjectBlob   =   "Order_Cancellation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Order_Cancellation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Order Cancellation Userform
' This module handles various functionalities related to the Order Cancellation process.
' It includes initialising the form with necessary data, submitting cancellation requests,
' And providing options to reset or close the order cancellation form.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Initialisation and Setup '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub userform_initialize()
    ' This subroutine is called when the userform is initialized.
    ' It populates the DateComboBox, sets up a named range "Reason" from Support_Data sheet,
    ' And initializes the cmbCancellationReason and txtEmpName controls.

    ' Populate the DateComboBox with dates.
    PopulateDateComboBox Me.DateComboBox
    
    ' Set up a named range "Reason" from the Support_Data sheet.
    Sheets("Support_Data").Range("G2", Sheets("Support_Data").Range("G" & Application.Rows.Count).End(xlUp)).Name = "Reason"
    
    ' Set the RowSource for cmbCancellationReason to the named range "Reason".
    cmbCancellationReason.RowSource = "Reason"
    
    ' Initialize cmbCancellationReason and txtEmpName controls.
    cmbCancellationReason.value = ""
    ' Automatically fill in the employee's name from the system environment.
    txtEmpName.value = Environ("username")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Form Control and Navigation '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub CloseButton_Click()
    ' Prompts the user to confirm if they want to close the userform, if confirmed, the form will be closed.
    
    ' Declare a variable to capture the user's response
    Dim i As VbMsgBoxResult
    
    ' Display a confirmation message box
    i = MsgBox("Do you want to close the Order Cancellation Form?", vbYesNo + vbQuestion, "Close Order Cancellation Form")
    
    ' If the user selects 'No', exit the subroutine without closing the form
    If i = vbNo Then Exit Sub
    
    ' Unload the form if the user selects 'Yes'
    Unload Me
End Sub

Private Sub ResetButton_Click()
    ' After confirmation this subroutine resets all Input fields except the customer name and email
    ' (as they are pulled from the data sheets).

    ' Declare a variable to capture the user's response.
    Dim i As VbMsgBoxResult

    ' Display a confirmation message box.
    i = MsgBox("Do you want to reset the Order Cancellation Form?", vbYesNo + vbQuestion, "Reset Order Cancellation Form")

    ' If the user selects 'No', exit the subroutine without resetting the form.
    If i = vbNo Then Exit Sub

    ' Loop through each control on the form and reset their values.
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        ' Check if the control is a TextBox or ComboBox, and is not the customer name or email fields.
        If (TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox") And ctrl.Name <> "txtCustName" And ctrl.Name <> "txtEmail" Then
            ' Reset the value of the control.
            ctrl.value = ""
        End If
    Next ctrl
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Data Entry and Updates '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdSubmit_Click()
    ' This subroutine handles the submission of cancellation data from a UserForm.
    ' It performs data validation, copies the data to the "Cancelled" worksheet, and updates related information.
    ' It also displays appropriate messages to the user for confirmation and validation.

    ' Display a confirmation message to ensure the user wants to submit the data.
    Dim msgResult As VbMsgBoxResult
    msgResult = MsgBox("Do you want to submit the data?", vbYesNo + vbQuestion, "Order Cancellation")

    ' If the user clicks "No," exit the subroutine.
    If msgResult = vbNo Then Exit Sub

    ' Data Validation - Check for empty input fields.
    Dim Inputs As Collection
    Set Inputs = New Collection
    Dim InputBoxes() As Variant
    Dim var As Variant
    Dim isError1 As Boolean

    ' Define an array of input fields to check.
    InputBoxes = Array(DateComboBox, txtEmpName, txtCustName, txtEmail, cmbCancellationReason, txtNotes)

    ' Loop through the input fields to check for empty values.
    For Each var In InputBoxes
        If var.value = "" Then
            var.BackColor = vbRed
            isError1 = True
        Else
            var.BackColor = vbWhite
        End If
    Next var

    ' If any input fields are empty, display an error message and exit the subroutine.
    If isError1 Then
        MsgBox "Please enter an input for the boxes that are highlighted red.", vbCritical + vbOKOnly, "No Input Error 1001"
        Exit Sub
    End If

    ' Check if the entered cancellation reason is valid.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Support_Data")
    Dim acceptableRange As Range
    Set acceptableRange = ws.Range("G2:G9")
    Dim enteredValue As String
    enteredValue = cmbCancellationReason.value

    ' Verify that the entered cancellation reason is in the acceptable range.
    If Application.WorksheetFunction.CountIf(acceptableRange, enteredValue) = 0 Then
        cmbCancellationReason.BackColor = vbRed
        MsgBox "Please ensure that the Cancellation Reason matches the reasons from the list.", vbCritical + vbOKOnly, "Validation Error 5001"
        Exit Sub
    Else
        cmbCancellationReason.BackColor = vbWhite
    End If
    
    ' Check if the user has inputted a Date in the future.
    ' This approach directly compares date values without formatting them as strings.
    Dim inputDate As Date
    Dim currentDate As Date
    
    ' Parse the input date from the DateComboBox.
    inputDate = CDate(DateComboBox.value)
    
    ' Get the current date.
    currentDate = Date ' 'Date' function returns the current system date.
    
    ' Compare the input date with the current date.
    If inputDate > currentDate Then
        Dim v As VbMsgBoxResult
        v = MsgBox("The Date you have inputted is in the future, are you sure you want to submit?", vbYesNo + vbQuestion, "Future Date")
        If v = vbNo Then Exit Sub  ' Exit if the user chooses not to proceed.
    End If
    
    ' Define variables for working with worksheets and rows.
    Dim iRow As Long
    Dim sh As Worksheet
    Dim lstRow As Long
    Dim orderDate As Date
    Dim cancellationDate As Date
    
    ' Set the target worksheet.
    Set sh = ThisWorkbook.Sheets("Cancelled")

    ' Find the next available row in the "Cancelled" worksheet.
    iRow = sh.Range("A" & Rows.Count).End(xlUp).Row + 1

    ' Access the selected row in the "lstCompleted" list box.
    With ASDEVS.lstCompleted
        lstRow = .ListIndex
        
    ' Parse the order date from the selected row in the "lstCompleted" list box.
    ' Assuming that the date is in the first column (index 0) and is properly formatted.
    orderDate = CDate(Format(.List(lstRow, 0), "dd/mm/yyyy"))

    ' Parse the cancellation date from the Order_Cancellation form.
    ' Assuming that the date is properly formatted in the ComboBox.
    cancellationDate = CDate(Order_Cancellation.DateComboBox.value)

    ' Check if the cancellation date is earlier than the order date.
    ' If cancellation date is earlier than the order date, prompt the user to confirm the entry,
    ' This helps avoid entry error, but does not block the entry outright to provide flexibility,
    ' E.g. stock ran out the day before the order.
    If cancellationDate < orderDate Then
        Dim vv As VbMsgBoxResult
        vv = MsgBox("The cancellation date is earlier than the order date, are you sure you want to proceed?", vbYesNo + vbQuestion, "Cancellation Date Before Order Date")
        If vv = vbNo Then Exit Sub  ' Exit if the user chooses not to proceed.
    End If
        
        ' Copy data from the selected row in "lstCompleted" to the "Cancelled" worksheet.
        ' And include additional information such as cancellation date, employee name/ID, etc.
        sh.Range("A" & iRow).value = Format(.List(lstRow, 0))
        sh.Range("B" & iRow).value = .List(lstRow, 1)
        sh.Range("C" & iRow).value = .List(lstRow, 2)
        sh.Range("D" & iRow).value = .List(lstRow, 3)
        sh.Range("E" & iRow).value = .List(lstRow, 4)
        sh.Range("F" & iRow).value = .List(lstRow, 5)
        sh.Range("G" & iRow).value = .List(lstRow, 6)
        sh.Range("H" & iRow).value = .List(lstRow, 7)
        sh.Range("I" & iRow).value = .List(lstRow, 8)
        sh.Range("J" & iRow).value = .List(lstRow, 9)
        sh.Range("K" & iRow).value = .List(lstRow, 10)
        sh.Range("L" & iRow).value = .List(lstRow, 11)
        sh.Range("M" & iRow).value = .List(lstRow, 12)
        sh.Range("N" & iRow).value = .List(lstRow, 13)

        sh.Range("P" & iRow).value = DateValue(Order_Cancellation.DateComboBox.value)
        sh.Range("Q" & iRow).value = Order_Cancellation.txtEmpName.value
        sh.Range("R" & iRow).value = Order_Cancellation.cmbCancellationReason.value
        sh.Range("S" & iRow).value = Order_Cancellation.txtNotes.value
        sh.Range("T" & iRow).value = Now
    End With

    ' Delete the selected row from the "Completed" worksheet.
    ThisWorkbook.Sheets("Completed").Rows(lstRow + 2).EntireRow.Delete

    ' Call the ListBoxReference subroutine to update the list box.
    Call ListBoxReference

    ' Record the current date and time in the "Support_Data" worksheet.
    Dim currentDateTime As Date
    currentDateTime = Now
    ws.Cells(2, 9).value = txtEmpName.value
    ws.Cells(2, 10).value = currentDateTime

    ' Call the LastEntry subroutine.
    Call LastEntry

    ' Unload the UserForm.
    Unload Me

    ' Display a success message
    MsgBox "Entry has been successful!", vbInformation + vbOKOnly, "Success"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Miscellaneous '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub info3_Click()
    ' Display the user manual.
    User_Manual.Show
End Sub


