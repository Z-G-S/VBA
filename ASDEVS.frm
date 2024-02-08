VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ASDEVS 
   Caption         =   "Sales Data Entry"
   ClientHeight    =   10486
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   18600
   OleObjectBlob   =   "ASDEVS.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ASDEVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Automated Sales Data Entry and Visualisation System (ASDEVS) Userform
' This module is part of a comprehensive user interface designed for managing and interacting with sales data in and between
' The Excel Sheets.

' Besides the code directly below, the rest of the module includes subroutines and functions that enable:
' 1. Data_Entry and Editing: Handling user inputs for sales data and performing validations.
' 2. List Management: Populating, managing, and interacting with lists of sales data.
' 3. Visualization and Reporting: Generating and displaying various charts and reports.
' 4. Form Control and Navigation: Managing the form's appearance and user navigation.
' 5. Data Export: Enabling the export of data to external formats like xlsm.
' 6. User Feedback and Interaction: Providing message boxes for confirmations and errors.
' 7. Utility Functions: Including utilities for counting rows, validating entries, and setting view settings.

' Declaring a custom type 'ValidationInfo' to store information for validating ComboBox inputs.
Private Type ValidationInfo
    cmb As ComboBox              ' ComboBox control to be validated.
    acceptableRange As Range     ' Range of acceptable values for the ComboBox.
    errorMessage As String       ' Error message to display if validation fails.
End Type

' Declare module-level variables to keep track of the last found index in search functionalities.
Private lastFoundIndex As Long         ' Stores the last found index for search results in a completed list.
Private lastFoundIndexCancelled As Long ' Stores the last found index for search results in a cancelled list.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Initialisation and Setup '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub userform_initialize()
    ' This subroutine is called when the user form is initialized.
    ' It sets up the form with necessary data and configurations.

    ' Call subroutines to set the initial view and populate list boxes.
    Call DataEntry_Click
    Call ListBoxReference
    Call RowCount_Completed
    Call RowCount_Cancelled
    Call LastEntry

    ' Populate the DateComboBox with dates.
    PopulateDateComboBox Me.DateComboBox
    
    ' Automatically fill in the employee's name from the system environment.
    txtEmpName.value = Environ("username")
    
    ' Define and set the RowSource for the 'Country' ComboBox.
    ' It uses data from the 'Support_Data' sheet.
    Sheets("Support_Data").Range("A2", Sheets("Support_Data").Range("A" & Application.Rows.Count).End(xlUp)).Name = "Country"
    cmbCountry.RowSource = "Country"
    cmbCountry.value = ""
    
    ' Define and set the RowSource for the 'ProductName' ComboBox.
    ' It uses data from the 'Support_Data' sheet.
    Sheets("Support_Data").Range("C2", Sheets("Support_Data").Range("C" & Application.Rows.Count).End(xlUp)).Name = "ProductName"
    cmbProductName.RowSource = "ProductName"
    cmbProductName.value = ""
    
    ' Define and set the RowSource for the 'PaymentType' ComboBox.
    ' It uses data from the 'Support_Data' sheet.
    Sheets("Support_Data").Range("E2", Sheets("Support_Data").Range("E" & Application.Rows.Count).End(xlUp)).Name = "PaymentType"
    cmbPaymentType.RowSource = "PaymentType"
    cmbPaymentType.value = ""
    
    ' Initialize various form controls to their default values.
    txtCustName.value = ""
    txtEmail.value = ""
    checkEngraved.value = False
    txtDiscount.value = ""
    txtRowNumber = ""
    
End Sub

Private Sub ListBoxPages_Change()
    ' This subroutine ensures that when the user changes from viewing the Completed list box to the
    ' Cancelled list box, their selection is deselected as without this, if the user decides to
    ' Delete, Edit or Cancel an entry, the system can get confused. Therefore, this code avoids this.
    
    ' Check if ListBoxPages value is 0 or 1 (Completed list box or Cancelled list box).
    If ListBoxPages.value = 0 Or ListBoxPages.value = 1 Then
        ' Deselect all items in lstCompleted.
        Dim i As Integer
        For i = 0 To lstCompleted.ListCount - 1
            lstCompleted.Selected(i) = False
        Next i
        
        ' Deselect all items in lstCancelled.
        Dim ii As Integer
        For ii = 0 To lstCancelled.ListCount - 1
            lstCancelled.Selected(ii) = False
        Next ii
    End If
End Sub

Sub SetNormalViewZoom()
    ' This subroutine is used to set the zoom level of the active worksheet to a specific value of 74%.
    ' This value has been chosen through testing how charts render in the user form image boxes at different
    ' Zoom levels and 74% renders all charts correctly.
    
    ' Declare a worksheet variable and set it to the "Data_Entry" sheet of the workbook.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_Entry")
    
    ' Set the zoom level of the worksheet's parent window to 74%.
    ws.Parent.Windows(1).Zoom = 74
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Form Control and Navigation '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub DataEntry_Click()
    ' This subroutine adjusts the user form's interface to reflect that the 'DataEntry' section is active.

    MultiPage1.value = 0

    ' Call the 'ResetButtonColours' subroutine to reset the colours of all navigation buttons.
    ' This ensures that all buttons are in their default state before highlighting the active one.
    Call ResetButtonColours

    ' Change the background color of the 'DataEntry' button to indicate that it is the active section.
    ' The colour code '&H8000000D' is used here for the active state.
    DataEntry.BackColor = &H8000000D
End Sub

Private Sub Visualisations_Click()
    ' This subroutine is triggered when the 'Visualisations' button is clicked on the user form.
    ' It handles the display and update of various charts and financial metrics.

    ' Call subroutine to set the optimal zoom for the Excel Sheets to capture an image of the charts.
    Call SetNormalViewZoom

    ' Call subroutine to render the charts on the form from the Excel Sheets.
    Call RenderCharts

    RevenuePieChart.value = True

    ' Reset the background colors of all navigation buttons and highlight the 'Visualisations' button.
    Call ResetButtonColours
    Visualisations.BackColor = &H8000000D  ' The color code '&H8000000D' is used to indicate active state.

    ' Define arrays to store chart names and corresponding image control names.
    Dim chartNames(1 To 4) As String
    chartNames(1) = "Stacked Column 1"
    chartNames(2) = "Stacked Column 2"
    chartNames(3) = "PieChart Sales"
    chartNames(4) = "PieChart Revenue"

    Dim imageControls(1 To 4) As String
    imageControls(1) = "ChartShow2"
    imageControls(2) = "ChartShow3"
    imageControls(3) = "ChartShow4"
    imageControls(4) = "ChartShow5"

    ' Iterate through each chart and corresponding image control to update the display.
    Dim i As Integer
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = "ASDEVS" Then
            For i = 1 To 4
                Dim imgControl As Object
                Set imgControl = uf.Controls(imageControls(i))

                Dim chartObj As ChartObject
                Set chartObj = ThisWorkbook.Sheets("Data_Entry").ChartObjects(chartNames(i))

                ' Export the chart as an image file and load it into the image control.
                Dim tempFile As String
                tempFile = Environ$("TEMP") & "\ChartImage" & i & ".jpg"
                chartObj.chart.Export tempFile
                imgControl.Picture = LoadPicture(tempFile)
                Kill tempFile  ' Delete the temporary file.
            Next i
            Exit For  ' Exit the loop once the form is found and updated.
        End If
    Next uf

    ' Update financial metrics from the 'Calculations' worksheet.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Calculations")
    AllTimeRevenue.value = Format(ws.Range("A18").value, "£#,##0.00")
    AllTimeProfit.value = Format(ws.Range("D18").value, "£#,##0.00")
    AllTimeCost.value = Format(ws.Range("B18").value, "£#,##0.00")
    ReturnPercentage.value = Format((AllTimeProfit / AllTimeRevenue), "0.00%")

    ' Set the MultiPage control to show the visualisations page.
    MultiPage1.value = 1
End Sub

Private Sub WorldSalesMap_Click()
    ' This subroutine is executed when the 'WorldSalesMap' button is clicked on the user form.
    ' It updates the user interface to display the world sales map.

    ' Call the 'RenderCharts' subroutine to render charts from the Excel Sheets on the user form.
    Call RenderCharts

    MultiPage1.value = 2

    ' Reset the background colours of navigation buttons to their default state.
    Call ResetButtonColours

    ' Change the background colour of the 'WorldSalesMap' button to indicate it is currently active.
    WorldSalesMap.BackColor = &H8000000D  ' The colour code '&H8000000D' is used to highlight the active state.

    ' Define a Worksheet variable to refer to the 'Data_Entry' sheet.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_Entry")
    
    ' Define a ChartObject variable to refer to the chart named 'ChartMap' on the 'Data_Entry' sheet.
    Dim chart1 As ChartObject
    Set chart1 = ws.ChartObjects("ChartMap")
    
    ' Define a string variable to store the path where the chart image will be temporarily saved.
    Dim tempPath1 As String
    tempPath1 = Environ$("TEMP") & "\ChartImage.jpg"
    
    ' Export the chart as an image to the temporary path.
    chart1.chart.Export tempPath1

    ' Load the exported chart image into a picture control ('ChartShow1') on the user form.
    ChartShow1.Picture = LoadPicture(tempPath1)
    
    ' Delete the temporary chart image file to clean up and free resources.
    Kill tempPath1
End Sub

Private Sub UserManual_Click()
    ' Display the user manual when the UserManual button is clicked.

    ' Reset the background colour of all buttons to their default state.
    Call ResetButtonColours

    ' Set the background colour of the 'UserManual' button a specific colour indicating that it is active.
    UserManual.BackColor = &H8000000D

    ' Show the user manual.
    User_Manual.Show
End Sub


Private Sub ResetButtonColours()
    ' This subroutine resets the background colours of specific buttons on the user form.

    ' Declare a collection to hold the button controls.
    Dim Inputs As Collection
    Set Inputs = New Collection

    ' Declare an array to store the buttons that need their colours reset.
    Dim InputBoxes() As Variant

    ' Declare a variant to iterate through the array of buttons.
    Dim var As Variant

    ' Initialize the array with the button controls that need their colours reset.
    InputBoxes = Array(DataEntry, Visualisations, WorldSalesMap, UserManual)
    
    ' Loop through each button control in the array.
    For Each var In InputBoxes
        ' Reset the background colour of the button.
        var.BackColor = &H3A1F1A  ' &H3A1F1A represents a specific colour code.
    Next
End Sub


Private Sub CloseButton_Click()
    ' This subroutine closes the user form if the user response equals "Yes".

    ' Declare a variable to capture the user's response to the confirmation message box.
    Dim i As VbMsgBoxResult
    
    ' Display a message box asking the user if they want to close the user form or not.
    i = MsgBox("Do you want to close the Entry Form?", vbYesNo + vbQuestion, "Close Entry Form")
    
    ' Check the user's response.
    If i = vbNo Then
        Exit Sub
    End If

' If the user selects 'Yes', unload the form.
Unload Me
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Data Entry, Editing, and Updates '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdSubmit_Click()
    ' This subroutine handles the submission of data.
    ' It includes validations for input fields, ComboBoxes, and discount calculations.

    ' Confirm submission with the user.
    Dim i As VbMsgBoxResult
    i = MsgBox("Do you want to submit the data?", vbYesNo + vbQuestion, "Submit")
    If i = vbNo Then Exit Sub  ' Exit if the user chooses not to proceed.

    ' Initialize input validation.
    Dim Inputs As Collection
    Set Inputs = New Collection
    Dim InputBoxes() As Variant
    Dim var As Variant
    Dim isError1 As Boolean  ' Flag for error in inputs.

    ' Array of input controls to check for empty values.
    InputBoxes = Array(DateComboBox, txtEmpName, txtCustName, txtEmail, cmbCountry, cmbProductName, txtDiscount, cmbPaymentType)

    ' Validate each input control in the array.
    For Each var In InputBoxes
        If var.value = "" Then
            var.BackColor = vbRed  ' Highlight empty fields in red.
            isError1 = True
        Else
            var.BackColor = vbWhite  ' Reset color if field is filled.
        End If
    Next var

    ' Show error message and exit if any field is empty.
    If isError1 Then
        MsgBox "Please enter an input for the boxes that are highlighted red.", vbCritical + vbOKOnly, "No Input Error 1001"
        Exit Sub
    End If

    ' Validate ComboBox data against ranges in Support_Data sheet.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Support_Data")
    
    Dim validations() As ValidationInfo
    ReDim validations(1 To 3)  ' Resize array for three ComboBox validations.

    ' Set up validations for Country, ProductName, and PaymentType ComboBoxes.
    Set validations(1).cmb = cmbCountry
    Set validations(1).acceptableRange = ws.Range("A2:A197")
    validations(1).errorMessage = "Please ensure that the Country inputted matches the countries from the list."
    
    Set validations(2).cmb = cmbProductName
    Set validations(2).acceptableRange = ws.Range("C2:C9")
    validations(2).errorMessage = "Please ensure that the Product Name inputted matches the products from the list."
    
    Set validations(3).cmb = cmbPaymentType
    Set validations(3).acceptableRange = ws.Range("E2:E13")
    validations(3).errorMessage = "Please ensure that the Payment Type inputted matches the types from the list."

    ' Perform the ComboBox validation checks.
    Dim ii As Integer
    Dim hasValidationError As Boolean
    For ii = LBound(validations) To UBound(validations)
        ValidateComboBox validations(ii)
        If validations(ii).cmb.BackColor = RGB(255, 0, 0) Then
            hasValidationError = True
            Exit For  ' Exit loop if validation error is found.
        End If
    Next ii

    ' Exit if any ComboBox validation fails.
    If hasValidationError Then Exit Sub

    ' Validate Discount field for numerical and range correctness.
    Dim discount As Double
    If Not IsNumeric(txtDiscount.value) Then
        txtDiscount.BackColor = vbRed
        MsgBox "Please enter a numerical discount value.", vbCritical, "Non-Numerical Value Error 2001"
        Exit Sub
    ElseIf txtDiscount.value < 0 Then
        txtDiscount.BackColor = vbRed
        MsgBox "Please ensure the discount value is greater than or equal to 0", vbCritical, "Negative Discount Value Error 2002"
        Exit Sub
    ElseIf txtDiscount.value > 100 Then
        txtDiscount.BackColor = vbRed
        MsgBox "Please ensure the discount value is less than or equal to 100", vbCritical, "Discount Value Too Large Error 2003"
        Exit Sub
    Else
        txtDiscount.BackColor = vbWhite
    End If

    ' Convert discount to a decimal value if not zero.
    If txtDiscount.value <> 0 Then
        discount = CDbl(txtDiscount.value) / 100
    Else
        discount = 0
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

    ' Determine the row for inserting or updating data.
    Dim iRow As Long
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Completed")

    ' Retrieve price and cost based on selected product and engraving status.
    Dim productInfo As Variant
    productInfo = GetPriceAndCost(cmbProductName.value, checkEngraved.value, ThisWorkbook.Sheets("Prices_Costs_Data"))

    ' Calculate financial details from input data.
    Dim OriginalPrice As Double, Price As Double, DiscountValue As Double, Revenue As Double, Cost As Double, Profit As Double
    OriginalPrice = productInfo(0)
    Price = OriginalPrice
    If txtRowNumber.value = "" Then
        iRow = sh.Range("A" & sh.Rows.Count).End(xlUp).Row + 1
    Else
        iRow = txtRowNumber.value
    End If

    ' Update the worksheet with the new data.
    With sh
        .Range("A" & iRow).value = DateValue(DateComboBox.value)
        .Range("B" & iRow).value = txtEmpName.value
        .Range("C" & iRow).value = txtCustName.value
        .Range("D" & iRow).value = txtEmail.value
        .Range("E" & iRow).value = cmbCountry.value
        .Range("F" & iRow).value = cmbProductName.value
        .Range("G" & iRow).value = IIf(checkEngraved.value, "Yes", "No")
        .Range("H" & iRow).value = discount
        .Range("I" & iRow).value = cmbPaymentType.value
        .Range("J" & iRow).value = Price
        .Range("K" & iRow).value = DiscountValue
        .Range("L" & iRow).value = Revenue
        .Range("M" & iRow).value = Cost
        .Range("N" & iRow).value = Profit
        .Range("O" & iRow).value = iRow - 1
        .Range("P" & iRow).value = Now
    End With
    
    ' Log the submission date and time.
    Dim currentDateTime As Date
    currentDateTime = Now
    ws.Cells(2, 9).value = txtEmpName.value
    ws.Cells(2, 10).value = currentDateTime

    ' Update list boxes, reinitialize form and update counters.
    Call ListBoxReference
    Call RowCount_Completed
    Call RowCount_Cancelled
    Call userform_initialize

    ' Load the last entry into the form.
    Call LastEntry

    ' Notify user of successful submission.
    MsgBox "Entry has been successful!", vbInformation + vbOKOnly, "Success"
End Sub

Private Sub cmdEdit_Click()
    ' This subroutine handles the event when the 'Edit' button is clicked.

    ' Check if no item is selected in either list.
    If Selected_List = 0 Then
        ' If no item is selected, inform the user and exit the subroutine.
        MsgBox "No row is selected.", vbOKOnly + vbInformation, "Edit"
        Exit Sub
    End If

    ' Determine which list has a selected item for editing.
    If ASDEVS.lstCompleted.ListIndex <> -1 Then
        ' If an item in the 'Completed' list is selected, go to the editing code for 'Completed'.
        GoTo listCompletedCode
    ElseIf ASDEVS.lstCancelled.ListIndex <> -1 Then
        ' If an item in the 'Cancelled' list is selected, go to the editing code for 'Cancelled'.
        GoTo lstCancelledCode
    End If

    ' Code block for editing an entry from the 'Completed' list.
listCompletedCode:
    Dim tempDateStore As Double
    With ASDEVS
        ' Load the selected entry's data into the form fields for editing.
        .txtRowNumber.value = Me.lstCompleted.ListIndex + 2
        tempDateStore = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 0)
        .DateComboBox.value = Format(tempDateStore, "dd/mm/yyyy")
        .txtEmpName.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 1)
        .txtCustName.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 2)
        .txtEmail.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 3)
        .cmbCountry.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 4)
        .cmbProductName.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 5)
        ' Check if the product is engraved and set the checkbox accordingly.
        If Me.lstCompleted.List(Me.lstCompleted.ListIndex, 6) = "Yes" Then
            .checkEngraved.value = True
        ElseIf Me.lstCompleted.List(Me.lstCompleted.ListIndex, 6) = "No" Then
            .checkEngraved.value = False
        End If
        .txtDiscount.value = (Me.lstCompleted.List(Me.lstCompleted.ListIndex, 7)) * 100
        .cmbPaymentType.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 8)
    End With
    ' Inform the user that the entry is ready for editing.
    MsgBox "The selected entry has been loaded into the form, ready to be edited. Simply make the necessary changes and click Submit.", vbInformation + vbOKOnly, "Edit Entry"
    Exit Sub
    
    ' Code block for handling attempts to edit an entry from the 'Cancelled' list.
lstCancelledCode:
    ' Inform the user that cancelled entries cannot be edited.
    MsgBox "Cancelled entries cannot be edited.", vbInformation + vbOKOnly, "Error"
    
End Sub

Private Sub cmdDelete_Click()
    ' This subroutine handles the deletion of a selected entry from the lists in the user form.

    ' Check if no item is selected in the list.
    If Selected_List = 0 Then
        ' If no item is selected, display a message box and exit the subroutine.
        MsgBox "No row is selected.", vbOKOnly + vbInformation, "Delete"
        Exit Sub
    End If
    
    ' Declare a variable to capture the user's response to the confirmation message box.
    Dim response As VbMsgBoxResult
    ' Display a confirmation message box to confirm deletion.
    response = MsgBox("Are you sure you want to delete the selected entry? This action cannot be undone.", vbYesNo + vbQuestion, "Confirm Delete")

    ' Check the user's response.
    If response = vbNo Then
        ' If the user selects 'No', exit the subroutine without performing the deletion.
        Exit Sub
    End If
    
    ' Declare variables for handling the deletion process.
    Dim lstRow As Long
    Dim currentDateTime As Date
    currentDateTime = Now  ' Store the current date and time.
        
    ' Determine which list ('Completed' or 'Cancelled') the deletion is to be performed on.
    If ASDEVS.lstCompleted.ListIndex <> -1 Then
        lstRow = lstCompleted.ListIndex
        GoTo listCompletedCode  ' Jump to the code section for handling 'Completed' list deletion.
    ElseIf ASDEVS.lstCancelled.ListIndex <> -1 Then
        lstRow = lstCancelled.ListIndex
        GoTo lstCancelledCode   ' Jump to the code section for handling 'Cancelled' list deletion.
    End If
    
    ' Code section for handling deletion from the 'Completed' list.
listCompletedCode:
    ThisWorkbook.Sheets("Completed").Rows(lstRow + 2).EntireRow.Delete  ' Delete the selected row.
    MsgBox "Selected entry deleted.", vbInformation + vbOKOnly, "Delete Entry"
    Call ListBoxReference  ' Refresh the list box reference.
    Call RowCount_Completed  ' Update the row count for 'Completed'.
    ' Update the 'Support_Data' sheet with the current user's name and date/time of deletion.
    ThisWorkbook.Worksheets("Support_Data").Cells(2, 9).value = txtEmpName.value
    ThisWorkbook.Worksheets("Support_Data").Cells(2, 10).value = currentDateTime
    Call LastEntry  ' Load the last entry into the form.
    Exit Sub

    ' Code section for handling deletion from the 'Cancelled' list.
lstCancelledCode:
    ThisWorkbook.Sheets("Cancelled").Rows(lstRow + 2).EntireRow.Delete  ' Delete the selected row.
    MsgBox "Selected entry deleted.", vbInformation + vbOKOnly, "Delete Entry"
    Call ListBoxReference  ' Refresh the list box reference.
    Call RowCount_Cancelled  ' Update the row count for 'Cancelled'.
    ' Update the 'Support_Data' sheet with the current user's name and date/time of deletion.
    ThisWorkbook.Worksheets("Support_Data").Cells(2, 9).value = txtEmpName.value
    ThisWorkbook.Worksheets("Support_Data").Cells(2, 10).value = currentDateTime
    Call LastEntry  ' Load the last entry into the form.
    Exit Sub

End Sub

Private Sub cmdCancel_Click()
    ' This subroutine is executed when the 'Cancel' button is clicked on the user form.

    ' Check if no item is selected in either the 'Completed' or 'Cancelled' list.
    If Selected_List = 0 Then
        ' If no item is selected, inform the user and exit the subroutine.
        MsgBox "No row is selected.", vbOKOnly + vbInformation, "Cancel"
        Exit Sub
    End If
    
    ' Determine which list has a selected item for cancellation.
    If ASDEVS.lstCompleted.ListIndex <> -1 Then
        ' If an item in the 'Completed' list is selected, proceed with cancellation for 'Completed'.
        GoTo listCompletedCode
    ElseIf ASDEVS.lstCancelled.ListIndex <> -1 Then
        ' If an item in the 'Cancelled' list is selected, display a message that cancellation is not possible.
        GoTo lstCancelledCode
    End If
    
    ' Code block for handling cancellation of an entry from the 'Completed' list.
listCompletedCode:
    ' Load and display the 'Order_Cancellation' form for further actions.
    Load Order_Cancellation
    With Order_Cancellation
        ' Pre-fill the customer name and email fields with the selected entry's data, and disable editing.
        .txtCustName.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 2)
        .txtCustName.Enabled = False
        .txtEmail.value = Me.lstCompleted.List(Me.lstCompleted.ListIndex, 3)
        .txtEmail.Enabled = False
        .txtNotes.value = ""  ' Initialize the notes field.
        .Show  ' Display the 'Order_Cancellation' form.
    End With
    
    ' Update the row counts for both 'Completed' and 'Cancelled' lists.
    Call RowCount_Completed
    Call RowCount_Cancelled
    
    ' Exit the subroutine after processing the cancellation for 'Completed'.
    Exit Sub
    
    ' Code block for handling attempts to cancel an entry from the 'Cancelled' list.
lstCancelledCode:
    ' Inform the user that entries in the 'Cancelled' list cannot be cancelled.
    MsgBox "Cancelled entries cannot be cancelled.", vbInformation + vbOKOnly, "Information"
    
End Sub

Private Sub ResetButton_Click()
    ' This subroutine handles the event when the 'Reset' button is clicked on the user form.
    ' It prompts the user to confirm the reset action and then clears all input fields and updates counters.

    ' Declare a variable to capture the user's response to the confirmation message box.
    Dim i As VbMsgBoxResult

    ' Display a message box asking the user to confirm the reset of the entry form.
    i = MsgBox("Do you want to reset the Entry Form?", vbYesNo + vbQuestion, "Reset Entry Form")

    ' Check the user's response.
    If i = vbNo Then
        ' If the user selects 'No', exit the subroutine without performing the reset.
        Exit Sub
    End If

    ' Declare a variable to iterate through all controls on the form.
    Dim ctrl As Control

    ' Loop through each control on the form.
    For Each ctrl In Me.Controls
        ' Check if the control is a TextBox, ComboBox, or CheckBox.
        If TypeName(ctrl) = "TextBox" Or TypeName(ctrl) = "ComboBox" Or TypeName(ctrl) = "CheckBox" Then
            ' Reset the value of the control to its default (usually an empty string or False for CheckBoxes).
            ctrl.value = ""
        End If
    Next ctrl
    
    OptionCompleted.value = False
    OptionCancelled.value = False

    ' Call subroutines to update counters on the form.
    Call RowCount_Completed
    Call RowCount_Cancelled
    Call LastEntry
End Sub

Sub RowCount_Completed()
    ' This subroutine calculates the number of rows in the 'Completed' sheet and updates a text box on the user form.

    ' Declare a Worksheet variable to refer to the 'Completed' sheet.
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Completed")

    ' Calculate the number of rows in the 'Completed' sheet that are populated with data.
    ' It finds the last used row in column A and subtracts 1 to exclude the header.
    Dim rowCount As Long
    rowCount = sh.Cells(sh.Rows.Count, "A").End(xlUp).Row - 1

    ' Update the text box 'txtRowCount1' on the user form with the row count.
    Me.txtRowCount1.value = rowCount
End Sub

Sub RowCount_Cancelled()
    ' This subroutine calculates the number of rows in the 'Cancelled' sheet and updates a text box on the user form.

    ' Declare a Worksheet variable to refer to the 'Cancelled' sheet.
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets("Cancelled")

    ' Calculate the number of rows in the 'Cancelled' sheet that are populated with data.
    ' It finds the last used row in column A and subtracts 1 to exclude the header row.
    Dim rowCount As Long
    rowCount = sh.Cells(sh.Rows.Count, "A").End(xlUp).Row - 1

    ' Update the text box 'txtRowCount2' on the user form with the row count.
    Me.txtRowCount2.value = rowCount
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Search Functionality '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdSearchID1_Click()
    ' This subroutine performs a case-insensitive search for the entered Customer ID within the list box
    ' And highlights the matching row if found.
    ' If the Customer ID is not found, it displays a message to inform the user.
    
    ' Check if the Customer ID input is empty.
    If txtSearchID1.value = "" Then
        MsgBox "Please enter a Customer ID before searching.", vbOKOnly + vbExclamation, "No Input Error 1001"
        Exit Sub
    End If

    ' Reset lastFoundIndex when a new search is initiated.
    lastFoundIndex = -1
    
    Dim searchValue As String
    Dim i As Long
    
    ' Get the search value from the text box and convert it to uppercase for case-insensitive search.
    searchValue = UCase(txtSearchID1.value)
    
    ' Loop through list box items and highlight only the first matching one (case-insensitive).
    For i = 0 To lstCompleted.ListCount - 1
        If UCase(lstCompleted.List(i, 2)) = searchValue Then
            ' If a match is found, highlight the entire row in the list box.
            lstCompleted.Selected(i) = True
            lastFoundIndex = i
            Exit For ' Exit the loop after the first match is found.
        End If
    Next i
    
    ' Display a message based on the search result.
    If lstCompleted.Selected(i) Then
        MsgBox "Record found.", vbOKOnly, "Successful Search"
    Else
        MsgBox "No records found.", vbOKOnly + vbCritical, "Unsuccessful Search Error 3001"
    End If
End Sub

Private Sub cmdNextMatch1_Click()
    ' This subroutine extends the cmdSearchID1_Click subroutine, by finding the next match in the search.
    ' It performs a case-insensitive search for the same Customer ID that was previously found and highlights
    ' The next matching row if found.
    ' If there are no more matches, it displays a message to inform the user.
    
    Dim searchValue As String
    Dim i As Long
    
    ' Get the search value from the text box and convert it to uppercase for case-insensitive search.
    searchValue = UCase(txtSearchID1.value)
    
    ' Continue searching from the index of the last found match.
    For i = lastFoundIndex + 1 To lstCompleted.ListCount - 1
        If UCase(lstCompleted.List(i, 2)) = searchValue Then
            ' If a match is found, highlight the entire row in the list box.
            lstCompleted.Selected(i) = True
            lastFoundIndex = i
            Exit For ' Exit the loop after the next match is found.
        End If
    Next i
    
    ' Display a message based on the search result and if there are no more matches.
    If lstCompleted.Selected(i) Then
        MsgBox "Next match found.", vbOKOnly, "Successful Match"
    Else
        MsgBox "No more matches found.", vbOKOnly + vbCritical, "Unsuccessful Match Error 3002"
    End If
End Sub

Private Sub cmdSearchID2_Click()
    ' This subroutine follows the same logic as cmdSearchID1_Click() as it provides the
    ' Same functionality but for the Cancelled list box.

    ' Check if the input text box is empty.
    If txtSearchID2.value = "" Then
        MsgBox "Please enter a Customer ID before searching.", vbOKOnly + vbExclamation, "No Input Error 1001"
        Exit Sub
    End If

    ' Reset lastFoundIndexCancelled when a new search is initiated.
    lastFoundIndexCancelled = -1
    
    Dim searchValue As String
    Dim i As Long
    
    ' Get the search value from the text box and convert it to uppercase for case-insensitive search.
    searchValue = UCase(txtSearchID2.value)
    
    ' Loop through list box items and highlight only the first matching one (case-insensitive).
    For i = 0 To lstCancelled.ListCount - 1
        If UCase(lstCancelled.List(i, 2)) = searchValue Then
            ' If a match is found, highlight the entire row in the list box.
            lstCancelled.Selected(i) = True
            lastFoundIndexCancelled = i
            Exit For ' Exit the loop after the first match is found.
        End If
    Next i
    
    ' Display a message based on the search result.
    If lstCancelled.Selected(i) Then
        MsgBox "Record found.", vbOKOnly, "Successful Search"
    Else
        MsgBox "No records found.", vbOKOnly + vbCritical, "Unsuccessful Search Error 3001"
    End If
End Sub

Private Sub cmdNextMatch2_Click()
    ' This subroutine follows the same logic as cmdNextMatch2_Click() as it provides the
    ' Same functionality but for the Cancelled list box.
    
    Dim searchValue As String
    Dim i As Long
    
    ' Get the search value from the text box.
    searchValue = UCase(txtSearchID2.value)
    
    ' Continue searching from the index after the last found match.
    For i = lastFoundIndexCancelled + 1 To lstCancelled.ListCount - 1
        If UCase(lstCancelled.List(i, 2)) = searchValue Then
            ' If a match is found, highlight the entire row in the list box.
            lstCancelled.Selected(i) = True
            lastFoundIndexCancelled = i
            Exit For ' Exit the loop after the next match is found.
        End If
    Next i
    
    ' Display a message based on the search result.
    If lstCancelled.Selected(i) Then
        MsgBox "Next match found.", vbOKOnly, "Successful Match"
    Else
        MsgBox "No more matches found.", vbOKOnly + vbCritical, "Unsuccessful Match Error 3002"
    End If
End Sub

Private Sub cmdNewtoOld1_Click()
    ' This subroutine sorts the data in columns A to P in the "Completed" sheet
    ' Based on the dates in column A. It sorts the dates from newest to oldest.
    
' Declare a variable to represent the worksheet.
    Dim ws As Worksheet
    
    ' Set the ws variable to the "Completed" sheet in the workbook.
    Set ws = ThisWorkbook.Sheets("Completed")

    ' Using the Sort object of the worksheet.
    With ws.Sort
        ' Clear any previous sort fields to ensure a fresh sort.
        .SortFields.Clear

        ' Add a new sort field - the range is column A from row 1 to the last row with data.
        ' Sorting is based on values, in ascending order (newest to oldest dates).
        .SortFields.Add Key:=ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        ' Define the range to sort, which is columns A to P up to the last row with data in column A.
        .SetRange ws.Range("A1:P" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

        ' Specify that the first row contains headers (change to xlNo if there are no headers).
        .Header = xlYes

        ' Apply the sort to the range.
        .Apply
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Sorting Data '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdOldtoNew1_Click()
    ' This subroutine sorts the data in columns A to P in the "Completed" sheet
    ' Based on the dates in column A. It sorts the dates from oldest to newest.
    
    ' Declare a variable to represent the worksheet.
    Dim ws As Worksheet
    
    ' Set the ws variable to the "Completed" sheet in the workbook.
    Set ws = ThisWorkbook.Sheets("Completed")

    ' Using the Sort object of the worksheet.
    With ws.Sort
        ' Clear any previous sort fields to ensure a fresh sort.
        .SortFields.Clear

        ' Add a new sort field - the range is column A from row 1 to the last row with data.
        ' Sorting is based on values, in ascending order (oldest to newest dates).
        .SortFields.Add Key:=ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        ' Define the range to sort, which is columns A to P up to the last row with data in column A.
        .SetRange ws.Range("A1:P" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

        ' Specify that the first row contains headers (change to xlNo if there are no headers).
        .Header = xlYes

        ' Apply the sort to the range.
        .Apply
    End With
End Sub

Private Sub cmdNewtoOld2_Click()
    ' Same logic as cmdNewtoOld1, applied to the Cancelled sheet.
    
' Declare a variable to represent the worksheet.
    Dim ws As Worksheet
    
    ' Set the ws variable to the "Cancelled" sheet in the workbook.
    Set ws = ThisWorkbook.Sheets("Cancelled")

    ' Using the Sort object of the worksheet.
    With ws.Sort
        ' Clear any previous sort fields to ensure a fresh sort.
        .SortFields.Clear

        ' Add a new sort field - the range is column A from row 1 to the last row with data.
        ' Sorting is based on values, in ascending order (newest to oldest dates).
        .SortFields.Add Key:=ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        ' Define the range to sort, which is columns A to P up to the last row with data in column A.
        .SetRange ws.Range("A1:P" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

        ' Specify that the first row contains headers (change to xlNo if there are no headers).
        .Header = xlYes

        ' Apply the sort to the range.
        .Apply
    End With
End Sub

Private Sub cmdOldtoNew2_Click()
    ' Same logic as cmdNewtoOld2, applied to the Cancelled sheet
    
    ' Declare a variable to represent the worksheet.
    Dim ws As Worksheet
    
    ' Set the ws variable to the "Cancelled" sheet in the workbook.
    Set ws = ThisWorkbook.Sheets("Cancelled")

    ' Using the Sort object of the worksheet.
    With ws.Sort
        ' Clear any previous sort fields to ensure a fresh sort
        .SortFields.Clear

        ' Add a new sort field - the range is column A from row 1 to the last row with data.
        ' Sorting is based on values, in ascending order (oldest to newest dates).
        .SortFields.Add Key:=ws.Range("A1:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

        ' Define the range to sort, which is columns A to P up to the last row with data in column A.
        .SetRange ws.Range("A1:P" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

        ' Specify that the first row contains headers (change to xlNo if there are no headers).
        .Header = xlYes

        ' Apply the sort to the range.
        .Apply
    End With
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Visualisation and Reporting '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub RenderCharts()
    ' This subroutine calls upon the UpdateAndActivateChart subroutine to update the charts listed in the array
    ' Ensuring that all charts are rendered correctly, even if the user has not manually rendered the charts
    ' By visiting their location of the Excel Sheets, avoiding errors.
    
' Array of chart names
    Dim chartNames As Variant
    chartNames = Array("Stacked Column 1", "Stacked Column 2", "PieChart Sales", _
                       "PieChart Revenue", "PieChart Revenue", "PieChart Profit", _
                       "PieChart Cost", "PieChart Cost Sources", "ChartMap")

    ' Update and activate each chart
    Dim chartName As Variant
    For Each chartName In chartNames
        UpdateAndActivateChart "Data_Entry", chartName
    Next chartName
    
    End Sub

Private Sub UpdateAndActivateChart(ByVal sheetName As String, ByVal chartName As String)
    ' This subroutine activates the specified Excel sheet, activates the specified chart on that sheet,
    ' And refreshes the chart if it exists. It also allows Excel to complete pending tasks, including chart rendering.
    ' It takes the sheet name and chart name as parameters.

    ' Activate the sheet containing the chart
    ' Set the optimal level of zoom to capture an image of the charts from the Excel Sheet.
    Call SetNormalViewZoom
    ThisWorkbook.Sheets(sheetName).Activate

    ' Activate the chart
    On Error Resume Next
    ThisWorkbook.Sheets(sheetName).ChartObjects(chartName).Activate
    On Error GoTo 0

    ' Check if an active chart (chartName specified) exists on the sheet.
    ' If an active chart exists, it is refreshed.
    If Not ActiveChart Is Nothing Then
        ActiveChart.Refresh
    End If

    ' Allow Excel to complete pending tasks, including chart rendering.
    ' This line is used to ensure that any pending tasks, such as chart rendering, are completed.
    DoEvents
End Sub

Private Sub RevenuePieChart_Click()
    ' Code to capture and display the chart when it is selected through an option button.
    ' Set the optimal zoom level to capture an image of the chart on the Excel Sheet.
    Call SetNormalViewZoom

    ' Reference the 'Data_Entry' worksheet.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_Entry")

    ' Reference the 'PieChart Revenue' chart object.
    Dim chart1 As ChartObject
    Set chart1 = ws.ChartObjects("PieChart Revenue")

    ' Define a temporary path for exporting the chart image.
    Dim tempPath1 As String
    tempPath1 = Environ$("TEMP") & "\ChartImage.jpg"

    ' Export the chart as an image to the temporary path.
    chart1.chart.Export tempPath1

    ' Load the exported chart image into the 'ChartShow5' control.
    ChartShow5.Picture = LoadPicture(tempPath1)

    ' Delete the temporary file to free up space.
    Kill tempPath1
End Sub

Private Sub ProfitPieChart_Click()
    ' Same logic as 'RevenuePieChart_Click'.
    ' Handles the display of the 'PieChart Profit' chart.

    Call SetNormalViewZoom
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_Entry")

    Dim chart1 As ChartObject
    Set chart1 = ws.ChartObjects("PieChart Profit")

    Dim tempPath1 As String
    tempPath1 = Environ$("TEMP") & "\ChartImage.jpg"

    chart1.chart.Export tempPath1
    ChartShow5.Picture = LoadPicture(tempPath1)
    Kill tempPath1
End Sub

Private Sub CostPieChart_Click()
    ' Same logic as 'RevenuePieChart_Click'.
    ' Handles the display of the 'PieChart Cost' chart.

    Call SetNormalViewZoom
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_Entry")

    Dim chart1 As ChartObject
    Set chart1 = ws.ChartObjects("PieChart Cost")

    Dim tempPath1 As String
    tempPath1 = Environ$("TEMP") & "\ChartImage.jpg"

    chart1.chart.Export tempPath1
    ChartShow5.Picture = LoadPicture(tempPath1)
    Kill tempPath1
End Sub

Private Sub SourcePieChart_Click()
    ' Same logic as 'RevenuePieChart_Click'.
    ' Handles the display of the 'PieChart Cost Sources' chart.

    Call SetNormalViewZoom
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data_Entry")

    Dim chart1 As ChartObject
    Set chart1 = ws.ChartObjects("PieChart Cost Sources")

    Dim tempPath1 As String
    tempPath1 = Environ$("TEMP") & "\ChartImage.jpg"

    chart1.chart.Export tempPath1
    ChartShow5.Picture = LoadPicture(tempPath1)
    Kill tempPath1
End Sub

Private Sub RefreshChartsandData_Click()
    ' Update the charts
    Call Visualisations_Click
End Sub
Private Sub RefreshWorldMap_Click()
    ' Update the chart
    Call WorldSalesMap_Click
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Utility and Validation Functions '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ValidateComboBox(ByRef info As ValidationInfo)
    ' This subroutine validates the input of a ComboBox control based on specified criteria.

    ' Retrieve the entered value from the ComboBox.
    Dim enteredValue As String
    enteredValue = info.cmb.value
    
    ' Check if the acceptable range for validation is defined.
    If Not (info.acceptableRange Is Nothing) Then
        ' If the acceptable range is defined, check if the entered value exists in the range.
        If Application.WorksheetFunction.CountIf(info.acceptableRange, enteredValue) = 0 Then
            ' If the value is not found in the range, highlight the ComboBox in red and show an error message.
            info.cmb.BackColor = RGB(255, 0, 0)  ' Set background to red for error indication.
            MsgBox info.errorMessage, vbCritical + vbOKOnly, "Validation Error 5001"
            Exit Sub  ' Exit the subroutine as the validation failed.
        Else
            ' If the value is found in the range, reset the background color of the ComboBox.
            info.cmb.BackColor = RGB(255, 255, 255)  ' Reset background to white.
        End If
    Else
        ' If the acceptable range is not defined, check against specific criteria.
        If enteredValue <> "In Progress/Fulfilled" And enteredValue <> "Cancelled" Then
            ' If the value does not meet the criteria, highlight the ComboBox in red and show an error message.
            info.cmb.BackColor = RGB(255, 0, 0)  ' Set background to red for error indication.
            MsgBox info.errorMessage, vbCritical + vbOKOnly, "Validation Error 5001"
            Exit Sub  ' Exit the subroutine as the validation failed.
        Else
            ' If the value meets the criteria, reset the background color of the ComboBox.
            info.cmb.BackColor = RGB(255, 255, 255)  ' Reset background to white.
        End If
    End If
End Sub

Private Function GetPriceAndCost(productName As String, isEngraved As Boolean, pricesCostsSheet As Worksheet) As Variant
    ' This function retrieves the price and cost information for a selected product based on its name and engraving status.
    ' It takes the product name, engraving status, and the worksheet containing price and cost data as parameters.
    ' It returns an array containing the price and cost values.
    
    Dim productNameRange As Range
    Dim priceColumn As Integer
    Dim costColumn As Integer

    ' Define the range of product names to column A from row 2 to 9 on the Prices_Costs_Data sheet.
    Set productNameRange = pricesCostsSheet.Range("A2:A9")

    ' Find the row number where the selected product name matches.
    Dim rowIndex As Variant
    rowIndex = Application.Match(productName, productNameRange, 0)

    If Not IsError(rowIndex) Then
        ' Determine the appropriate columns based on engraving status.
        If isEngraved Then
            priceColumn = 8 ' Column H for price with engraving.
            costColumn = 10 ' Column J for cost with engraving.
        Else
            priceColumn = 7 ' Column G for price without engraving.
            costColumn = 9 ' Column I for cost without engraving.
        End If

        ' Get the price and cost from the Prices_Costs_Data sheet based on the row and columns.
        GetPriceAndCost = Array(pricesCostsSheet.Cells(rowIndex + 1, priceColumn).value, _
                                pricesCostsSheet.Cells(rowIndex + 1, costColumn).value)
    Else
        ' Handle error if the product name is not found.
        GetPriceAndCost = Array(0, 0) ' Return zeros for price and cost.
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Data Export '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdExport_Click()
    ' This subroutine is used to export data from either the "Completed" or "Cancelled" worksheet to a xlsm file.
    ' It allows the user to select which data to export based on the selected option button.
    
    Dim wsCompleted As Worksheet
    Dim wsCancelled As Worksheet
    Dim wb As Workbook
    Dim savePath As String
    Dim confirmation As VbMsgBoxResult
    
    ' Display an error if no data sheet is selected before clicking Export.
    If OptionCompleted = False And OptionCancelled = False Then
        MsgBox "Please select a data sheet before exporting.", vbOKOnly + vbCritical, "No Selection Error 4001"
        Exit Sub
    End If

    ' Get confirmation from the user.
    confirmation = MsgBox("Are you sure you want to export the data?", vbQuestion + vbYesNo, "Export Data Confirmation")

    ' Check user's response.
    If confirmation <> vbYes Then
        Exit Sub
    End If
        
    ' Set the references to the "Completed" and "Cancelled" sheets.
    Set wsCompleted = ThisWorkbook.Sheets("Completed")
    Set wsCancelled = ThisWorkbook.Sheets("Cancelled")

    ' Create a new workbook to copy data.
    Set wb = Workbooks.Add

    ' Export data based on the selected option button.
    If OptionCompleted.value Then
        ExportSheet wsCompleted, wb
    ElseIf OptionCancelled.value Then
        ExportSheet wsCancelled, wb
    Else
        MsgBox "Please select a sheet.", vbExclamation
        Exit Sub
    End If

   ' Save the new workbook as an .xlsm file with a specified path and filename and display message.
    savePath = Application.GetSaveAsFilename(FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm")
    If savePath <> "False" Then
        wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        wb.Close
        MsgBox "Data exported successfully!", vbInformation
    Else
        MsgBox "Export canceled.", vbExclamation
    End If
End Sub

Sub ExportSheet(srcSheet As Worksheet, destWorkbook As Workbook)
    ' This subroutine is used to copy data from the source worksheet to a new workbook.
    ' It takes the source worksheet and the destination workbook as parameters.
    
    ' Copy data from the original sheet to the new workbook.
    srcSheet.Copy Before:=destWorkbook.Sheets(1)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Miscellaneous '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdFullScreen_Click()
    ' This subroutine is triggered when the 'Full Screen' button is clicked.
    ' It warns the user about potential sizing and formatting issues when changing to full screen mode and
    ' back to restore mode, and proceeds to maximize or restore the user form based on the user's response.

    ' Declare a variable to store the user's response to the message box.
    Dim response As VbMsgBoxResult

    ' Display a message box warning about full screen mode issues and asking for confirmation to proceed.
    response = MsgBox("Please note that changing Full Screen mode can cause sizing and formatting issues on some screen sizes." & _
                      vbNewLine & vbNewLine & "Do you want to proceed?", vbYesNo + vbInformation, "Full Screen/Restore")

    ' Check the user's response.
    If response = vbNo Then
        ' If the user clicks 'No', exit the subroutine without changing the screen size.
        Exit Sub
    End If
    
    ' If the user clicks 'Yes', call the 'Maximize_Restore' subroutine to toggle the screen size.
    Call Maximize_Restore
End Sub

Private Sub info1_Click()
    'Display the user manual
    User_Manual.Show
End Sub

Private Sub info2_Click()
    'Display the user manual
    User_Manual.Show
End Sub








  

   















