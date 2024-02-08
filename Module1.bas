Attribute VB_Name = "Module1"
Option Explicit
' This module contains subroutines and functions that are used across various user forms.
' It includes procedures for setting up list boxes managing form display, and populating combo boxes with dates.

Public iWidth As Integer
Public iHeight As Integer
Public iLeft As Integer
Public iTop As Integer
Public bState As Boolean

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' List Box Parameters '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ListBoxReference()
    ' This subroutine configures the list boxes in the 'ASDEVS' for displaying 'Completed' and 'Cancelled' data.
    ' It sets up the source data range and column settings for each list box.

    Dim iRow_Completed As Long
    Dim iRow_Cancelled As Long
    
    ' Find the last row with data in the 'Completed' sheet.
    iRow_Completed = ThisWorkbook.Worksheets("Completed").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Find the last row with data in the 'Cancelled' sheet.
    iRow_Cancelled = ThisWorkbook.Worksheets("Cancelled").Range("A" & Rows.Count).End(xlUp).Row
    
    ' Configure the 'lstCompleted' list box for completed orders & associated data.
    ASDEVS.lstCompleted.ColumnCount = 16  ' Set the number of columns.
    ASDEVS.lstCompleted.ColumnHeads = True ' Display column headers.

    ' Set the row source for the 'lstCompleted' list box based on the number of rows with data.
    If iRow_Completed = 1 Then
        ASDEVS.lstCompleted.RowSource = "Completed!A2:P2"  ' Use if only header row is present.
    Else
        ASDEVS.lstCompleted.RowSource = "Completed!A2:P" & iRow_Completed  ' Use if data rows are present.
    End If
    
    ' Configure the 'lstCancelled' list box for cancelled orders & associated data.
    ASDEVS.lstCancelled.ColumnCount = 20  ' Set the number of columns.
    ASDEVS.lstCancelled.ColumnHeads = True ' Display column headers.

    ' Set the row source for the 'lstCancelled' list box based on the number of rows with data.
    If iRow_Cancelled = 1 Then
        ASDEVS.lstCancelled.RowSource = "Cancelled!A2:T2"  ' Use if only header row is present.
    Else
        ASDEVS.lstCancelled.RowSource = "Cancelled!A2:T" & iRow_Cancelled  ' Use if data rows are present.
    End If
End Sub

Function Selected_List() As Long
    ' This function checks which item is selected in either of two list boxes (Completed or Cancelled)
    ' in the 'ASDEVS' and returns the index (position) of the selected item.
    ' If no item is selected, it returns 0.

    Dim i As Long  ' Variable to iterate through the list items.

    ' Initialise the return value to 0, indicating no item is selected initially.
    Selected_List = 0
    
    ' Determine which list box's selection to check.
    ' The value of 'ListBoxPages' determines which list box is active.
    If ASDEVS.ListBoxPages.value = 0 Then
        ' If 'ListBoxPages' value is 0, it indicates the 'Completed' list box is active.
        GoTo lstCompletedRow
    Else
        ' Any other value indicates the 'Cancelled' list box is active.
        GoTo lstCancelledRow
    End If

lstCompletedRow:
    ' Loop through the items in the 'Completed' list box.
    For i = 0 To ASDEVS.lstCompleted.ListCount - 1
        ' Check if the current item in the loop is selected.
        If ASDEVS.lstCompleted.Selected(i) = True Then
            ' If an item is selected, set the function's return value to the item's index (1-based) and exit the loop.
            Selected_List = i + 1
            Exit For
        End If
    Next i
    ' After checking the 'Completed' list box, proceed to the end of the function.

lstCancelledRow:
    ' Loop through the items in the 'Cancelled' list box.
    For i = 0 To ASDEVS.lstCancelled.ListCount - 1
        ' Check if the current item in the loop is selected.
        If ASDEVS.lstCancelled.Selected(i) = True Then
            ' If an item is selected, set the function's return value to the item's index (1-based) and exit the loop.
            Selected_List = i + 1
            Exit For
        End If
    Next i

    ' The function ends here, returning the index of the selected item or 0 if no item is selected.
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Last Entry/Edit Counter '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub LastEntry()
    ' This subroutine retrieves the latest entry details from the 'Support_Data' worksheet
    ' and updates the corresponding fields in the 'ASDEVS'.

    Dim sh As Worksheet

    ' Set a reference to the 'Support_Data' worksheet.
    Set sh = ThisWorkbook.Sheets("Support_Data")

    ' Retrieve and assign the value from the 'Edit By' field
    ' to the 'txtEditBy' textbox in 'ASDEVS'.
    ASDEVS.txtEditBy.value = sh.Cells(2, 9).value

    ' Retrieve and format the date and time from the 'Date/Time' field
    ' and assign it to the 'txtDateTime' textbox in 'ASDEVS'.
    ASDEVS.txtDateTime.value = Format(sh.Cells(2, 10).value, "dd/mm/yyyy hh:mm")
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Full Screen Functionality '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Maximize_Restore()
    ' This subroutine toggles the user form between a maximized state and its original size.
    ' It uses global variables (stated at the top of this module) to store the form's original dimensions and position.

    ' Check the current state of the form.
    ' If 'bState' is False (meaning the form is not maximized), maximize the form.
    If Not bState = True Then
        
        ' Store the current dimensions and position of the form in global variables.
        iWidth = ASDEVS.Width
        iHeight = ASDEVS.Height
        iTop = ASDEVS.Top
        iLeft = ASDEVS.Left
        
        ' Code to maximize the form to full screen.
        With Application
            ' Set the Excel application window state to maximized.
            .WindowState = xlMaximized
            
            ' Adjust the zoom of the user form to fit the maximized window.
            ASDEVS.Zoom = Int(.Width / ASDEVS.Width * 100)
            
            ' Set the form's position and size to fill the entire Excel window.
            ASDEVS.StartUpPosition = 0
            ASDEVS.Left = .Left
            ASDEVS.Top = .Top
            ASDEVS.Width = .Width
            ASDEVS.Height = .Height
        End With
        
        ' Change the caption of the full screen button to "Restore" and update the state flag.
        ASDEVS.cmdFullScreen.Caption = "Restore"
        bState = True
    
    ' If 'bState' is True (meaning the form is currently maximized), restore it to its original size.
    Else
        ' Code to restore the form to its original size and position.
        With Application
            ' Set the Excel application window state to normal.
            .WindowState = xlNormal
            
            ' Reset the zoom of the user form to 100%.
            ASDEVS.Zoom = 100
            
            ' Restore the form's original dimensions and position.
            ASDEVS.StartUpPosition = 0
            ASDEVS.Left = iLeft
            ASDEVS.Width = iWidth
            ASDEVS.Height = iHeight
            ASDEVS.Top = iTop
        End With
        
        ' Change the caption of the full screen button to "Full Screen" and update the state flag.
        ASDEVS.cmdFullScreen.Caption = "Full Screen"
        bState = False
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Function to populate the DateComboBox '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub PopulateDateComboBox(ByVal DateComboBox As MSForms.ComboBox)
    ' This subroutine populates a given DateComboBox with dates between a specified start and end date.
    ' It also includes additional dates for leap years.

    ' Declare variables to store the start and end dates, and the current date in the loop.
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date

    ' Define the start and end dates of the date range to populate.
    startDate = #1/1/2010#   ' Start date set to January 1, 2010.
    endDate = #12/31/2030#   ' End date set to December 31, 2030.
    ' Above dates can be altered as required.

    ' Loop through each date in the specified range.
    currentDate = startDate
    Do While currentDate <= endDate
        ' Add each date to the ComboBox, formatted as "dd/mm/yyyy".
        DateComboBox.AddItem Format(currentDate, "dd/mm/yyyy")
        currentDate = currentDate + 1 ' Increment the date by one day.
    Loop

    ' Manually add leap year dates to the ComboBox, although these are present in the iteration
    ' This provides a fail-safe.
    DateComboBox.AddItem Format("29/02/2012", "dd/mm/yyyy")
    DateComboBox.AddItem Format("29/02/2016", "dd/mm/yyyy")
    DateComboBox.AddItem Format("29/02/2020", "dd/mm/yyyy")
    DateComboBox.AddItem Format("29/02/2024", "dd/mm/yyyy")
    DateComboBox.AddItem Format("29/02/2028", "dd/mm/yyyy")

    ' Set the ComboBox's default value to today's date.
    DateComboBox.value = Format(Now, "dd/mm/yyyy")
End Sub






