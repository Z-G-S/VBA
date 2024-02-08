VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} User_Manual 
   Caption         =   "User Manual"
   ClientHeight    =   7329
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   12720
   OleObjectBlob   =   "User_Manual.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "User_Manual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' User_Manual Userform
' All code in this Module adds functionality to the User Manual userform, allowing the user to navigate to each page.
' It highlights the menus buttons based on page selected.



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Initialisation and Setup '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub userform_initialize()
    ' Highlights the DataEntry button if it's the default selected page.
    If MultiPage1.value = 0 Then
        DataEntryCB.BackColor = &H8000000D
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Form Control and Navigation '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub ResetButtonColours()
    ' Resets the background color of all control buttons on the User_Manual form.
    
    Dim Inputs As Collection
    Set Inputs = New Collection
    Dim CB() As Variant
    Dim var As Variant

    CB = Array(DataEntryCB, ListBoxCB, OrderCancellationCB, ExportOtherCB, LoginFormCB, ErrorCodesCB, HomeCB, PricesCostsCB, SupportDataCB, _
               CompletedCB, CancelledCB, CalculationsCB)
    
    For Each var In CB
        var.BackColor = &H3A1F1A
    Next
End Sub

Private Sub BackHelpButton_Click()
    ' Ensures that the colour of the User Manual button on the ASDEVS userform is reset and the now active page on the ASDEVS
    ' Userform is correctly highlighted.

    Select Case ASDEVS.MultiPage1.value
        Case 0
            ASDEVS.DataEntry.BackColor = &H8000000D
        Case 1
            ASDEVS.Visualisations.BackColor = &H8000000D
        Case 2
            ASDEVS.WorldSalesMap.BackColor = &H8000000D
    End Select
    
    ASDEVS.UserManual.BackColor = &H3A1F1A
    
    Unload Me
End Sub

' The remaining code enables the user to navigate to each page of the User_Manual, calls ResetButtonColours to
' Reset all button colours to default before highlighting the active page, providing an enhanced UX.

Private Sub DataEntryCB_Click()
    MultiPage1.value = 0
    Call ResetButtonColours
    DataEntryCB.BackColor = &H8000000D
End Sub

Private Sub ListBoxCB_Click()
    MultiPage1.value = 1
    Call ResetButtonColours
    ListBoxCB.BackColor = &H8000000D
End Sub

Private Sub OrderCancellationCB_Click()
    MultiPage1.value = 2
    Call ResetButtonColours
    OrderCancellationCB.BackColor = &H8000000D
End Sub
Private Sub ExportOtherCB_Click()
MultiPage1.value = 3
Call ResetButtonColours
    ExportOtherCB.BackColor = &H8000000D
End Sub
Private Sub LoginFormCB_Click()
MultiPage1.value = 4
Call ResetButtonColours
    LoginFormCB.BackColor = &H8000000D
End Sub
Private Sub ErrorCodesCB_Click()
MultiPage1.value = 5
Call ResetButtonColours
    ErrorCodesCB.BackColor = &H8000000D
End Sub
Private Sub HomeCB_Click()
MultiPage1.value = 6
Call ResetButtonColours
    HomeCB.BackColor = &H8000000D
End Sub
Private Sub PricesCostsCB_Click()
MultiPage1.value = 7
Call ResetButtonColours
    PricesCostsCB.BackColor = &H8000000D
End Sub
Private Sub SupportDataCB_Click()
MultiPage1.value = 8
Call ResetButtonColours
    SupportDataCB.BackColor = &H8000000D
End Sub
Private Sub CompletedCB_Click()
MultiPage1.value = 9
Call ResetButtonColours
    CompletedCB.BackColor = &H8000000D
End Sub
Private Sub CancelledCB_Click()
MultiPage1.value = 10
Call ResetButtonColours
    CancelledCB.BackColor = &H8000000D
End Sub
Private Sub CalculationsCB_Click()
MultiPage1.value = 10
Call ResetButtonColours
    CalculationsCB.BackColor = &H8000000D
End Sub
