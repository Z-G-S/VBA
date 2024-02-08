VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LoginForm 
   Caption         =   "Login Form"
   ClientHeight    =   2576
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   4440
   OleObjectBlob   =   "LoginForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' LoginForm Userform
' This module handles the sign-in process to access the Excel Sheets and VBA Application.
' It checks the user's credentials and, if correct, makes all sheets visible.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' User Authentication and Workbook/Data Sheets access '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub cmdSignIn_Click()
    ' This subroutine is executed when the Sign-In button is clicked.

    ' Initialize flag for empty fields
    Dim emptyFields As Boolean
    emptyFields = False

    ' Reset background color for text boxes
    LoginForm.txtID.BackColor = vbWhite
    LoginForm.txtPassword.BackColor = vbWhite

    ' Check for empty ID field
    If Trim(LoginForm.txtID.value) = "" Then
        LoginForm.txtID.BackColor = vbRed
        emptyFields = True
    End If

    ' Check for empty Password field
    If Trim(LoginForm.txtPassword.value) = "" Then
        LoginForm.txtPassword.BackColor = vbRed
        emptyFields = True
    End If

    ' If either field is empty, prompt the user and exit the subroutine
    If emptyFields Then
        MsgBox "Please enter an input for the boxes that are highlighted red.", vbCritical + vbOKOnly, "No Input Error 1001"
        Exit Sub
    End If

    ' Check if the entered ID and Password match the predefined credentials ("Test" and "123").
    If LoginForm.txtID.value = "Test" And LoginForm.txtPassword.value = "123" Then
    
        ' Make several sheets visible to the user after successful login.
        ActiveWorkbook.Sheets("Data_Entry").Visible = True
        ActiveWorkbook.Sheets("Completed").Visible = True
        ActiveWorkbook.Sheets("Cancelled").Visible = True
        ActiveWorkbook.Sheets("Calculations").Visible = True
        ActiveWorkbook.Sheets("Support_Data").Visible = True
        ActiveWorkbook.Sheets("Prices_Costs_Data").Visible = True
        
        ' Save the current state of the workbook.
        ThisWorkbook.Save
        
        ' Activate the Data_Entry sheet
        Sheet2.Activate
        
        ' Unload the LoginForm and show a success message.
        Unload Me
        MsgBox "Login Successful!"
        Exit Sub
    Else
        ' If login credentials do not match, show an error message and clear the input fields.
        MsgBox "Incorrect Login Information!", vbOKOnly + vbCritical, "Incorrect Entry Error 6001"
        LoginForm.txtID.value = ""
        LoginForm.txtPassword.value = ""
    End If

End Sub



