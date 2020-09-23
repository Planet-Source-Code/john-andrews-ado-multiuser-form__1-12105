VERSION 5.00
Begin VB.Form frmMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main Menu"
   ClientHeight    =   3240
   ClientLeft      =   7695
   ClientTop       =   6015
   ClientWidth     =   5025
   Icon            =   "MainMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5025
   Begin VB.CommandButton cmdOpenWizardForm 
      Caption         =   "&VB Wizard Form"
      Height          =   495
      Left            =   1125
      TabIndex        =   1
      ToolTipText     =   "Examine the VB wizard's creation for comparison"
      Top             =   1500
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   495
      Left            =   3510
      TabIndex        =   2
      ToolTipText     =   "Close the application"
      Top             =   2550
      Width           =   1215
   End
   Begin VB.CommandButton cmdSingleRecord 
      Caption         =   "&Multiuser Data Form"
      Height          =   495
      Left            =   1155
      TabIndex        =   0
      ToolTipText     =   "Launch an instance of the ADO multiuser data form"
      Top             =   885
      Width           =   2160
   End
   Begin VB.Shape Shape1 
      Height          =   2190
      Left            =   120
      Top             =   120
      Width           =   4620
   End
   Begin VB.Label Label1 
      Caption         =   "Each click opens a new instance of the data form."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   330
      TabIndex        =   3
      Top             =   240
      Width           =   3660
   End
End
Attribute VB_Name = "frmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
 'Close the application
 
     On Error GoTo ErrorHandler
     
    Unload Me
    
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub

Private Sub cmdOpenWizardForm_Click()
 'Open an instance of the data form

    Dim frmNew As frmVBWizard
    Static intFormCount As Integer
    
    On Error GoTo ErrorHandler
    
    intFormCount = intFormCount + 1
    
    Set frmNew = New frmVBWizard
    Load frmNew
    frmNew.Caption = "VB Wizard Form, Instance #" & intFormCount
    frmNew.Show
    
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub


Private Sub cmdSingleRecord_Click()
 'Open an instance of the data form

    Dim frmNew As frmSingleRec
    Static intFormCount As Integer
    
    On Error GoTo ErrorHandler
    
    intFormCount = intFormCount + 1
    
    Set frmNew = New frmSingleRec
    Load frmNew
    frmNew.Caption = "Multiuser Data Form, Instance #" & intFormCount
    frmNew.Show
    
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Unload application properly ensuring restoration of resources

    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    While Forms.Count > 1
        ' Find first form besides "me" to unload
        i = 0
        While Forms(i).Caption = Me.Caption
             i = i + 1
        Wend
        Unload Forms(i)
    Wend
    
    ' Last thing to be done...
    Unload Me
    End

    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)

End Sub

Private Sub DisplayError(rstrErrorMessage As String)
    On Error Resume Next
    MsgBox rstrErrorMessage, vbInformation, "Error Message"
End Sub
