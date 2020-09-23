VERSION 5.00
Begin VB.Form frmSingleRec 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Form"
   ClientHeight    =   2715
   ClientLeft      =   1635
   ClientTop       =   4905
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   Icon            =   "SingleRec.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2715
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkHasSquirrels 
      Appearance      =   0  'Flat
      DataField       =   "HasSquirrels"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1035
      TabIndex        =   20
      Top             =   1530
      Width           =   330
   End
   Begin VB.TextBox txtLeafTypeID 
      Appearance      =   0  'Flat
      DataField       =   "LeafTypeID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1035
      TabIndex        =   18
      Top             =   1155
      Width           =   660
   End
   Begin VB.TextBox txtTreeName 
      Appearance      =   0  'Flat
      DataField       =   "TreeName"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1035
      TabIndex        =   16
      Top             =   780
      Width           =   1170
   End
   Begin VB.TextBox txtTreeID 
      Appearance      =   0  'Flat
      DataField       =   "TreeID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1035
      TabIndex        =   14
      Top             =   390
      Width           =   660
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4290
      TabIndex        =   4
      ToolTipText     =   "Close this form"
      Top             =   2070
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3210
      TabIndex        =   3
      ToolTipText     =   "Requery database to get latest data"
      Top             =   2070
      Width           =   1095
   End
   Begin VB.PictureBox picStatBox 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   5400
      TabIndex        =   9
      Top             =   2415
      Width           =   5400
      Begin VB.CommandButton cmdMoveNext 
         Height          =   300
         Left            =   3000
         Picture         =   "SingleRec.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdMovePrevious 
         Height          =   300
         Left            =   360
         Picture         =   "SingleRec.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdMoveFirst 
         Height          =   300
         Left            =   0
         Picture         =   "SingleRec.frx":0AC6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdMoveLast 
         Height          =   300
         Left            =   3360
         Picture         =   "SingleRec.frx":0E08
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   290
         Left            =   705
         TabIndex        =   10
         Top             =   15
         Width           =   2280
      End
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1050
      TabIndex        =   1
      ToolTipText     =   "Add a new record"
      Top             =   2070
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2130
      TabIndex        =   2
      ToolTipText     =   "Delete current record"
      Top             =   2070
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -30
      TabIndex        =   0
      ToolTipText     =   "Edit current record"
      Top             =   2070
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1050
      TabIndex        =   12
      ToolTipText     =   "Cancel changes"
      Top             =   2070
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -30
      TabIndex        =   11
      ToolTipText     =   "Save changes to record"
      Top             =   2070
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "No duplicate names"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   2370
      TabIndex        =   24
      Top             =   810
      Width           =   2280
   End
   Begin VB.Label Label4 
      Caption         =   "Autonumber fields can't be edited"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1785
      TabIndex        =   23
      Top             =   405
      Width           =   3330
   End
   Begin VB.Label Label2 
      Caption         =   "Type ID links to another table. Values of 1 or 2 are allowed."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1785
      TabIndex        =   22
      Top             =   1140
      Width           =   3330
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   2100
      TabIndex        =   21
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Squirrels:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   210
      TabIndex        =   19
      Top             =   1575
      Width           =   795
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Type ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   345
      TabIndex        =   17
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Tree Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   30
      TabIndex        =   15
      Top             =   810
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   765
      TabIndex        =   13
      Top             =   435
      Width           =   225
   End
End
Attribute VB_Name = "frmSingleRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "hello john!"
Option Explicit

Enum NavButtonModeEnum 'passed to UpdateNavigationButtonsAndRecPosStatus
       Addmode = 3
       EditMode = 4
End Enum

Enum DisplayModeEnum 'passed to UpdateNavigationButtonsAndRecPosStatus
       BrowseOnly = False
       AllowEdit = True
End Enum

Private WithEvents mrecPrimary As Recordset
Attribute mrecPrimary.VB_VarHelpID = -1

Private Sub cmdClose_Click()
 'Close form
On Error GoTo ErrorHandler
    Unload Me
      
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub

Private Sub cmdMoveFirst_Click()
' Move to first record in set

    On Error GoTo ErrorHandler
    
    mrecPrimary.MoveFirst
    UpdateNavigationButtonsAndRecPosStatus
     
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub

Private Sub cmdMoveLast_Click()
 'Move to last item in recordset
 
    On Error GoTo ErrorHandler
    
    If mrecPrimary.EOF And mrecPrimary.BOF Then Exit Sub
    mrecPrimary.MoveLast
    UpdateNavigationButtonsAndRecPosStatus
     
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
     
End Sub

Sub UpdateNavigationButtonsAndRecPosStatus(Optional rMode As NavButtonModeEnum)
'Disable the record navigation buttons at end of recordset
'Should be impossible for record to be on EOF, because button disabling will prevent the user
'from moving that far.
'Also updates record position caption: ie: 'Record x of y'

    Dim blnCanMoveForward As Boolean, blnCanMoveBack As Boolean
    On Error GoTo ErrorHandler
    
    blnCanMoveForward = True
    blnCanMoveBack = True
    
    Select Case rMode
        Case Addmode
            lblStatus = "New record"
        Case EditMode
                lblStatus = "Editing record " & CStr(mrecPrimary.AbsolutePosition)
        Case Else
                lblStatus = " Record " & CStr(mrecPrimary.AbsolutePosition) & " of " & mrecPrimary.RecordCount
    End Select
        
    'disable navigation buttons while user is adding or editing record
    If rMode = Addmode Or rMode = EditMode Then
        cmdMoveLast.Enabled = False
        cmdMoveNext.Enabled = False
        cmdMoveFirst.Enabled = False
        cmdMovePrevious.Enabled = False
    Else
        'User is browsing. Enable buttons as appropriate.
        'Disable buttons if neccessary to prevent user moving past end of recordset.
        
        'Take a look at the next record to see if it's EOF
        '(EOF can't be detected by merely sitting on last record)
        mrecPrimary.MoveNext
        If mrecPrimary.EOF Then blnCanMoveForward = False
        mrecPrimary.MovePrevious 'back to original position
        
        'Check for BOF
        'Note that both BOF and EOF can be valid simultaneously if the recordset is empty
        mrecPrimary.MovePrevious
        If mrecPrimary.BOF Then blnCanMoveBack = False
        mrecPrimary.MoveNext 'back to original position
        
        'Set buttons
        cmdMoveLast.Enabled = blnCanMoveForward
        cmdMoveNext.Enabled = blnCanMoveForward
        cmdMoveFirst.Enabled = blnCanMoveBack
        cmdMovePrevious.Enabled = blnCanMoveBack
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox Err.Description & " (Source: UpdateNavigationButtonsAndRecPosStatus)", vbExclamation, "Error"
   
End Sub

Private Sub cmdMoveNext_Click()
 'Move to next record
 
    On Error GoTo ErrorHandler
    
    mrecPrimary.MoveNext
    'Button disabling will have prevented user moving onto EOF
    
    'automatically skip over deleted rows
    Do While mrecPrimary.Status = adRecDBDeleted
        mrecPrimary.MoveNext
        If mrecPrimary.EOF Then
            mrecPrimary.MovePrevious
            Exit Do
        End If
    Loop
    UpdateNavigationButtonsAndRecPosStatus
     
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
         
End Sub

Private Sub cmdMovePrevious_Click()
 'Move to previous record

    On Error GoTo ErrorHandler

    mrecPrimary.MovePrevious
    'Button disabling will have prevented user moving onto EOF
    
    'automatically skip over deleted rows (these are not removed from a client-side recordset)
    'If we were using the ado data control, it would skip over deleted rows automatically
    Do While mrecPrimary.Status = adRecDBDeleted
        mrecPrimary.MovePrevious
        If mrecPrimary.EOF Then
            mrecPrimary.MoveNext
            Exit Do
        End If
    Loop
    UpdateNavigationButtonsAndRecPosStatus
     
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub

Private Sub cmdAddNew_Click()
 'Add a new record

    On Error GoTo ErrorHandler

    SetFormEditMode (AllowEdit)
    UpdateNavigationButtonsAndRecPosStatus Addmode
    mrecPrimary.AddNew
         
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub
Sub SetFormEditMode(rblnEditMode As DisplayModeEnum)
 'browsemode: Lock Databound Controls and hide cancel and save buttons
 'editmode: Unlock Databound Controls and show cancel and save buttons


    
    Dim oControl As Control
    Const EDIT_BACKCOLOUR As Long = &HC000C0
    Const LOCKED_BACKCOLOUR As Long = &H8000000F
    
    On Error GoTo ErrorHandler
    
    For Each oControl In Me.Controls  'Optionally apply a visual effect to indicate to user whether controls can be edited
        Select Case TypeName(oControl)
            Case "TextBox", "DataCombo", "CheckBox" 'add other control types to this list if required
                If oControl.DataField <> "" Then oControl.Enabled = rblnEditMode 'it's a databound control
                'the enabled property is the only one that applies to all controls including checkboxes.
                'has the disadvantage that everything turns grey. If only checkboxes had a 'locked' prop...
                If mrecPrimary.Fields(oControl.DataField).Properties("ISAUTOINCREMENT") Then oControl.Enabled = False
                'lock autonumber column, an error would occur if user attempted to edit it
                'note that newly-added autoincrement fields are not returned to the cursor by Access 97 databases
                    
                'This will colour-change the control's backgrounds
                'oControl.Locked = rblnEditMode - this technique will not work for all types of controls, eg checkboxes
                'oControl.BackColor = IIf(rblnEditMode, EDIT_BACKCOLOUR, LOCKED_BACKCOLOUR)
        End Select
   Next oControl
   
   'Now show or hide command buttons
    cmdAddNew.Visible = Not rblnEditMode
    cmdDelete.Visible = Not rblnEditMode
    cmdEdit.Visible = Not rblnEditMode
    cmdRefresh.Visible = Not rblnEditMode
    cmdClose.Visible = Not rblnEditMode
    cmdSave.Visible = rblnEditMode
    cmdCancel.Visible = rblnEditMode
         
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub
Sub RebindControls()
 'Forces a display update
 
    On Error GoTo ErrorHandler
        
    Dim oControl As Control
    For Each oControl In Me.Controls
        If TypeName(oControl) = "TextBox" Or TypeName(oControl) = "CheckBox" Or TypeName(oControl) = "DataCombo" Then
            Set oControl.DataSource = mrecPrimary
        End If
   Next oControl
        
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub

Private Sub cmdCancel_Click()
 'User cancels change
 '
 
    On Error GoTo ErrorHandler
    mrecPrimary.CancelUpdate
    'databound controls will revert to previous values
    SetFormEditMode (BrowseOnly)
    UpdateNavigationButtonsAndRecPosStatus
         
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub

Private Sub cmdDelete_Click()
 'Delete the current record

    On Error GoTo ErrorHandler
    
    If MsgBox("Record will be deleted permanently. Proceed?", vbYesNoCancel + vbExclamation, "Confirm Delete") <> vbYes Then Exit Sub
        
    On Error Resume Next
    mrecPrimary.Delete
    
    Select Case Err.Number
        Case 0: 'delete was successful
            If cmdMoveNext.Enabled Then
                cmdMoveNext_Click
            Else
                If cmdMovePrevious.Enabled Then cmdMovePrevious_Click
            End If
        Case -2147217864
            MsgBox "This row had already been deleted by another user!", vbInformation
            mrecPrimary.CancelUpdate
        Case Else
            MsgBox "Could not delete record." + vbCrLf + Err.Description
            mrecPrimary.CancelUpdate
    End Select
       
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub

Private Sub cmdEdit_Click()
 'Enable editing of the databound controls
 
    On Error GoTo ErrorHandler
    
    mrecPrimary.Resync adAffectCurrent
    'get latest data for the edit
    mrecPrimary.Move 0  'update the bound controls
    
    SetFormEditMode (AllowEdit)
    UpdateNavigationButtonsAndRecPosStatus (EditMode)
    
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case -2147217885 'row was deleted (the resync might trigger this)
            MsgBox "This row has been deleted by another user.", vbInformation
        Case Else
            DisplayError (Err.Description)
    End Select
        
End Sub

Private Sub cmdRefresh_Click()
'Get latest data
  On Error GoTo RefreshErr
  Screen.MousePointer = vbHourglass
  mrecPrimary.Requery
  'the resync method won't show new records, and confuses ADO if there are any - gives 'deleted from local data store error'
  Screen.MousePointer = vbNormal
  UpdateNavigationButtonsAndRecPosStatus
  Exit Sub
RefreshErr:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub cmdSave_Click()
 'Save edits to database

    Dim vFieldArray(), x As Integer, intUpdateError As Integer, strErrorMessage As String, oError As Error
    Dim blnAdd As Boolean
    On Error Resume Next 'this also clears the error object
    
    blnAdd = mrecPrimary.EditMode = adEditAdd
    mrecPrimary.ActiveConnection.Errors.Clear
    
    Screen.MousePointer = vbHourglass 'the update might take a while
    mrecPrimary.Update
    Screen.MousePointer = vbNormal
    intUpdateError = Err.Number 'remember any errors (error object will be reset by the next line)
    On Error GoTo ErrorHandler
    
    Select Case intUpdateError
        Case 0:
            'Check the underlying connection object for errors too. Provider-specific errors don't show up in the above error trap
            If mrecPrimary.ActiveConnection.Errors.Count = 0 Then 'no provider-specific errors - save was successful
                SetFormEditMode (BrowseOnly)
                UpdateNavigationButtonsAndRecPosStatus
                If blnAdd Then
                    mrecPrimary.Resync adAffectCurrent 'show default field values that may have been entered by the db
                    mrecPrimary.Move 0 'force databound controls to display new info
                End If
            Else
                For Each oError In mrecPrimary.ActiveConnection.Errors
                    strErrorMessage = strErrorMessage & oError.Description & vbCr
                Next
                If mrecPrimary.ActiveConnection.Errors.Count = 1 Then strErrorMessage = "The following error" & _
                    IIf(mrecPrimary.ActiveConnection.Errors.Count > 1, "s were", " was") & " reported by the provider: " & vbCr & strErrorMessage
                MsgBox strErrorMessage 'display all the errors
            End If
            'leave the Save and Cancel buttons showing so that user can backtrack
            
        Case 3640 + vbObjectError 'Record changed by another user
            If MsgBox("Another user has changed this record since you started editing it. If you save the record, you will overwrite the changes the other user made." & vbCr & vbCr & "Do you want to overwrite the other user's changes?", vbExclamation + vbYesNoCancel, "Write Conflict") = vbYes Then
                'To force an overwrite, we must cache this user's edits, refresh the cursor, and then re-save.
                ReDim vFieldArray(mrecPrimary.Fields.Count - 1)
                For x = 0 To mrecPrimary.Fields.Count - 1 'stash user's changes (the database won't accept them yet)
                vFieldArray(x) = mrecPrimary.Fields(x).Value
                Next x
                
                mrecPrimary.CancelUpdate
                mrecPrimary.Resync adAffectCurrent 'fool database into thinking these are the latest changes
                
                For x = 0 To mrecPrimary.Fields.Count - 1 'save this user's cached edits into database
                    If mrecPrimary.Fields(x).Value <> vFieldArray(x) Then mrecPrimary.Fields(x) = vFieldArray(x)
                Next x
                
                mrecPrimary.Update
                'LockDataboundControls (True)
                SetFormEditMode (BrowseOnly)
                UpdateNavigationButtonsAndRecPosStatus
            Else 'user elects not to overwrite database record
                mrecPrimary.CancelUpdate
                mrecPrimary.Resync adAffectCurrent 'show the latest data
                mrecPrimary.Move 0
            End If
        Case -2147467259
            MsgBox "The changes you have made can't be saved at the moment. The record is currently locked by another user." & vbCr & "You can cancel your changes, or retry the save again later.", vbExclamation, "Write Error"  'eg record has been deleted
            Exit Sub
        Case Else:
                MsgBox Err.Description + vbCr & "(Source: cmdSave_Click)", vbExclamation, "Error"
    End Select
    
    Exit Sub
ErrorHandler:
    If Err.Number = -2147217885 Then 'Key value for this row was changed or deleted at the data store. The local row is now deleted.
        MsgBox "Access 97 databases don't update client with newly added autonumber fields. Press the Refresh button to see the new record.", vbInformation
        If cmdRefresh.Visible Then cmdRefresh.SetFocus
    Else
        DisplayError (Err.Description)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 'Enable keyboard shortcuts for record navigation whilst browsing
 
    On Error GoTo ErrorHandler
    
    If cmdSave.Visible Then Exit Sub 'user editing or adding new record, don't use accellerators
    
    Select Case KeyCode
      Case vbKeyEnd
         If cmdMoveLast.Enabled Then cmdMoveLast_Click
      Case vbKeyHome
         If cmdMoveFirst.Enabled Then cmdMoveFirst_Click
      Case vbKeyUp, vbKeyPageUp
        If Shift = vbCtrlMask Then
           If cmdMoveFirst.Enabled Then cmdMoveFirst_Click
        Else
           If cmdMovePrevious.Enabled Then cmdMovePrevious_Click
        End If
      Case vbKeyDown, vbKeyPageDown
        If Shift = vbCtrlMask Then
          If cmdMoveLast.Enabled Then cmdMoveLast_Click
        Else
          If cmdMoveNext.Enabled Then cmdMoveNext_Click
        End If
    End Select
     
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Dim db As Connection
    Set db = New Connection
    db.CursorLocation = adUseClient
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Trees97.mdb;"
    
    Set mrecPrimary = New ADODB.Recordset
    mrecPrimary.Open "Select HasSquirrels, LeafTypeID, TreeID, TreeName from tblTree", db, adOpenStatic, adLockOptimistic

    RebindControls
    AutoConfigureDataboundControls

    SetFormEditMode (BrowseOnly)

    cmdMoveFirst_Click
    
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
   
End Sub

Private Sub AutoConfigureDataboundControls()
 'Apply input restrictions to the databound controls to match their bound datatypes
 'eg. Set maxlength property of text boxes to match their respective bound field lengths.
 
 'This function can be expanded later, at present it just sets the maxlength of textboxes
 
    Dim oControl As Control
    On Error GoTo ErrorHandler
    
    For Each oControl In Me.Controls
        If TypeName(oControl) = "TextBox" Then
            'set maxlength for text fields
            If oControl.DataField <> "" Then
                If mrecPrimary.Fields(oControl.DataField).Type = adVarWChar Then oControl.MaxLength = mrecPrimary.Fields(oControl.DataField).DefinedSize
            End If
        End If
    Next
    
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error GoTo ErrorHandler

    If cmdSave.Visible Then If MsgBox("You are still in edit mode. Any changes you have made will be lost if you close the form now." & _
               vbCr & vbCr & "Are you sure you want to close the form?", vbInformation + vbYesNoCancel) <> vbYes Then Cancel = True

    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'Tidy up used rescources
On Error GoTo ErrorHandler

    mrecPrimary.Close
    Set mrecPrimary = Nothing
    
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)

End Sub



Private Sub mrecPrimary_FieldChangeComplete(ByVal cFields As Long, ByVal Fields As Variant, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'Trap invalid field data as soon as it's entered by the user.
 'If these errors are ignored, then ADO will throw an 'operation cancelled' error when we attempt to move the cursor

    'This event is fired when focus leaves an edited databound control.
    'and is called again after the update method has been called.
    'The passed 'Fields' object contains the field with errors
    
    On Error GoTo ErrorHandler
    
    If adStatus = adStatusErrorsOccurred Then
        Beep
        'TO_DO: restore focus to previous control and force user to re-enter
        MsgBox "The value you entered isn't valid for this field." & vbCr & vbCr & _
            "For example, you may have entered text in a numeric field or a number that is larger than the FieldSize setting permits.", vbInformation
        Me.ActiveControl.Text = mrecPrimary.Fields(Me.ActiveControl.DataField).OriginalValue
        'restore control to its previous value
    End If
    
    Exit Sub
ErrorHandler:
    DisplayError (Err.Description)
    
End Sub

Private Sub DisplayError(rstrErrorMessage As String)
    On Error Resume Next
    MsgBox rstrErrorMessage, vbInformation, "Error Message"
End Sub
