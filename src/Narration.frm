VERSION 5.00
Begin VB.Form frmNarration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Narration Details"
   ClientHeight    =   3165
   ClientLeft      =   2160
   ClientTop       =   1545
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data datNarration 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraNarration 
      Height          =   735
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   5895
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   270
         Left            =   135
         TabIndex        =   19
         Top             =   285
         Width           =   645
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify"
         Height          =   270
         Left            =   780
         TabIndex        =   18
         Top             =   285
         Width           =   645
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   270
         Left            =   1425
         TabIndex        =   17
         Top             =   285
         Width           =   645
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   270
         Left            =   2205
         TabIndex        =   16
         Top             =   285
         Width           =   645
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   270
         Left            =   2850
         TabIndex        =   15
         Top             =   285
         Width           =   645
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Ca&ncel"
         Height          =   270
         Left            =   3495
         TabIndex        =   14
         Top             =   285
         Width           =   645
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   270
         Left            =   4290
         TabIndex        =   13
         Top             =   285
         Width           =   645
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   270
         Left            =   5085
         TabIndex        =   12
         Top             =   285
         Width           =   645
      End
   End
   Begin VB.Frame fraNarration 
      Height          =   1815
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5895
      Begin VB.TextBox txtCode 
         DataField       =   "Narr_Code"
         DataSource      =   "datNarration"
         Height          =   300
         Left            =   1455
         MaxLength       =   2
         TabIndex        =   8
         Top             =   465
         Width           =   1440
      End
      Begin VB.TextBox txtDesc 
         DataField       =   "Narr_Desc"
         DataSource      =   "datNarration"
         Height          =   300
         Left            =   1455
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1140
         Width           =   4290
      End
      Begin VB.Frame fraNarration 
         Height          =   600
         Index           =   2
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   2010
         Begin VB.CommandButton cmdFirst 
            Caption         =   "<<"
            Height          =   255
            Left            =   165
            TabIndex        =   6
            Top             =   210
            Width           =   420
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "<"
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   210
            Width           =   420
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   ">"
            Height          =   255
            Left            =   1020
            TabIndex        =   4
            Top             =   210
            Width           =   420
         End
         Begin VB.CommandButton cmdLast 
            Caption         =   ">>"
            Height          =   255
            Left            =   1440
            TabIndex        =   3
            Top             =   210
            Width           =   420
         End
      End
      Begin VB.ComboBox cboCode 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   458
         Width           =   1440
      End
      Begin VB.Label lblNarration 
         Caption         =   "Narration Code"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   1125
      End
      Begin VB.Label lblNarration 
         Caption         =   "Description"
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1170
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmNarration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mdbsAccounts As Database
Dim mrstNarration As Recordset

Dim mblnTrans As Boolean
Dim mblnFind As Boolean
Dim mblnClose As Boolean

Dim mvntBookmark As Variant
Private Sub cmdAdd_Click()
    mblnTrans = True
    If Not mrstNarration.EOF Then
        mvntBookmark = mrstNarration.Bookmark
        mrstNarration.MoveLast
    End If
    Call Enable_Controls(False)
    Call Lock_Controls(False)
    Call DisableNaviButtons
    
    mrstNarration.AddNew
    txtCode.SetFocus
End Sub


Private Sub cmdCancel_Click()

    If mblnTrans Then
        mblnTrans = False
        mrstNarration.CancelUpdate
        If Not mrstNarration.EOF Then
            mrstNarration.Bookmark = mvntBookmark
        End If
    ElseIf mblnFind Then
        mblnFind = False
        mrstNarration.Bookmark = mvntBookmark
        txtCode.Visible = True
        cboCode.Visible = False
        cmdFind.Caption = "&Find"
    End If
    Call Enable_Controls(True)
    Call Lock_Controls(True)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If MsgBox("Delete the Entry?", vbYesNo) = vbYes Then
        mrstNarration.Delete
        mrstNarration.MoveNext
        If mrstNarration.EOF And mrstNarration.RecordCount > 0 Then
            mrstNarration.MoveLast
        End If
        Call Enable_Controls(True)
    End If
    
End Sub

Private Sub cmdFind_Click()

Dim rstCode As Recordset
    If Not mblnFind Then
        mvntBookmark = mrstNarration.Bookmark
        mblnFind = True
        cmdFind.Caption = "&Display"
        Call DisableNaviButtons
        Call Enable_Controls(False)
        txtCode.Visible = False
        cboCode.Visible = True
        cboCode.Clear
        cboCode.Text = ""
        Set rstCode = mdbsAccounts.OpenRecordset("Select Narr_Code from NarrCode", dbOpenSnapshot)
        While Not rstCode.EOF
            cboCode.AddItem rstCode!Narr_Code
            rstCode.MoveNext
        Wend
        cboCode.SetFocus
    ElseIf mblnFind Then
        If cboCode.Text = "" Then
            MsgBox "Select Code from the list"
            cboCode.SetFocus
        Else
            mblnFind = False
            cmdFind.Caption = "&Find"
            mrstNarration.FindFirst "Narr_Code = '" & cboCode.Text & "'"
            If mrstNarration.NoMatch Then
                MsgBox "Entered Code does not exist"
                mrstNarration.Bookmark = mvntBookmark
            End If
            Call Enable_Controls(True)
            Call Lock_Controls(True)
            txtCode.Visible = True
            cboCode.Visible = False
        End If
    End If
End Sub

Private Sub cmdFirst_Click()
    mrstNarration.MoveFirst
End Sub

Private Sub cmdLast_Click()
    mrstNarration.MoveLast
End Sub

Private Sub cmdModify_Click()
    mblnTrans = True
    Call Enable_Controls(False)
    Call Lock_Controls(False)
    Call DisableNaviButtons
    mvntBookmark = mrstNarration.Bookmark
    mrstNarration.Edit
    txtCode.Locked = True
    txtDesc.SetFocus
End Sub

Private Sub cmdNext_Click()
    mrstNarration.MoveNext
End Sub

Private Sub cmdPrevious_Click()
    mrstNarration.MovePrevious
End Sub


Private Sub cmdReset_Click()
    datNarration.UpdateControls
    If mrstNarration.EditMode = dbEditAdd Then
        txtCode.SetFocus
    ElseIf mrstNarration.EditMode = dbEditInProgress Then
        txtDesc.SetFocus
    End If
End Sub

Private Sub cmdSave_Click()

On Error GoTo ErrorHandler
    If FieldCheck Then
        mblnTrans = False
        mrstNarration.Update
        mrstNarration.Bookmark = mrstNarration.LastModified
        Call Enable_Controls(True)
        Call Lock_Controls(True)
    End If
    Exit Sub

ErrorHandler:
    If Err.Number = 3022 Then
        MsgBox "Code already exist"
        mblnTrans = True
        txtCode.SetFocus
    Else
        MsgBox Err.Description
    End If
    Exit Sub
End Sub

Private Sub datNarration_Reposition()

    If Not mblnTrans And Not mblnFind Then
        Call RecordPosition
    End If
    
End Sub

Private Sub Form_Load()

    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    Set mrstNarration = mdbsAccounts.OpenRecordset("NarrCode", dbOpenDynaset)
    Set datNarration.Recordset = mrstNarration
    
    If Not mrstNarration.EOF Then
        mrstNarration.MoveLast
        mrstNarration.MoveFirst
    End If
    
    mblnTrans = False
    mblnFind = False
    mblnClose = False
    
    Call Lock_Controls(True)
    Call Enable_Controls(True)
    Call RecordPosition
End Sub
Private Sub Enable_Controls(blnEnable As Boolean)
    cmdAdd.Enabled = blnEnable
    If blnEnable And mrstNarration.RecordCount <> 0 Then
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
    If (blnEnable And mrstNarration.RecordCount > 1) Or mblnFind Then
        cmdFind.Enabled = True
    Else
        cmdFind.Enabled = False
    End If
    
    cmdSave.Enabled = mblnTrans
    cmdReset.Enabled = mblnTrans
    
    If mblnTrans Or mblnFind Then
        cmdCancel.Enabled = True
    Else
        cmdCancel.Enabled = False
    End If
    cmdClose.Enabled = blnEnable
End Sub



Private Sub Lock_Controls(blnLock As Boolean)
    txtCode.Locked = blnLock
    txtDesc.Locked = blnLock
End Sub

Private Sub RecordPosition()
    If mrstNarration.RecordCount = 0 Or mrstNarration.RecordCount = 1 Then
        cmdFirst.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = False
        cmdLast.Enabled = False
    ElseIf mrstNarration.AbsolutePosition = 0 Then
        cmdFirst.Enabled = False
        cmdPrevious.Enabled = False
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    ElseIf mrstNarration.AbsolutePosition = (mrstNarration.RecordCount - 1) Then
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = False
        cmdLast.Enabled = False
    Else
        cmdFirst.Enabled = True
        cmdPrevious.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    End If
End Sub

Private Sub DisableNaviButtons()

    cmdFirst.Enabled = False
    cmdPrevious.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    
End Sub


Private Function FieldCheck() As Boolean
    FieldCheck = False
    If txtCode.Text = "" Then
        MsgBox "Code should not be blank"
        txtCode.SetFocus
    ElseIf txtDesc.Text = "" Then
        MsgBox "Description should not be blank"
        txtDesc.SetFocus
    Else
        FieldCheck = True
    End If
End Function

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) And (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
        (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
        (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        Beep
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub


Private Sub txtDesc_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) And (KeyAscii <> Asc(" ")) And _
        (KeyAscii <> Asc("(")) And (KeyAscii <> Asc(")")) And _
        (KeyAscii <> Asc(".")) And (KeyAscii <> Asc(",")) And _
        (KeyAscii <> Asc("&")) And (KeyAscii <> Asc("-")) And _
        (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
        (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
        (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        Beep
        KeyAscii = 0
    Else
        If Len(txtDesc.Text) = 0 Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
End Sub
