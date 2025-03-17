VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Journal Transaction"
   ClientHeight    =   4995
   ClientLeft      =   825
   ClientTop       =   2055
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8055
   Begin VB.Data datJournal 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   285
      Left            =   7290
      TabIndex        =   14
      Top             =   255
      Width           =   465
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   285
      Left            =   6810
      TabIndex        =   13
      Top             =   255
      Width           =   465
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   285
      Left            =   6345
      TabIndex        =   12
      Top             =   255
      Width           =   465
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   285
      Left            =   5880
      TabIndex        =   11
      Top             =   255
      Width           =   465
   End
   Begin TabDlg.SSTab sstabJournal 
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Form View"
      TabPicture(0)   =   "Journal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraJournal(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraJournal(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "SpreedSheet View"
      TabPicture(1)   =   "Journal.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbgrdJournal"
      Tab(1).ControlCount=   1
      Begin MSDBGrid.DBGrid dbgrdJournal 
         Bindings        =   "Journal.frx":0038
         Height          =   3615
         Left            =   -74760
         OleObjectBlob   =   "Journal.frx":0051
         TabIndex        =   35
         Top             =   600
         Width           =   7095
      End
      Begin VB.Frame fraJournal 
         Height          =   2910
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7050
         Begin VB.TextBox txtEntryNo 
            DataField       =   "EntryNo"
            DataSource      =   "datJournal"
            Height          =   300
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   397
            Width           =   1065
         End
         Begin VB.TextBox txtDate 
            DataField       =   "Date"
            DataSource      =   "datJournal"
            Height          =   300
            Left            =   3060
            TabIndex        =   19
            Top             =   390
            Width           =   1080
         End
         Begin VB.TextBox txtAccCode 
            DataField       =   "Acc_Code"
            DataSource      =   "datJournal"
            Height          =   300
            Left            =   5340
            MaxLength       =   4
            TabIndex        =   21
            Top             =   390
            Width           =   1110
         End
         Begin VB.TextBox txtAccDesc 
            Height          =   300
            Left            =   1170
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   1080
            Width           =   5580
         End
         Begin VB.TextBox txtNarration 
            DataField       =   "Narration"
            DataSource      =   "datJournal"
            Height          =   300
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1695
            Width           =   5250
         End
         Begin VB.TextBox txtQuantity 
            Height          =   300
            Left            =   3315
            TabIndex        =   32
            Top             =   2265
            Width           =   1095
         End
         Begin VB.TextBox txtAmount 
            DataField       =   "Amount"
            DataSource      =   "datJournal"
            Height          =   300
            Left            =   5550
            TabIndex        =   34
            Top             =   2265
            Width           =   1230
         End
         Begin VB.TextBox txtNo 
            Height          =   300
            Left            =   1155
            TabIndex        =   17
            Top             =   397
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Frame fraJournal 
            Caption         =   "Transaction Type"
            Height          =   660
            HelpContextID   =   2
            Index           =   2
            Left            =   255
            TabIndex        =   28
            Top             =   2092
            Width           =   1875
            Begin VB.OptionButton optDebit 
               Caption         =   "&Debit"
               Height          =   210
               Left            =   120
               TabIndex        =   29
               Top             =   285
               Width           =   675
            End
            Begin VB.OptionButton optCredit 
               Caption         =   "&Credit"
               Height          =   210
               Left            =   945
               TabIndex        =   30
               Top             =   285
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdSelNarr 
            Caption         =   "..."
            Height          =   300
            Left            =   6480
            TabIndex        =   27
            Top             =   1695
            Width           =   330
         End
         Begin VB.CommandButton cmdSelAccCode 
            Caption         =   "..."
            Height          =   300
            Left            =   6495
            TabIndex        =   22
            Top             =   390
            Width           =   330
         End
         Begin VB.Label lblJournal 
            Caption         =   "Entry No."
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   15
            Top             =   420
            Width           =   840
         End
         Begin VB.Label lblJournal 
            Caption         =   "Date"
            Height          =   255
            Index           =   1
            Left            =   2445
            TabIndex        =   18
            Top             =   420
            Width           =   540
         End
         Begin VB.Label lblJournal 
            Caption         =   "Account Code"
            Height          =   435
            Index           =   2
            Left            =   4470
            TabIndex        =   20
            Top             =   330
            Width           =   750
         End
         Begin VB.Label lblJournal 
            Caption         =   "Account Description"
            Height          =   435
            Index           =   3
            Left            =   210
            TabIndex        =   23
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label lblJournal 
            Caption         =   "Narration"
            Height          =   255
            Index           =   4
            Left            =   210
            TabIndex        =   25
            Top             =   1725
            Width           =   840
         End
         Begin VB.Label lblJournal 
            Caption         =   "Quantity"
            Height          =   255
            Index           =   5
            Left            =   2430
            TabIndex        =   31
            Top             =   2295
            Width           =   660
         End
         Begin VB.Label lblJournal 
            Caption         =   "Amount"
            Height          =   240
            Index           =   6
            Left            =   4740
            TabIndex        =   33
            Top             =   2295
            Width           =   660
         End
      End
      Begin VB.Frame fraJournal 
         Height          =   810
         Index           =   1
         Left            =   255
         TabIndex        =   2
         Top             =   3405
         Width           =   7005
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   270
            Left            =   210
            TabIndex        =   3
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "&Modify"
            Height          =   270
            Left            =   915
            TabIndex        =   4
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "D&elete"
            Height          =   270
            Left            =   1620
            TabIndex        =   5
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   270
            Left            =   2550
            TabIndex        =   6
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Ca&ncel"
            Height          =   270
            Left            =   3960
            TabIndex        =   8
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset"
            Height          =   270
            Left            =   3255
            TabIndex        =   7
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   270
            Left            =   4980
            TabIndex        =   9
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Cl&ose"
            Height          =   270
            Left            =   6030
            TabIndex        =   10
            Top             =   330
            Width           =   705
         End
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Fo&rm"
      Begin VB.Menu mnuClose 
         Caption         =   "Cl&ose"
      End
   End
   Begin VB.Menu mnuNavigation 
      Caption         =   "Na&vigation"
      Begin VB.Menu mnuFirst 
         Caption         =   "&First"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "&Previous"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "frmJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mdbsAccounts As Database
    Dim mrstJournal As Recordset

    Dim mblnTrans As Boolean
    Dim mblnFind As Boolean
    Dim mblnClose As Boolean

    Dim mstrDate As String
    Dim mstrNarr As String
    Dim mstrQty As String
    Dim mstrAmount As String
    
    Dim mvntBookmark As Variant
    Dim mblnQtyToRec As Boolean
    Dim mstrAccCode As String


Private Sub cmdAdd_Click()
    Dim intEntryNo As Integer
    Dim rstCode As Recordset
    
    mblnTrans = True
    Call DisableNaviButtons
    Call Enable_Controls(False)
    Call Lock_Controls(False)
    
    If Not mrstJournal.BOF And Not mrstJournal.EOF Then
        mvntBookmark = mrstJournal.Bookmark
        mrstJournal.MoveLast
        intEntryNo = mrstJournal!Entryno + 1
    Else
        intEntryNo = 1
    End If
    
    mrstJournal.AddNew
    mstrAccCode = ""
    txtAccDesc.Text = ""
    txtQuantity.Text = ""
    txtEntryNo.Text = intEntryNo
    txtDate.Text = mstrDate
    optDebit.Value = False
    optCredit.Value = False
    
    cmdSelAccCode.Enabled = True
    cmdSelNarr.Enabled = True
    sstabJournal.TabEnabled(1) = False
    txtDate.SetFocus
End Sub

Private Sub cmdCancel_Click()
    If mblnTrans Then
        mblnTrans = False
        mblnQtyToRec = False
        mrstJournal.CancelUpdate
        If Not mrstJournal.EOF Then
            mrstJournal.Bookmark = mvntBookmark
        Else
            txtAccDesc.Text = ""
            optDebit.Value = False
            optCredit.Value = False
            txtQuantity.Text = ""
        End If
        
        'txtNarrCode.Visible = False
        'txtNarration.Visible = True
        cmdSelAccCode.Enabled = False
        cmdSelNarr.Enabled = False
        txtQuantity.Enabled = True
    ElseIf mblnFind Then
        mblnFind = False
        cmdFind.Caption = "&Find"
        mrstJournal.Bookmark = mvntBookmark
        txtEntryNo.Visible = True
        txtEntryNo.Visible = True
        txtNo.Visible = False
    End If
    Call Enable_Controls(True)
    Call Lock_Controls(True)
    sstabJournal.TabEnabled(1) = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete the Journal Entry?", vbYesNo) = vbYes Then
        mrstJournal.Delete
        mrstJournal.MoveNext
        If mrstJournal.EOF And mrstJournal.RecordCount > 0 Then
            mrstJournal.MoveLast
        Else
            txtAccDesc.Text = ""
            txtQuantity.Text = ""
            optDebit.Value = False
            optCredit.Value = False
        End If
        Call Enable_Controls(True)
    End If
End Sub

Private Sub cmdFind_Click()
    
    If Not mblnFind Then
        mblnFind = True
        txtEntryNo.Visible = False
        txtNo.Visible = True
        txtNo.Text = ""
        cmdFind.Caption = "&Display"
        
        Call DisableNaviButtons
        Call Enable_Controls(False)
        mvntBookmark = mrstJournal.Bookmark
        sstabJournal.TabEnabled(1) = False
        txtNo.SetFocus
    ElseIf mblnFind Then
        If txtNo.Text = "" Then
            MsgBox "Enter the Entry Number"
            txtNo.SetFocus
        Else
            mblnFind = False
            mrstJournal.FindFirst "EntryNo = " & txtNo.Text
            If mrstJournal.NoMatch Then
               MsgBox "Invalid Entry Number", vbOKOnly
               mrstJournal.Bookmark = mvntBookmark
            End If
            txtEntryNo.Visible = True
            txtNo.Visible = False
            cmdFind.Caption = "&Find"
            sstabJournal.TabEnabled(1) = True
            Call Enable_Controls(True)
        End If
    End If
End Sub

Private Sub cmdFirst_Click()
    mrstJournal.MoveFirst
End Sub

Private Sub cmdLast_Click()
    mrstJournal.MoveLast
End Sub

Private Sub cmdModify_Click()
    Dim rstCode As Recordset
    
    mblnTrans = True
    Call DisableNaviButtons
    Call Enable_Controls(False)
    Call Lock_Controls(False)
    
    mvntBookmark = mrstJournal.Bookmark
    mrstJournal.Edit
    mstrAccCode = txtAccCode.Text
    
    Set rstCode = mdbsAccounts.OpenRecordset("Select Qty_ToRec from AccCode where Acc_Code = '" & mrstJournal!Acc_Code & "'", dbOpenSnapshot)
    If rstCode!Qty_ToRec = "N" Then
        txtQuantity.Enabled = False
        'txtQuantity.Locked = True
        mblnQtyToRec = False
    Else
        mblnQtyToRec = True
    End If
    
    cmdSelAccCode.Enabled = True
    cmdSelNarr.Enabled = True
    sstabJournal.TabEnabled(1) = False
    txtDate.SetFocus
End Sub

Private Sub cmdNext_Click()
    mrstJournal.MoveNext
End Sub

Private Sub cmdPrev_Click()
    mrstJournal.MovePrevious
End Sub


Private Sub cmdReset_Click()
    
    Dim rstCode As Recordset
    If mrstJournal.EditMode = dbEditAdd Then
        txtDate.Text = mstrDate
        txtAccCode.Text = ""
        txtAccDesc.Text = ""
        txtNarration.Text = ""
        txtAmount.Text = ""
        txtQuantity.Text = ""
        optDebit.Value = False
        optCredit.Value = False
        txtQuantity.Locked = False
        mblnQtyToRec = False
        txtQuantity.Enabled = True
        'txtNarrCode.Visible = True
        'txtNarrCode.Text = ""
        'txtNarration.Visible = False
        mstrAccCode = ""
    ElseIf mrstJournal.EditMode = dbEditInProgress Then
        datJournal.UpdateControls
        Set rstCode = mdbsAccounts.OpenRecordset("Select acc_desc, Qty_ToRec from AccCode where acc_code = '" & mrstJournal!Acc_Code & "'", dbOpenSnapshot)
        txtAccDesc = rstCode!Acc_Desc
        If mrstJournal![Debit/Credit] = "D" Then
            optDebit.Value = True
        ElseIf mrstJournal![Debit/Credit] = "C" Then
            optCredit.Value = True
        End If
        If rstCode!Qty_ToRec = "N" Then
            txtQuantity.Text = ""
            txtQuantity.Enabled = False
            'txtQuantity.Locked = True
            mblnQtyToRec = False
        ElseIf rstCode!Qty_ToRec = "Y" Then
            txtQuantity.Text = mrstJournal!Quantity
            txtQuantity.Enabled = True
            'txtQuantity.Locked = False
            mblnQtyToRec = True
        End If
        rstCode.Close
        mstrAccCode = txtAccCode.Text
    End If
    txtDate.SetFocus
End Sub

Private Sub cmdSave_Click()
    Dim blnRetVal As Boolean
     
    blnRetVal = FieldCheck
    If blnRetVal Then
        mrstJournal!BillNo = 0
        If mrstJournal.EditMode = dbEditAdd Then
            'mrstJournal![Cash/Journal] = "J"
            mstrDate = txtDate.Text
        End If
        mstrNarr = txtNarration.Text
        If optCredit.Value Then
            mrstJournal![Debit/Credit] = "C"
        ElseIf optDebit.Value Then
            mrstJournal![Debit/Credit] = "D"
        End If
        If txtQuantity.Text <> "" Then
            mrstJournal!Quantity = Val(txtQuantity.Text)
            mstrQty = txtQuantity.Text
        Else
            If Not mblnQtyToRec Then
                mrstJournal!Quantity = Null
                mstrQty = ""
            Else
                mrstJournal!Quantity = 0
                mstrQty = 0
            End If
        End If
        mrstJournal!Amount = Format(txtAmount.Text, "#,##,##0.00")
        mstrAmount = txtAmount.Text
        mrstJournal.Update
        mblnQtyToRec = False
        mblnTrans = False
        mrstJournal.Bookmark = mrstJournal.LastModified
        
        Call Enable_Controls(True)
        Call Lock_Controls(True)
        
        cmdSelAccCode.Enabled = False
        cmdSelNarr.Enabled = False
        txtQuantity.Enabled = True
        sstabJournal.TabEnabled(1) = True
    End If
End Sub

Private Sub cmdSelAccCode_Click()
    gstrFormName = "Journal"
    gstrDetName = "Account"
    frmDetails.Show
End Sub

Private Sub cmdSelNarr_Click()
    gstrFormName = "Journal"
    gstrDetName = "Narration"
    frmDetails.Show
End Sub


Private Sub datJournal_Reposition()
    Dim rstCode As Recordset
    If Not mblnTrans And Not mblnFind Then
        If txtAccCode.Text <> "" And Not mrstJournal.EOF Then
            Set rstCode = mdbsAccounts.OpenRecordset("Select acc_desc from AccCode where acc_code = '" & txtAccCode.Text & "'", dbOpenSnapshot)
            txtAccDesc.Text = rstCode!Acc_Desc
            rstCode.Close
            If mrstJournal![Debit/Credit] = "C" Then
                optCredit.Value = True
            ElseIf mrstJournal![Debit/Credit] = "D" Then
                optDebit.Value = True
            End If
            If Not IsNull(mrstJournal!Quantity) Then
                txtQuantity.Text = mrstJournal!Quantity
            Else
                txtQuantity.Text = ""
            End If
            txtAmount.Text = Format(mrstJournal!Amount, "#####.00")
        Else
            txtAccDesc.Text = ""
            txtQuantity.Text = ""
            optCredit.Value = False
            optDebit.Value = False
        End If
        Call RecordPosition
    End If

End Sub

Private Sub Form_Load()
    
    Dim strCondition As String
    Dim rstJournalTemp As Recordset
    Dim rstCashTemp As Recordset
    
    Dim strJourDate As String
    Dim strCashDate As String
    
    mblnTrans = False
    mblnFind = False
    mblnClose = False
    mblnQtyToRec = False
    
    Set mdbsAccounts = OpenDatabase(gstrDatabase)
        
    strCondition = "Select * from Journal Order By EntryNo"
    Set mrstJournal = mdbsAccounts.OpenRecordset(strCondition, dbOpenDynaset)
    Set datJournal.Recordset = mrstJournal
    
    If Not mrstJournal.EOF And Not mrstJournal.BOF Then
        mrstJournal.MoveLast
        If Not IsNull(mrstJournal!Quantity) Then
            mstrQty = mrstJournal!Quantity
        Else
            mstrQty = ""
        End If
        If Not IsNull(mrstJournal!Narration) Then
            mstrNarr = mrstJournal!Narration
        Else
            mstrNarr = ""
        End If
        mstrAmount = mrstJournal!Amount
        mrstJournal.MoveFirst
    Else
        mstrNarr = ""
        mstrQty = ""
        mstrAmount = ""
    End If
    
    Set rstJournalTemp = mdbsAccounts.OpenRecordset("Select Date from Journal Order by Date", dbOpenSnapshot)
    If Not rstJournalTemp.EOF Then
        rstJournalTemp.MoveLast
        strJourDate = rstJournalTemp!Date
    Else
        strJourDate = ""
    End If
    rstJournalTemp.Close

    Set rstCashTemp = mdbsAccounts.OpenRecordset("Select Date from Cash Order by Date", dbOpenSnapshot)
    If Not rstCashTemp.EOF Then
        rstCashTemp.MoveLast
        strCashDate = rstCashTemp!Date
    Else
        strCashDate = ""
    End If
    rstCashTemp.Close

    If strJourDate = "" And strCashDate = "" Then
        mstrDate = ""
    ElseIf strJourDate = "" Then
        mstrDate = strCashDate
    ElseIf strCashDate = "" Then
        mstrDate = strJourDate
    Else
        If strJourDate > strCashDate Then
            mstrDate = strJourDate
        Else
            mstrDate = strCashDate
        End If
    End If
    
    Call RecordPosition
    Call Lock_Controls(True)
    Call Enable_Controls(True)
    
    cmdSelAccCode.Enabled = False
    cmdSelNarr.Enabled = False
    'txtNarrCode.Visible = False
    
End Sub





Private Sub RecordPosition()
    If mrstJournal.RecordCount = 0 Or mrstJournal.RecordCount = 1 Then
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdLast.Enabled = False
        
        mnuFirst.Enabled = False
        mnuPrev.Enabled = False
        mnuNext.Enabled = False
        mnuLast.Enabled = False
    ElseIf mrstJournal.AbsolutePosition = 0 Then
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        
        mnuFirst.Enabled = False
        mnuPrev.Enabled = False
        mnuNext.Enabled = True
        mnuLast.Enabled = True
    ElseIf mrstJournal.AbsolutePosition = (mrstJournal.RecordCount - 1) Then
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
        cmdLast.Enabled = False
        
        mnuFirst.Enabled = True
        mnuPrev.Enabled = True
        mnuNext.Enabled = False
        mnuLast.Enabled = False
    Else
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        
        mnuFirst.Enabled = True
        mnuPrev.Enabled = True
        mnuNext.Enabled = True
        mnuLast.Enabled = True
    End If
End Sub

Private Sub Lock_Controls(blnLock As Boolean)
    txtDate.Locked = blnLock
    txtAccCode.Locked = blnLock
    txtNarration.Locked = blnLock
    txtQuantity.Locked = blnLock
    txtAmount.Locked = blnLock
    
    optDebit.Enabled = Not blnLock
    optCredit.Enabled = Not blnLock
End Sub

Private Sub Enable_Controls(blnEnable As Boolean)
    cmdAdd.Enabled = blnEnable
    If blnEnable And mrstJournal.RecordCount <> 0 Then
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
    If (blnEnable And mrstJournal.RecordCount > 1) Or mblnFind Then
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
    mnuClose.Enabled = blnEnable
End Sub

Private Sub DisableNaviButtons()
    cmdFirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
    
    mnuFirst.Enabled = False
    mnuPrev.Enabled = False
    mnuNext.Enabled = False
    mnuLast.Enabled = False
End Sub

Private Sub mnuClose_Click()
    cmdClose.Value = True
End Sub

Private Sub mnuFirst_Click()
    cmdFirst.Value = True
End Sub

Private Sub mnuLast_Click()
    cmdLast.Value = True
End Sub


Private Sub mnuNext_Click()
    cmdNext.Value = True
End Sub




Private Sub mnuPrev_Click()
    cmdPrev.Value = True
End Sub

Private Sub txtAccCode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And mblnTrans Then
        If (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            Beep
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Len(txtAccCode.Text) = 0 Then
                If KeyAscii <> Asc("A") And KeyAscii <> Asc("L") And _
                    KeyAscii <> Asc("E") And KeyAscii <> Asc("P") And _
                    KeyAscii <> Asc("F") And KeyAscii <> Asc("I") Then
                    MsgBox "First letter must begin with A,L,E,P,F,I"
                    KeyAscii = 0
                End If
            End If
        End If
    End If
End Sub


Private Sub txtAccCode_LostFocus()
    
    Dim rstCode As Recordset
    If mblnTrans Then
        If txtAccCode.Text <> "" And txtAccCode.Text <> mstrAccCode Then
            Set rstCode = mdbsAccounts.OpenRecordset("Select Acc_Desc,Qty_ToRec from AccCode where Acc_Code = '" & txtAccCode.Text & "'", dbOpenSnapshot)
            If rstCode.EOF Then
                MsgBox "Invalid Account Code"
                txtAccDesc.Text = ""
                mstrAccCode = ""
                txtAccCode.SetFocus
            Else
                txtAccDesc.Text = rstCode!Acc_Desc
                If rstCode!Qty_ToRec = "N" Then
                    txtQuantity.Text = ""
                    txtQuantity.Enabled = False
                    'txtQuantity.Locked = True
                    mblnQtyToRec = False
                ElseIf rstCode!Qty_ToRec = "Y" Then
                    txtQuantity.Enabled = True
                    'txtQuantity.Locked = False
                    mblnQtyToRec = True
                    'If Left$(txtAccCode.Text, 1) = "P" Then 'And mrstJournal.EditMode = dbEditAdd Then
                    '    txtQuantity.Text = "-"
                    'End If
                End If
                mstrAccCode = txtAccCode.Text
            End If
        ElseIf txtAccCode.Text = "" Then
            mstrAccCode = ""
            txtAccDesc.Text = ""
        End If
    End If
End Sub


Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        txtAmount.Text = mstrAmount
    End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)

    If mblnTrans Then
        If (KeyAscii <> 8) And (KeyAscii <> Asc(".")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            KeyAscii = 0
        End If
    End If

End Sub


Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If mblnTrans Then
        If (KeyAscii <> 8) And (KeyAscii <> Asc("/")) And _
        (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
           Beep
           KeyAscii = 0
        End If
    End If
End Sub


Private Sub txtNarration_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        txtNarration.Text = mstrNarr
    End If
End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
    If mblnTrans Then
        If (KeyAscii <> 8) And (KeyAscii <> Asc(",")) And _
            (KeyAscii <> Asc(".")) And (KeyAscii <> Asc("-")) And _
            (KeyAscii <> Asc(" ")) And (KeyAscii <> Asc("&")) And _
            (KeyAscii <> Asc("(")) And (KeyAscii <> Asc(")")) And _
            (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> 13) Then
            Beep
            KeyAscii = 0
        ElseIf KeyAscii = 13 Then
            Dim rstCode As Recordset
            Set rstCode = mdbsAccounts.OpenRecordset("Select Narr_Desc from NarrCode where Narr_Code  = '" & txtNarration.Text & " '", dbOpenSnapshot)
            If rstCode.EOF Then
                MsgBox "Invalid Narration"
                txtNarration.SetFocus
                'txtNarrCode.SetFocus
            Else
                'txtNarrCode.Visible = False
                'txtNarration.Visible = True
                txtNarration.Text = rstCode!Narr_Desc
                txtNarration.SetFocus
            End If
        End If
    End If
End Sub


Private Sub txtNarration_LostFocus()
    
'    Dim rstcode As Recordset
'    If mblnTrans And txtNarration.Tag = "Find" Then
'        If txtNarration.Text <> "" Then
'            Set rstcode = mdbsAccounts.OpenRecordset("Select Narr_Desc from NarrCode where Narr_Code = '" & txtNarration.Text & "'", dbOpenSnapshot)
'            If rstcode.EOF Then
'                MsgBox "Invalid Narration Code"
'                txtNarration.SetFocus
'            Else
'                txtNarration.Text = rstcode!Narr_Desc
'            End If
'        End If
'    End If
End Sub





Private Sub txtNo_KeyPress(KeyAscii As Integer)
    
    If mblnFind Then
        If (KeyAscii <> 8) And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub txtQuantity_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        txtQuantity.Text = mstrQty
    End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)

    If mblnTrans Then
        If (KeyAscii <> 8) And (KeyAscii <> Asc(".")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> Asc("-")) Then
            KeyAscii = 0
        End If
    End If

End Sub



Private Function FieldCheck() As Boolean

    Dim rstCode As Recordset
    FieldCheck = False
    If txtDate.Text = "" Then
        MsgBox "Date should not blank"
        txtDate.SetFocus
    'elseif
    ElseIf txtAccCode.Text = "" Then
        MsgBox "Account Code should not be blank"
        txtAccDesc.Text = ""
        mstrAccCode = ""
        txtAccCode.SetFocus
    Else
        Set rstCode = mdbsAccounts.OpenRecordset( _
            "Select Acc_Code from AccCode where Acc_Code = '" _
            & txtAccCode.Text & "'", dbOpenSnapshot)
        If rstCode.EOF Then
            MsgBox "Invalid AccountCode"
            txtAccDesc.Text = ""
            mstrAccCode = ""
            txtAccCode.SetFocus
        ElseIf Not optDebit.Value And Not optCredit.Value Then
            MsgBox "Specify the type of transaction"
            optDebit.SetFocus
        ElseIf txtQuantity.Text <> "" Then
            If Not CheckingNum(txtQuantity.Text) Then
                MsgBox "Invalid Quantity"
                txtQuantity.SetFocus
            Else
                FieldCheck = True
            End If
        Else
            FieldCheck = True
        End If
        If FieldCheck Then
            If txtAmount.Text = "" Then
                MsgBox "Amount Should not be blank"
                txtAmount.SetFocus
                FieldCheck = False
            ElseIf Not CheckingNum(txtAmount.Text) Then
                MsgBox "Invalid Amount"
                FieldCheck = False
                txtAmount.SetFocus
            End If
        End If
    End If
        
End Function

Private Function CheckingNum(strText As String) As Boolean
    
    Dim strNum As String
    Dim intPos As Integer
    
    CheckingNum = False
    If IsNumeric(strText) Then
        intPos = InStr(1, strText, ".")
        If intPos <> 0 Then
            strNum = Right$(strText, Len(strText) - intPos)
            If Len(strNum) <= 2 Then
                CheckingNum = True
            End If
        Else
            CheckingNum = True
        End If
    End If
End Function


