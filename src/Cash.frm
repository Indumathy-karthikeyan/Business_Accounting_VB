VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Transaction"
   ClientHeight    =   5445
   ClientLeft      =   2160
   ClientTop       =   1545
   ClientWidth     =   8085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8085
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data datCash 
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
      Width           =   1320
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   285
      Left            =   7320
      TabIndex        =   14
      Top             =   255
      Width           =   465
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   285
      Left            =   6855
      TabIndex        =   13
      Top             =   255
      Width           =   465
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   285
      Left            =   6390
      TabIndex        =   12
      Top             =   255
      Width           =   465
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   285
      Left            =   5925
      TabIndex        =   11
      Top             =   255
      Width           =   465
   End
   Begin TabDlg.SSTab sstabCash 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Form View"
      TabPicture(0)   =   "Cash.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraCash(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCash(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "SpreedSheet View"
      TabPicture(1)   =   "Cash.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbgrdCash"
      Tab(1).ControlCount=   1
      Begin MSDBGrid.DBGrid dbgrdCash 
         Bindings        =   "Cash.frx":0038
         Height          =   4095
         Left            =   -74760
         OleObjectBlob   =   "Cash.frx":004E
         TabIndex        =   35
         Top             =   600
         Width           =   7095
      End
      Begin VB.Frame fraCash 
         Height          =   3315
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   7050
         Begin VB.Frame fraCash 
            Caption         =   "Transaction Type"
            Height          =   990
            Index           =   2
            Left            =   240
            TabIndex        =   28
            Top             =   2130
            Width           =   2025
            Begin VB.OptionButton optCashPay 
               Caption         =   "Cash &Payment"
               Height          =   300
               Left            =   330
               TabIndex        =   30
               Top             =   570
               Width           =   1350
            End
            Begin VB.OptionButton optCashRec 
               Caption         =   "Cash R&eceipt"
               Height          =   315
               Left            =   330
               TabIndex        =   29
               Top             =   225
               Width           =   1335
            End
         End
         Begin VB.CommandButton cmdSelNarr 
            Caption         =   "..."
            Height          =   300
            Left            =   6465
            TabIndex        =   27
            Top             =   1695
            Width           =   285
         End
         Begin VB.CommandButton cmdSelAccCode 
            Caption         =   "..."
            Height          =   300
            Left            =   6465
            TabIndex        =   22
            Top             =   390
            Width           =   285
         End
         Begin VB.TextBox txtNo 
            Height          =   300
            Left            =   1170
            TabIndex        =   16
            Top             =   397
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtAmount 
            DataField       =   "Amount"
            DataSource      =   "datCash"
            Height          =   300
            Left            =   5490
            TabIndex        =   34
            Top             =   2475
            Width           =   1260
         End
         Begin VB.TextBox txtQuantity 
            Height          =   300
            Left            =   3225
            TabIndex        =   32
            Top             =   2475
            Width           =   1245
         End
         Begin VB.TextBox txtNarration 
            DataField       =   "Narration"
            DataSource      =   "datCash"
            Height          =   300
            Left            =   1170
            MaxLength       =   50
            TabIndex        =   26
            Top             =   1695
            Width           =   5190
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
         Begin VB.TextBox txtAccCode 
            DataField       =   "Acc_Code"
            DataSource      =   "datCash"
            Height          =   300
            Left            =   5250
            MaxLength       =   4
            TabIndex        =   21
            Top             =   375
            Width           =   1110
         End
         Begin VB.TextBox txtDate 
            DataField       =   "Date"
            DataSource      =   "datCash"
            Height          =   300
            Left            =   3030
            TabIndex        =   19
            Top             =   375
            Width           =   1020
         End
         Begin VB.TextBox txtEntryNo 
            DataField       =   "EntryNo"
            DataSource      =   "datCash"
            Height          =   300
            Left            =   1155
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   397
            Width           =   990
         End
         Begin VB.Label lblCash 
            Caption         =   "Amount"
            Height          =   255
            Index           =   6
            Left            =   4755
            TabIndex        =   33
            Top             =   2505
            Width           =   630
         End
         Begin VB.Label lblCash 
            Caption         =   "Quantity"
            Height          =   255
            Index           =   5
            Left            =   2475
            TabIndex        =   31
            Top             =   2505
            Width           =   705
         End
         Begin VB.Label lblCash 
            Caption         =   "Narration"
            Height          =   255
            Index           =   4
            Left            =   210
            TabIndex        =   25
            Top             =   1725
            Width           =   840
         End
         Begin VB.Label lblCash 
            Caption         =   "Account Description"
            Height          =   435
            Index           =   3
            Left            =   210
            TabIndex        =   23
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label lblCash 
            Caption         =   "Account Code"
            Height          =   435
            Index           =   2
            Left            =   4320
            TabIndex        =   20
            Top             =   330
            Width           =   750
         End
         Begin VB.Label lblCash 
            Caption         =   "Date"
            Height          =   255
            Index           =   1
            Left            =   2385
            TabIndex        =   18
            Top             =   420
            Width           =   540
         End
         Begin VB.Label lblCash 
            Caption         =   "Entry No."
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   15
            Top             =   420
            Width           =   840
         End
      End
      Begin VB.Frame fraCash 
         Height          =   810
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   3810
         Width           =   7050
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Height          =   270
            Left            =   6045
            TabIndex        =   10
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   270
            Left            =   5025
            TabIndex        =   9
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset"
            Height          =   270
            Left            =   3330
            TabIndex        =   7
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Ca&ncel"
            Height          =   270
            Left            =   4035
            TabIndex        =   8
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   270
            Left            =   2625
            TabIndex        =   6
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   270
            Left            =   1650
            TabIndex        =   5
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "&Modify"
            Height          =   270
            Left            =   945
            TabIndex        =   4
            Top             =   330
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   270
            Left            =   240
            TabIndex        =   3
            Top             =   330
            Width           =   705
         End
      End
   End
End
Attribute VB_Name = "frmCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mdbsAccounts As Database
    Dim mrstCash As Recordset

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
    
    Dim rstCode As Recordset
    Dim strEntryNo As String
    
    mblnTrans = True
    Call DisableNaviButtons
    Call Enable_Controls(False)
    Call Lock_Controls(False)
    
    If Not mrstCash.BOF And Not mrstCash.EOF Then
        mvntBookmark = mrstCash.Bookmark
        mrstCash.MoveLast
        strEntryNo = mrstCash!Entryno + 1
    Else
        strEntryNo = 1
    End If
    
    mrstCash.AddNew
    txtAccDesc.Text = ""
    txtQuantity.Text = ""
    txtDate.Text = mstrDate
    txtEntryNo = strEntryNo
    mstrAccCode = ""
    optCashRec.Value = False
    optCashPay.Value = False
        
    cmdSelAccCode.Enabled = True
    cmdSelNarr.Enabled = True
    sstabCash.TabEnabled(1) = False
    txtDate.SetFocus
End Sub

Private Sub cmdCancel_Click()
    If mblnTrans Then
        mblnTrans = False
        mblnQtyToRec = False
        mrstCash.CancelUpdate
        If Not mrstCash.EOF And Not mrstCash.BOF Then
            mrstCash.Bookmark = mvntBookmark
        Else
            txtAccDesc.Text = ""
            txtQuantity.Text = ""
            optCashRec.Value = False
            optCashPay.Value = False
        End If
        
        cmdSelAccCode.Enabled = False
        cmdSelNarr.Enabled = False
        txtQuantity.Enabled = True
    ElseIf mblnFind Then
        mblnFind = False
        cmdFind.Caption = "&Find"
        mrstCash.Bookmark = mvntBookmark
        
        txtEntryNo.Visible = True
        txtNo.Visible = False
    End If
    Call Enable_Controls(True)
    Call Lock_Controls(True)
    sstabCash.TabEnabled(1) = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Delete the Entry?", vbYesNo) = vbYes Then
        mrstCash.Delete
        mrstCash.MoveNext
        If mrstCash.EOF And mrstCash.RecordCount <> 0 Then
            mrstCash.MoveLast
        End If
        Call Enable_Controls(True)
        If mrstCash.EOF And mrstCash.BOF Then
            txtAccDesc.Text = ""
            txtQuantity.Text = ""
            optCashRec.Value = False
            optCashPay.Value = False
        End If
    End If
End Sub

Private Sub cmdFind_Click()
    
    If Not mblnFind Then
        mblnFind = True
        txtEntryNo.Visible = False
        txtNo.Visible = True
        txtNo.Text = ""
        cmdFind.Caption = "&Display"
        
        mvntBookmark = mrstCash.Bookmark
        Call DisableNaviButtons
        Call Enable_Controls(False)
        sstabCash.TabEnabled(1) = False
        txtNo.SetFocus
    ElseIf mblnFind Then
        If txtNo.Text = "" Then
            MsgBox "Enter the Enter No."
            txtNo.SetFocus
        Else
            mblnFind = False
            mrstCash.FindFirst "entryno = " & txtNo.Text
            If mrstCash.NoMatch Then
                MsgBox "Invalid Entry Number", vbOKOnly
                mrstCash.Bookmark = mvntBookmark
            End If
            txtEntryNo.Visible = True
            txtNo.Visible = False
            cmdFind.Caption = "&Find"
        
            Call Enable_Controls(True)
            sstabCash.TabEnabled(1) = True
        End If
    End If
End Sub

Private Sub cmdFirst_Click()
    mrstCash.MoveFirst
End Sub

Private Sub cmdLast_Click()
    mrstCash.MoveLast
End Sub

Private Sub cmdModify_Click()

    Dim rstCode As Recordset
    
    mblnTrans = True
    Call DisableNaviButtons
    Call Enable_Controls(False)
    Call Lock_Controls(False)
    
    mvntBookmark = mrstCash.Bookmark
    mrstCash.Edit
    mstrAccCode = txtAccCode.Text
    
    Set rstCode = mdbsAccounts.OpenRecordset("Select Qty_ToRec from AccCode where Acc_Code = '" & mrstCash!Acc_Code & "'", dbOpenSnapshot)
    If rstCode!Qty_ToRec = "N" Then
        txtQuantity.Enabled = False
        mblnQtyToRec = False
    Else
        txtQuantity.Enabled = True
        mblnQtyToRec = True
    End If
    
    cmdSelAccCode.Enabled = True
    cmdSelNarr.Enabled = True
    sstabCash.TabEnabled(1) = False
    txtDate.SetFocus
End Sub

Private Sub cmdNext_Click()
    mrstCash.MoveNext
End Sub

Private Sub cmdPrev_Click()
    mrstCash.MovePrevious
End Sub


Private Sub cmdReset_Click()
    
    Dim rstCode As Recordset
    If mrstCash.EditMode = dbEditAdd Then
        txtDate.Text = mstrDate
        txtAccCode.Text = ""
        txtAccDesc.Text = ""
        txtNarration.Text = ""
        txtAmount.Text = ""
        txtQuantity.Text = ""
        mblnQtyToRec = False
        mstrAccCode = ""
        txtQuantity.Enabled = True
        optCashRec.Value = False
        optCashPay.Value = False
    ElseIf mrstCash.EditMode = dbEditInProgress Then
        datCash.UpdateControls
        Set rstCode = mdbsAccounts.OpenRecordset("Select Acc_Desc, Qty_ToRec from AccCode where acc_code = '" & txtAccCode.Text & "'", dbOpenSnapshot)
        txtAccDesc = rstCode!Acc_Desc
        If rstCode!Qty_ToRec = "N" Then
            txtQuantity.Text = ""
            txtQuantity.Enabled = False
            mblnQtyToRec = False
        ElseIf rstCode!Qty_ToRec = "Y" Then
            txtQuantity.Text = mrstCash!Quantity
            txtQuantity.Enabled = True
            mblnQtyToRec = True
        End If
        rstCode.Close
        mstrAccCode = txtAccCode.Text
        If mrstCash![Debit/Credit] = "C" Then
            optCashRec.Value = True
        ElseIf mrstCash![Debit/Credit] = "D" Then
            optCashPay.Value = True
        End If
    End If
    txtDate.SetFocus
End Sub

Private Sub cmdSave_Click()
    
    If FieldCheck Then
        mrstCash!BillNo = 0
        If mrstCash.EditMode = dbEditAdd Then
            'mrstCash![Cash/Journal] = "C"
            mstrDate = txtDate.Text
        End If
        mstrNarr = txtNarration.Text
        If optCashRec.Value Then
            mrstCash![Debit/Credit] = "C"
        ElseIf optCashPay.Value Then
            mrstCash![Debit/Credit] = "D"
        End If
        If txtQuantity.Text <> "" And txtQuantity.Text <> "-" Then
            mrstCash!Quantity = Val(txtQuantity.Text)
            mstrQty = txtQuantity.Text
        Else
            If mblnQtyToRec Then
                mrstCash!Quantity = 0
                mstrQty = 0
            Else
                mrstCash!Quantity = Null
                mstrQty = ""
            End If
        End If
        mrstCash!Amount = Format(txtAmount.Text, "#,##,##0.00")
        mstrAmount = txtAmount.Text
        mrstCash.Update
        mblnQtyToRec = False
        mblnTrans = False
        mrstCash.Bookmark = mrstCash.LastModified
        
        Call Enable_Controls(True)
        Call Lock_Controls(True)
        
        cmdSelAccCode.Enabled = False
        cmdSelNarr.Enabled = False
        txtQuantity.Enabled = True
        
        sstabCash.TabEnabled(1) = True
    End If
End Sub

Private Sub cmdSelAccCode_Click()
    gstrFormName = "Cash"
    gstrDetName = "Account"
    frmDetails.Show
End Sub

Private Sub cmdSelNarr_Click()
    gstrFormName = "Cash"
    gstrDetName = "Narration"
    frmDetails.Show
End Sub

Private Sub datCash_Reposition()
    Dim rstCode As Recordset
    If Not mblnTrans And Not mblnFind Then
        If txtAccCode.Text <> "" Then
            Set rstCode = mdbsAccounts.OpenRecordset("Select Acc_Desc from AccCode where acc_code = '" & txtAccCode.Text & "'", dbOpenSnapshot)
            If Not rstCode.EOF Then
                txtAccDesc = rstCode!Acc_Desc
            End If
            rstCode.Close
            If Not IsNull(mrstCash!Quantity) Then
                txtQuantity.Text = mrstCash!Quantity
            Else
                txtQuantity.Text = ""
            End If
            txtAmount.Text = Format(mrstCash!Amount, "#####.00")
            If mrstCash![Debit/Credit] = "C" Then
                optCashRec.Value = True
            ElseIf mrstCash![Debit/Credit] = "D" Then
                optCashPay.Value = True
            End If
        Else
            txtAccDesc.Text = ""
            txtQuantity.Text = ""
            optCashRec.Value = False
            optCashPay.Value = False
        End If
        Call RecordPosition
    End If
End Sub

Private Sub Form_Load()

    Dim strCondition As String
    Dim rstCashTemp As Recordset
    Dim rstJourTemp As Recordset
    
    Dim strJourDate As String
    Dim strCashDate As String
    
    mblnTrans = False
    mblnFind = False
    mblnClose = False
    mblnQtyToRec = False
    
    strCondition = "Select * from Cash Order By EntryNo "
    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    Set mrstCash = mdbsAccounts.OpenRecordset(strCondition, dbOpenDynaset)
    Set datCash.Recordset = mrstCash
    
    If Not mrstCash.EOF And Not mrstCash.BOF Then
        mrstCash.MoveLast
        If Not IsNull(mrstCash!Narration) Then
            mstrNarr = mrstCash!Narration
        Else
            mstrNarr = ""
        End If
        If Not IsNull(mrstCash!Quantity) Then
            mstrQty = mrstCash!Quantity
        Else
            mstrQty = ""
        End If
        mstrAmount = mrstCash!Amount
        mrstCash.MoveFirst
    Else
        mstrNarr = ""
        mstrQty = ""
        mstrAmount = ""
    End If
    Call RecordPosition
    Call Lock_Controls(True)
    Call Enable_Controls(True)
    
    Set rstCashTemp = mdbsAccounts.OpenRecordset("Select Date from Cash Order By Date", dbOpenSnapshot)
    If Not rstCashTemp.EOF And Not rstCashTemp.BOF Then
        rstCashTemp.MoveLast
        strCashDate = rstCashTemp!Date
    Else
        strCashDate = ""
    End If
    rstCashTemp.Close
    
    Set rstJourTemp = mdbsAccounts.OpenRecordset("Select Date from Journal Order By Date", dbOpenSnapshot)
    If Not rstJourTemp.EOF And Not rstJourTemp.BOF Then
        rstJourTemp.MoveLast
        strJourDate = rstJourTemp!Date
    Else
        strJourDate = ""
    End If
    rstJourTemp.Close

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
    cmdSelAccCode.Enabled = False
    cmdSelNarr.Enabled = False
End Sub

Private Function FieldCheck() As Boolean

    Dim rstCode As Recordset
    FieldCheck = False
    If txtDate.Text = "" Then
        MsgBox "Date should not blank"
        txtDate.SetFocus
    ElseIf Not IsDate(txtDate.Text) Then
        MsgBox "Invalid Date"
        txtDate.SetFocus
    ElseIf txtAccCode.Text = "" Then
         MsgBox "Account Code should not be blank"
         txtAccDesc.Text = ""
         mstrAccCode = ""
         txtAccCode.SetFocus
    Else
        Set rstCode = mdbsAccounts.OpenRecordset("Select Acc_Code from AccCode where Acc_Code = '" & txtAccCode.Text & "'", dbOpenSnapshot)
        If rstCode.EOF Then
            MsgBox "Invalid Account Code"
            txtAccDesc.Text = ""
            mstrAccCode = ""
            txtAccCode.SetFocus
        ElseIf Not optCashRec.Value And Not optCashPay.Value Then
            MsgBox "Specify the Type of Transaction"
            optCashRec.SetFocus
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




Private Sub RecordPosition()
    If mrstCash.RecordCount = 0 Or mrstCash.RecordCount = 1 Then
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdLast.Enabled = False
    ElseIf mrstCash.AbsolutePosition = 0 Then
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    ElseIf mrstCash.AbsolutePosition = (mrstCash.RecordCount - 1) Then
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
        cmdLast.Enabled = False
    Else
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    End If
End Sub

Private Sub Lock_Controls(blnLock As Boolean)
    txtDate.Locked = blnLock
    txtAccCode.Locked = blnLock
    txtNarration.Locked = blnLock
    txtQuantity.Locked = blnLock
    txtAmount.Locked = blnLock
    
    optCashRec.Enabled = Not blnLock
    optCashPay.Enabled = Not blnLock
End Sub

Private Sub Enable_Controls(blnEnable As Boolean)
    cmdAdd.Enabled = blnEnable
    If blnEnable And mrstCash.RecordCount <> 0 Then
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
    If (blnEnable And mrstCash.RecordCount > 1) Or mblnFind Then
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

Private Sub DisableNaviButtons()
    cmdFirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
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
'            txtAccCode.Tag = "Find"
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
                    mblnQtyToRec = False
                ElseIf rstCode!Qty_ToRec = "Y" Then
                    txtQuantity.Enabled = True
                    mblnQtyToRec = True
                    'If Left$(txtAccCode.Text, 1) = "P" Then
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
            Beep
            KeyAscii = 0
        End If
    End If

End Sub


Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If mblnTrans Then
        If (KeyAscii <> 8) And (KeyAscii <> Asc("/")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            KeyAscii = 0
            Beep
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
            Else
                txtNarration.Text = rstCode!Narr_Desc
                txtNarration.SetFocus
            End If
            rstCode.Close
        End If
    End If

End Sub


Private Sub txtNarration_LostFocus()
    
'    Dim rstCode As Recordset
'    If mblnTrans Then 'And txtNarration.Tag = "Find" Then
'        If txtNarration.Text <> "" Then
'            Set rstCode = mdbsAccounts.OpenRecordset("Select Narr_Desc from NarrCode where Narr_Code = '" & txtNarration.Text & "'", dbOpenSnapshot)
'            If rstCode.EOF Then
'                MsgBox "Invalid Narration Code"
'                txtNarration.SetFocus
'            Else
'                txtNarration.Text = rstCode!Narr_Desc
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
            Beep
        End If
    End If

End Sub





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


