VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAccDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Details"
   ClientHeight    =   5580
   ClientLeft      =   1950
   ClientTop       =   1260
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   285
      Left            =   7245
      TabIndex        =   30
      Top             =   255
      Width           =   510
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   285
      Left            =   6720
      TabIndex        =   29
      Top             =   255
      Width           =   510
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   285
      Left            =   6195
      TabIndex        =   28
      Top             =   255
      Width           =   510
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      Height          =   285
      Left            =   5670
      TabIndex        =   27
      Top             =   255
      Width           =   510
   End
   Begin VB.Data datAccDet 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\Accounts\YR2002\Sntc.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1230
   End
   Begin TabDlg.SSTab sstabAccDet 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Form View"
      TabPicture(0)   =   "AccDet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAccDet(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraAccDet(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "SpreedSheet View"
      TabPicture(1)   =   "AccDet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbgrdAccDet"
      Tab(1).ControlCount=   1
      Begin MSDBGrid.DBGrid dbgrdAccDet 
         Bindings        =   "AccDet.frx":0038
         Height          =   4215
         Left            =   -74760
         OleObjectBlob   =   "AccDet.frx":0050
         TabIndex        =   26
         Top             =   600
         Width           =   6975
      End
      Begin VB.Frame fraAccDet 
         Height          =   720
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Top             =   3960
         Width           =   6855
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Canc&el"
            Height          =   285
            Left            =   3960
            TabIndex        =   25
            Top             =   285
            Width           =   705
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Cl&ose"
            Height          =   285
            Left            =   5910
            TabIndex        =   24
            Top             =   285
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   285
            Left            =   4935
            TabIndex        =   23
            Top             =   285
            Width           =   705
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset"
            Height          =   285
            Left            =   3255
            TabIndex        =   22
            Top             =   285
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   285
            Left            =   2550
            TabIndex        =   21
            Top             =   285
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "De&lete"
            Height          =   285
            Left            =   1605
            TabIndex        =   20
            Top             =   285
            Width           =   705
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "&Modify"
            Height          =   285
            Left            =   900
            TabIndex        =   19
            Top             =   285
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   285
            Left            =   195
            TabIndex        =   18
            Top             =   285
            Width           =   705
         End
      End
      Begin VB.Frame fraAccDet 
         Height          =   3450
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   6855
         Begin VB.TextBox txtAccCode 
            DataField       =   "Acc_Code"
            DataSource      =   "datAccDet"
            Height          =   285
            Left            =   1530
            MaxLength       =   4
            TabIndex        =   32
            Top             =   375
            Width           =   1350
         End
         Begin VB.ComboBox cboAccCode 
            Height          =   315
            Left            =   1530
            TabIndex        =   31
            Text            =   "Combo1"
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtQtyUnit 
            DataField       =   "Qty_Unit"
            DataSource      =   "datAccDet"
            Height          =   285
            Left            =   4860
            MaxLength       =   4
            TabIndex        =   11
            Top             =   1628
            Width           =   1620
         End
         Begin VB.TextBox txtAccDesc 
            DataField       =   "Acc_Desc"
            DataSource      =   "datAccDet"
            Height          =   285
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   10
            Top             =   945
            Width           =   4950
         End
         Begin VB.TextBox txtYTopQty 
            DataField       =   "YTop_Qty"
            DataSource      =   "datAccDet"
            Height          =   285
            Left            =   4860
            TabIndex        =   9
            Top             =   2888
            Width           =   1620
         End
         Begin VB.TextBox txtOpBal 
            DataField       =   "Open_Bal"
            DataSource      =   "datAccDet"
            Height          =   285
            Left            =   4860
            TabIndex        =   8
            Top             =   2258
            Width           =   1620
         End
         Begin VB.Frame fraOpBalType 
            Caption         =   "Opening Balance Type"
            Height          =   690
            Left            =   285
            TabIndex        =   5
            Top             =   2490
            Width           =   2505
            Begin VB.OptionButton OptCredit 
               Caption         =   "&Credit"
               Height          =   255
               Left            =   1455
               TabIndex        =   7
               Top             =   300
               Width           =   780
            End
            Begin VB.OptionButton optDebit 
               Caption         =   "&Debit"
               Height          =   255
               Left            =   270
               TabIndex        =   6
               Top             =   300
               Width           =   780
            End
         End
         Begin VB.Frame fraQtyToRec 
            Caption         =   "Quantity To Record"
            Height          =   690
            Left            =   285
            TabIndex        =   2
            Top             =   1485
            Width           =   2505
            Begin VB.OptionButton optNo 
               Caption         =   "&No"
               Height          =   225
               Left            =   1470
               TabIndex        =   4
               Top             =   315
               Width           =   660
            End
            Begin VB.OptionButton optYes 
               Caption         =   "&Yes"
               Height          =   225
               Left            =   300
               TabIndex        =   3
               Top             =   315
               Width           =   660
            End
         End
         Begin VB.Label lblAccDet 
            Caption         =   "Year Top Quantity"
            Height          =   210
            Index           =   4
            Left            =   3270
            TabIndex        =   16
            Top             =   2925
            Width           =   1290
         End
         Begin VB.Label lblAccDet 
            Caption         =   "Opening Balance"
            Height          =   210
            Index           =   3
            Left            =   3270
            TabIndex        =   15
            Top             =   2295
            Width           =   1290
         End
         Begin VB.Label lblAccDet 
            Caption         =   "Quantity Unit"
            Height          =   210
            Index           =   2
            Left            =   3270
            TabIndex        =   14
            Top             =   1665
            Width           =   1290
         End
         Begin VB.Label lblAccDet 
            Caption         =   "Description"
            Height          =   210
            Index           =   1
            Left            =   285
            TabIndex        =   13
            Top             =   982
            Width           =   1125
         End
         Begin VB.Label lblAccDet 
            Caption         =   "Account Code"
            Height          =   210
            Index           =   0
            Left            =   285
            TabIndex        =   12
            Top             =   442
            Width           =   1125
         End
      End
   End
End
Attribute VB_Name = "frmAccDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mdbsAccounts As Database    'To open the database
    Dim mrstCode As Recordset       'To open the Account Code table
    
    Dim mblnTrans As Boolean        'to store the transaction status
    Dim mblnFind As Boolean         'to store the find status
    Dim mvntBookmark As Variant     'to store the bookmark
Private Sub cmdAdd_Click()
    'Disables the spreedsheet view tab
    sstabAccDet.TabEnabled(1) = False
    'sets the transaction flag
    mblnTrans = True
    'invoking function to enable and disable respective controls
    'and to lock and unlock the controls
    Call DisableNaviButtons
    Call Lock_Controls(False)
    Call Enable_Controls(False)
    
    If Not mrstCode.EOF Then
        'storing the bookmark of the last viewed record
        mvntBookmark = mrstCode.Bookmark
        'moving to the end of table
        mrstCode.MoveLast
    End If
    'preparing for adding new info
    mrstCode.AddNew
    'resetting the controls so as to accept new information
    optYes.Value = False
    optNo.Value = False
    optDebit.Value = False
    OptCredit.Value = False
    txtQtyUnit.Enabled = False
    'setting the focus to Account code field
    txtAccCode.SetFocus
End Sub
Private Sub cmdCancel_Click()
    'resetting the transaction flag
    mblnTrans = False
    'cancelling the updation done so for
    mrstCode.CancelUpdate
    If Not mrstCode.EOF Then
        'moving to the last viewed record in case if one exists
        mrstCode.Bookmark = mvntBookmark
    Else
        'resetting the controls
        optYes.Value = False
        optNo.Value = False
        OptCredit.Value = False
        optDebit.Value = False
        txtQtyUnit.Enabled = False
        txtQtyUnit.Text = ""
    End If
    'invoking functions to enable and disable controls
    'and to lock and unlock the controls
    Call Lock_Controls(True)
    Call Enable_Controls(True)
    'enables the spreedsheet view tab
    sstabAccDet.TabEnabled(1) = True
End Sub
Private Sub cmdClose_Click()
    'unloads the form
    Unload Me
End Sub
Private Sub cmdDelete_Click()
    'getting confirmation from the user
    If MsgBox("Delete the Account?", vbYesNo) = vbYes Then
        'if yes, deletes the current info
        mrstCode.Delete
        'moving to the next info
        mrstCode.MoveNext
        If mrstCode.EOF And mrstCode.RecordCount <> 0 Then
            'if last info was deleted and still some info exist
            'moves to the last one
            mrstCode.MoveLast
        End If
        'invoking function to enable and disable controls
        Call Enable_Controls(True)
    End If
End Sub
Private Sub cmdFind_Click()
Dim rstTempCode As Recordset    'to open the account code table

    If Not mblnFind Then
        'if not in find mode
        'disables the spreedsheet view tab
        sstabAccDet.TabEnabled(1) = False
        'sets the find flag
        mblnFind = True
        cmdFind.Caption = "&Display"
        'bringing the code combo box for listing and hiding the code text box
        cboAccCode.Visible = True
        txtAccCode.Visible = False
        'invoking function to enable and disable controls
        Call DisableNaviButtons
        Call Enable_Controls(False)
        'saving the bookmark of the last viewed info
        mvntBookmark = mrstCode.Bookmark
        'opening the code table to list all the codes available
        Set rstTempCode = mdbsAccounts.OpenRecordset("Select Acc_Code from AccCode", dbOpenSnapshot)
        'clear the code combo box
        cboAccCode.Clear
        cboAccCode.Text = ""
        'adding all the codes to the combo box
        While Not rstTempCode.EOF
            cboAccCode.AddItem rstTempCode!Acc_Code
            rstTempCode.MoveNext
        Wend
        'setting focus to the code combo box
        cboAccCode.SetFocus
    ElseIf mblnFind Then
        'if in transaction mode
        'verifying the information
        If cboAccCode.Text = "" Then
            MsgBox "Select Account Code from the List"
            cboAccCode.SetFocus
        Else
            'if one selected
            'resetting the find flag
            mblnFind = False
            cmdFind.Caption = "&Find"
            'finding the info of the selected code
            mrstCode.FindFirst "Acc_Code = '" & cboAccCode.Text & "'"
            If mrstCode.NoMatch Then
                'if an invalid one is selected
                MsgBox "Entered Account Code does not exist"
                'moves to the last viewed info
                mrstCode.Bookmark = mvntBookmark
            End If
            'bringing the code text box and hiding the code combo box
            cboAccCode.Visible = False
            txtAccCode.Visible = True
            'invoking function to enable and disable controls
            'and to lock and unlock controls
            Call Enable_Controls(True)
            Call Lock_Controls(True)
            'enbles the spreedsheet view tab
            sstabAccDet.TabEnabled(1) = True
        End If
    End If
End Sub
Private Sub cmdFirst_Click()
    'moves to the first record
    mrstCode.MoveFirst
End Sub
Private Sub cmdLast_Click()
    'moves to the last record
    mrstCode.MoveLast
End Sub
Private Sub cmdModify_Click()
    'disables the spreedsheet view tab
    sstabAccDet.TabEnabled(1) = False
    'setting the transaction flag
    mblnTrans = True
    'invoking function to enable and disable controls
    'and to lock and unlock controls
    Call Lock_Controls(False)
    Call Enable_Controls(False)
    Call DisableNaviButtons
    'saves the bookmark of the current viewed record
    mvntBookmark = mrstCode.Bookmark
    'editing the current record
    mrstCode.Edit
    'locking the account code field so as to avoid editing
    txtAccCode.Locked = True
    'setting focus to description field
    txtAccDesc.SetFocus
End Sub
Private Sub cmdNext_Click()
    'moves to the next record
    mrstCode.MoveNext
End Sub
Private Sub cmdPrev_Click()
    'moving to the prevous record
    mrstCode.MovePrevious
End Sub
Private Sub cmdReset_Click()
    'Resetting the addition or modification action
    If mrstCode.EditMode = dbEditAdd Then
        'if in add mode
        'resetting controls
        optYes.Value = False
        optNo.Value = False
        optDebit.Value = False
        OptCredit.Value = False
        txtQtyUnit.Text = ""
        txtQtyUnit.Enabled = False
    ElseIf mrstCode.EditMode = dbEditInProgress Then
        'if in modification mode
        'setting Qty_toRec and Qty_Unit according to the current record
        If mrstCode!Qty_ToRec = "Y" Then
            optYes.Value = True
            txtQtyUnit.Enabled = False
        ElseIf mrstCode!Qty_ToRec = "N" Then
            optNo.Value = True
            txtQtyUnit.Enabled = False
        Else
            optYes.Value = False
            optNo.Value = False
        End If
        'setting the bal_type according to the one in the current record
        If mrstCode!Bal_Type = "C" Then
            OptCredit.Value = True
        ElseIf mrstCode!Bal_Type = "D" Then
            optDebit.Value = True
        Else
            OptCredit.Value = False
            optDebit.Value = False
        End If
    End If
    'updating controls according to the current record
    datAccDet.UpdateControls
End Sub
Private Sub cmdSave_Click()
'checking for error
On Error GoTo ErrorHandler
    'invoking function that validates the input data
    If FieldCheck Then
        'resetting the transaction flag
        mblnTrans = False
        'recording the Qty_ToRec field with the value
        'according to the option selected
        If optYes.Value Then
            mrstCode!Qty_ToRec = "Y"
        ElseIf optNo.Value Then
            mrstCode!Qty_ToRec = "N"
        End If
        'recording the Bal_Type field with the value
        'according to the option selected
        If OptCredit.Value Then
            mrstCode!Bal_Type = "C"
        ElseIf optDebit.Value Then
            mrstCode!Bal_Type = "D"
        Else
            mrstCode!Bal_Type = ""
        End If
        'saving the values entered
        mrstCode.Update
        'moving to the last saved record
        mrstCode.Bookmark = mrstCode.LastModified
        'invoking funtion that enables and disables controls,
        'and locks and unlocks controls according to the requirements
        Call Enable_Controls(True)
        Call Lock_Controls(True)
        'enabling the spreedsheet view tab
        sstabAccDet.TabEnabled(1) = True
    End If
    Exit Sub

ErrorHandler:
    'checking for error
    If Err.Number = 3022 Then
        MsgBox "Account Code already exists"
        mblnTrans = True
        txtAccCode.SetFocus
    Else
        MsgBox Err.Description
    End If
End Sub
Private Sub datAccDet_Reposition()
    'setting the controls according to the datas in the current record
    If Not mblnTrans And Not mblnFind Then
        'if not transaction and find mode
        If Not mrstCode.EOF Then
            'if some records exists
            'setting the qty_torec and qty_unit fields according
            'to the one in the current record
            If mrstCode!Qty_ToRec = "Y" Then
                optYes.Value = True
                txtQtyUnit.Enabled = True
            ElseIf mrstCode!Qty_ToRec = "N" Then
                optNo.Value = True
                txtQtyUnit.Enabled = False
            End If
            'formatting the open_Bal field
            txtOpBal.Text = Format(mrstCode!Open_Bal, "#####.00")
            'setting the Bal_Type fields according to the one
            'in the current record
            If mrstCode!Bal_Type = "C" Then
                OptCredit.Value = True
            ElseIf mrstCode!Bal_Type = "D" Then
                optDebit.Value = True
            Else
                OptCredit.Value = False
                optDebit.Value = False
            End If
            'invoking the function that enables and disables
            'navigation buttons according to the current record
            Call RecordPosition
        Else
            'resetting all the fields if no record exists
            optYes.Value = False
            optNo.Value = False
            optDebit.Value = False
            OptCredit.Value = False
        End If
    End If
End Sub
Private Sub Form_Load()
    Dim strCond As String   'to store the query for opening a table

    'opening the database
    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    'setting the query for opening the AccCode table
    strCond = "Select * from AccCode Order By Acc_Desc"
    Set mrstCode = mdbsAccounts.OpenRecordset(strCond, dbOpenDynaset)
    'setting the datacontrols to the table openned
    Set datAccDet.Recordset = mrstCode
    If Not mrstCode.EOF Then
        'populating the recordset
        mrstCode.MoveLast
        mrstCode.MoveFirst
    End If
    'resetting the transaction and find flag
    mblnTrans = False
    mblnFind = False
    'invoking funtion that enables and disables controls,
    'and locks and unlocks controls according to the requirements
    Call Lock_Controls(True)
    Call Enable_Controls(True)
    Call RecordPosition
End Sub
Private Sub Enable_Controls(blnEnable As Boolean)
    'Enables or disables the controls according to the requirements
    'blnEnable  'to store the flag to enable controls
    cmdAdd.Enabled = blnEnable
    If blnEnable And mrstCode.RecordCount <> 0 Then
        'if blnenable flag is true and if records exists
        'enables the Modify and Delete buttons
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    Else
        'disables the Modify and Delete buttons
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
    End If
    If (blnEnable And mrstCode.RecordCount > 1) Or mblnFind Then
        'if blnenable flag is true and if more than one record exist
        'or if find mode
        'enables the find button
        cmdFind.Enabled = True
    Else
        'disables the find button
        cmdFind.Enabled = False
    End If
    'setting the save,reset and cancel buttons
    'according to the transaction flag
    cmdSave.Enabled = mblnTrans
    cmdReset.Enabled = mblnTrans
    cmdCancel.Enabled = mblnTrans
End Sub
Private Sub RecordPosition()
    'enables or disables the navigaiton according
    'to the current record
    If mrstCode.RecordCount = 0 Or mrstCode.RecordCount = 1 Then
        'if no record exists or if only one record exists
        'disables all the navigation button
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdLast.Enabled = False
    ElseIf mrstCode.AbsolutePosition = 0 Then
        'if the first record is the current one
        'disables first and previous buttons
        'enables the next and last butons
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    ElseIf mrstCode.AbsolutePosition = (mrstCode.RecordCount - 1) Then
        'if the last record  is the current one
        'enables first and prev buttons
        'disables next and last buttons
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = False
        cmdLast.Enabled = False
    Else
        'if middle record is the current one
        'enables all the navigation buttons
        cmdFirst.Enabled = True
        cmdPrev.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
    End If
End Sub
Private Sub Lock_Controls(blnLock As Boolean)
    'enabling or disabling the controls
    'according to the input
    txtAccCode.Locked = blnLock
    txtAccDesc.Locked = blnLock
    txtQtyUnit.Locked = blnLock
    txtOpBal.Locked = blnLock
    txtYTopQty.Locked = blnLock
    
    optYes.Enabled = Not blnLock
    optNo.Enabled = Not blnLock
    optDebit.Enabled = Not blnLock
    OptCredit.Enabled = Not blnLock
End Sub
Private Sub DisableNaviButtons()
    'disables all the navigation buttons
    cmdFirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False
End Sub
Private Sub optNo_Click()
    If mblnTrans Then
        'if in transaction mode
        'disables the qty_unit field
        txtQtyUnit.Enabled = False
        txtQtyUnit.Text = ""
    End If
End Sub
Private Sub optYes_Click()
    If mblnTrans Then
        'if in transaction mode
        'prepares the qty_unit field for edition
        txtQtyUnit.Enabled = True
        txtQtyUnit.Text = ""
    End If
End Sub
Private Sub txtAccCode_KeyPress(KeyAscii As Integer)
    'restricting values entered
    If KeyAscii <> 8 And mblnTrans Then
        If (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            Beep
            KeyAscii = 0
        Else
            'converting the character entered into uppercase
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Len(txtAccCode.Text) = 0 Then
                'restricting for only certain character as the first one
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
Private Sub txtAccDesc_KeyPress(KeyAscii As Integer)
    'restricting values entered
    If mblnTrans Then
        'if in transaction mode
        If (KeyAscii <> 8) And (KeyAscii <> Asc(",")) And _
            (KeyAscii <> Asc("-")) And (KeyAscii <> Asc(".")) And _
            (KeyAscii <> Asc(" ")) And (KeyAscii <> Asc("/")) And _
            (KeyAscii <> Asc("(")) And (KeyAscii <> Asc(")")) And _
            (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> Asc("&")) Then
            KeyAscii = 0
        ElseIf Len(txtAccDesc.Text) = 0 Then
            'converting the characted entered into uppercase
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If

    End If

End Sub
Private Sub txtOpBal_KeyPress(KeyAscii As Integer)
    'restricting values entered
    If mblnTrans Then
        'if in transaction mode
        If (KeyAscii <> 8) And (KeyAscii <> Asc(".")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtQtyUnit_KeyPress(KeyAscii As Integer)
    'restricting values entered
    If mblnTrans Then
        'if in transaction mode
        If (KeyAscii <> 8) And (KeyAscii <> Asc("&")) And _
            (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtYTopQty_KeyPress(KeyAscii As Integer)
    'restricting values entered
    If mblnTrans Then
        'if in transaction mode
        If (KeyAscii <> 8) And (KeyAscii <> Asc("-")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> Asc(".")) Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub
Private Function FieldCheck() As Boolean

    FieldCheck = False
    If txtAccCode.Text = "" Then
        MsgBox "Account Code should not be blank"
        txtAccCode.SetFocus
    ElseIf txtAccDesc.Text = "" Then
        MsgBox "Description should not be blank"
        txtAccDesc.SetFocus
    ElseIf Not optYes.Value And Not optNo.Value Then
        MsgBox "Specify whether Quantity should be recorded or not?"
        optYes.SetFocus
    ElseIf optYes.Value And txtQtyUnit.Text = "" Then
        MsgBox "Quantity Unit should not be blank"
        txtQtyUnit.SetFocus
    ElseIf txtOpBal.Text <> "" Then
        If Not CheckingNum(txtOpBal.Text) Then
            MsgBox "Invalid Opening Balance"
            txtOpBal.SetFocus
        ElseIf Not OptCredit.Value And Not optDebit.Value Then
            MsgBox "Specify the type of the opening balance"
            optDebit.SetFocus
        Else
            FieldCheck = True
        End If
    ElseIf txtOpBal.Text = "" Then
        If OptCredit.Value Or optDebit.Value Then
            MsgBox "Enter the Opening Balance"
            txtOpBal.SetFocus
        Else
            FieldCheck = True
        End If
    End If
    If FieldCheck Then
        If txtYTopQty.Text <> "" Then
            If Not CheckingNum(txtYTopQty.Text) Then
                MsgBox "Invalid Quantity"
                txtYTopQty.SetFocus
                FieldCheck = False
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
