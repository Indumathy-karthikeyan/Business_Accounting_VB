VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSales 
   BorderStyle     =   0  'None
   Caption         =   "Sales Register"
   ClientHeight    =   8925
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraRegister 
      Height          =   8775
      Index           =   0
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton cmdAccCodeSel 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   375
         Left            =   3120
         TabIndex        =   53
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdNarrationSel 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   375
         Left            =   7320
         TabIndex        =   52
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox txtNarration 
         Height          =   375
         Left            =   1800
         TabIndex        =   51
         Top             =   1680
         Width           =   5415
      End
      Begin VB.TextBox txtAccDesc 
         Height          =   375
         Left            =   5280
         TabIndex        =   50
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtAccCode 
         Height          =   375
         Left            =   1800
         TabIndex        =   49
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Frame fraRegister 
         Height          =   975
         Index           =   3
         Left            =   480
         TabIndex        =   37
         Top             =   7440
         Width           =   10575
         Begin VB.CommandButton cmdClose 
            Caption         =   "Cl&ose"
            Height          =   495
            Left            =   9120
            TabIndex        =   45
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   495
            Left            =   3960
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset"
            Height          =   495
            Left            =   5040
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   495
            Left            =   6120
            TabIndex        =   42
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   495
            Left            =   360
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "&Modify"
            Height          =   495
            Left            =   1440
            TabIndex        =   40
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   495
            Left            =   2520
            TabIndex        =   39
            Top             =   240
            Width           =   975
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   495
            Left            =   7800
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         Height          =   375
         Left            =   8520
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   375
         Left            =   9000
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   375
         Left            =   9480
         TabIndex        =   17
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "<<"
         Height          =   375
         Left            =   8040
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtBillNo 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtDate 
         Height          =   375
         Left            =   5280
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtAmount 
         Height          =   375
         Left            =   9000
         TabIndex        =   24
         Top             =   1680
         Width           =   1815
      End
      Begin TabDlg.SSTab sstabRegister 
         Height          =   5055
         Left            =   480
         TabIndex        =   22
         Top             =   2280
         Width           =   10605
         _ExtentX        =   18706
         _ExtentY        =   8916
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Form View"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fraRegister(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "fraRegister(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdProdFirst"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdProdPrev"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdProdNext"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdProdLast"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Spreedsheet View"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "msflxgrdRegister"
         Tab(1).ControlCount=   1
         Begin VB.CommandButton cmdProdLast 
            Caption         =   ">>"
            Height          =   375
            Left            =   8760
            TabIndex        =   21
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton cmdProdNext 
            Caption         =   ">"
            Height          =   375
            Left            =   8280
            TabIndex        =   20
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton cmdProdPrev 
            Caption         =   "<"
            Height          =   375
            Left            =   7800
            TabIndex        =   19
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton cmdProdFirst 
            Caption         =   "<<"
            Height          =   375
            Left            =   7320
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
         Begin VB.Frame fraRegister 
            Height          =   3015
            Index           =   2
            Left            =   360
            TabIndex        =   28
            Top             =   600
            Width           =   9855
            Begin VB.CommandButton cmdProdAccCodeSel 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   375
               Left            =   2760
               TabIndex        =   3
               Top             =   600
               Width           =   495
            End
            Begin VB.CommandButton cmdProdNarrationSel 
               Appearance      =   0  'Flat
               Caption         =   "..."
               Height          =   375
               Left            =   7920
               TabIndex        =   5
               Top             =   1320
               Width           =   495
            End
            Begin VB.TextBox txtProdAmount 
               Height          =   375
               Left            =   4680
               TabIndex        =   7
               Top             =   2040
               Width           =   1575
            End
            Begin VB.TextBox txtProdQuantity 
               Height          =   405
               Left            =   1440
               TabIndex        =   6
               Top             =   2010
               Width           =   1935
            End
            Begin VB.TextBox txtProdNarration 
               Height          =   375
               Left            =   1440
               TabIndex        =   4
               Top             =   1320
               Width           =   6255
            End
            Begin VB.TextBox txtProdAccDesc 
               Height          =   375
               Left            =   4680
               TabIndex        =   29
               Top             =   600
               Width           =   4815
            End
            Begin VB.TextBox txtProdAccCode 
               Height          =   375
               Left            =   1440
               TabIndex        =   2
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lblRegister 
               Caption         =   "Amount"
               Height          =   255
               Index           =   9
               Left            =   3720
               TabIndex        =   34
               Top             =   2160
               Width           =   735
            End
            Begin VB.Label lblRegister 
               Caption         =   "Quantity"
               Height          =   255
               Index           =   8
               Left            =   240
               TabIndex        =   33
               Top             =   2160
               Width           =   1095
            End
            Begin VB.Label lblRegister 
               Caption         =   "Narration "
               Height          =   255
               Index           =   10
               Left            =   240
               TabIndex        =   32
               Top             =   1440
               Width           =   735
            End
            Begin VB.Label lblRegister 
               Caption         =   "Account Description"
               Height          =   495
               Index           =   7
               Left            =   3720
               TabIndex        =   31
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label lblRegister 
               Caption         =   "Account Code"
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   30
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.Frame fraRegister 
            Height          =   975
            Index           =   1
            Left            =   360
            TabIndex        =   26
            Top             =   3720
            Width           =   9855
            Begin VB.CommandButton cmdProdModify 
               Caption         =   "&Modify"
               Height          =   375
               Left            =   2280
               TabIndex        =   9
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdProdAdd 
               Caption         =   "&Add"
               Height          =   375
               Left            =   1200
               TabIndex        =   8
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdProdDelete 
               Caption         =   "&Delete"
               Height          =   375
               Left            =   3360
               TabIndex        =   10
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdProdSave 
               Caption         =   "&Save"
               Height          =   375
               Left            =   5640
               TabIndex        =   11
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdProdReset 
               Caption         =   "&Reset"
               Height          =   375
               Left            =   6720
               TabIndex        =   12
               Top             =   360
               Width           =   735
            End
            Begin VB.CommandButton cmdProdCancel 
               Caption         =   "&Cancel"
               Height          =   375
               Left            =   7800
               TabIndex        =   13
               Top             =   360
               Width           =   735
            End
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdRegister 
            Height          =   4215
            Left            =   -74760
            TabIndex        =   27
            Top             =   600
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7435
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
         End
      End
      Begin VB.Label lblRegister 
         Caption         =   "Narration"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   48
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label lblRegister 
         Caption         =   "Account Description"
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   47
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblRegister 
         Caption         =   "Account Code"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   46
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label lblRegister 
         Caption         =   "Bill No."
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   36
         Top             =   540
         Width           =   615
      End
      Begin VB.Label lblRegister 
         Caption         =   "Date"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   35
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblRegister 
         Caption         =   "Amount"
         Height          =   255
         Index           =   5
         Left            =   8160
         TabIndex        =   25
         Top             =   1740
         Width           =   735
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuNavigator 
      Caption         =   "Navigator"
      Begin VB.Menu mnuFirst 
         Caption         =   "First"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "Previous"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLast 
         Caption         =   "Last"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'database
Dim mdbsAccounts As Database
'recordset of register
Dim mrstRegister As Recordset
'Bill Transaction flag
Dim mblnTrans As Boolean
'product transaction flag
Dim mblnProd As Boolean
'finding bill flag
Dim mblnFind As Boolean
'current billno
Dim mintBillNo As Integer
'new billno
Dim mintNewBillNo As Integer
'new bill date
Dim mstrNewDate As String
'current product no
Dim mintProdNo As Integer
'type of bill transaction
Dim mstrTransType As String
'type of prod transaction
Dim mstrProdTransType As String
'total no of product in the current bill
Dim mintProdTotNo As Integer
'to store the old date
Dim mstrDate As String
Private Sub PreparingForm()
    'hiding acccode,accdesc and narration text controls in case of cash transaction
    lblRegister(2).Visible = False
    txtAccCode.Visible = False
    lblRegister(3).Visible = False
    txtAccDesc.Visible = False
    lblRegister(4).Visible = False
    txtNarration.Visible = False
    
    'displaying amount field in the place of acccode field in case of cash transaction
    lblRegister(5).Top = lblRegister(2).Top
    lblRegister(5).Left = lblRegister(2).Left
    txtAmount.Left = txtAccCode.Left
    txtAmount.Top = txtAccCode.Top
    
    'moving tab control and transaction frame up
    sstabRegister.Top = txtNarration.Top
    fraRegister(3).Top = fraRegister(3).Top - 600
    'reducing the height of the frame
    fraRegister(0).Height = fraRegister(0).Height - 600
    Me.Height = Me.Height - 600
End Sub

Private Sub cmdAccCodeSel_Click()
    gstrFormName = "Sales AccCode"
    gstrDetName = "Account"
    frmDetails.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    If (mblnTrans) Then
        'cancelling bill transaction
        If mstrTransType = "Add" Then
            'cancelling addition of Bill
            If (mrstRegister.BOF And mrstRegister.EOF) Then
                'if the register is empty
                'clearing all controls
                Call ClearingMainText
                Call ClearingProdText
            Else
                'displaying previously displayed details
                Call DisplayingDetails
                Call DisplayingProdDetails
            End If
        ElseIf (mstrTransType = "Modify") Then
            'cancelling modification of bill
            'displaying origiinal bill
            Call DisplayingDetails
            Call DisplayingProdDetails
        End If
        'resetting bill transaction flag
        mblnTrans = False
        mstrTransType = ""
        cmdAccCodeSel.Visible = False
        cmdNarrationSel.Visible = False
    ElseIf (mblnFind) Then
        'cancelling finding bill
        mblnFind = False
        cmdFind.Caption = "&Find"
        'enabling spreadsheet view tab
        sstabRegister.TabEnabled(1) = True
        'displaying current bill details
        Call DisplayingDetails
        Call DisplayingProdDetails
    End If
    'locking text controls
    Call LockTextControls
    'enabling transaction controls
    Call EnableControls
    Call ProdEnableControls
    'enabling navigation controls
    Call RecordPosition
    Call ProdRecordPosition
End Sub

Private Sub cmdClose_Click()
    If (Not mblnTrans) Then
        'if not in transation mode closing form
        Unload Me
    End If
End Sub

Private Sub cmdDelete_Click()
'condition for openning recordset
Dim strCondition As String
'condition for finding a record
Dim strFindCondition As String
'recordset
Dim rstBillNo As Recordset

    If MsgBox("Do you want to delete it? ", vbOKCancel) = vbOK Then
        'Saving cancelled bill details for further use
        Call SavingCancelledBills
        'deleting current bill
        mintProdTotNo = msflxgrdRegister.Rows - 1
        Call Deleting_Bill
        'moving to next bill
        'getting Bill details
        If gstrRegister = "Cash Sales" Then
            strCondition = "select distinct BillNo from Cash where (BillNo <> 0) order by BillNo"
        ElseIf gstrRegister = "Credit Sales" Then
            strCondition = "select distinct BillNo from Journal where (BillNo <> 0) order by BillNo"
        End If

        Set rstBillNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
        If (rstBillNo.BOF And rstBillNo.EOF) Then
            'if no bill exist
            'resetting text controls
            Call ClearingMainText
            Call ClearingProdText
            msflxgrdRegister.Rows = 1
            mintBillNo = 0
        Else
            'finding the next bill
            strFindCondition = "BillNo = " & mintBillNo + 1
            rstBillNo.FindFirst strFindCondition
            If (rstBillNo.NoMatch) Then
                'if the one deleted is the last bill
                'getting the last bill present
                rstBillNo.MoveLast
            End If
            'displaying bill details
            mintBillNo = rstBillNo!BillNo
            Call DisplayingDetails
            Call DisplayingProdDetails
        End If
        'resetting navigation buttons
        Call RecordPosition
        Call ProdRecordPosition
        'locking the text controls
        Call LockTextControls
        'enabling the transaction buttons
        Call EnableControls
        Call ProdEnableControls
    End If
End Sub
Private Sub SavingCancelledBills()
'condition for opening the recordset
Dim strCondition As String
'recordset
Dim rstCancelledBills As Recordset

    'saving the cancelled bill
    strCondition = "Select * from CancelledBills"
    Set rstCancelledBills = mdbsAccounts.OpenRecordset(strCondition, dbOpenDynaset)
    With rstCancelledBills
        .AddNew
        !BillNo = txtBillNo.Text
        !Date = CDate(txtDate)
        If gstrRegister = "Cash Sales" Then
            !Type = "C"
        ElseIf gstrRegister = "Credit Sales" Then
            !Type = "J"
        End If
        .Update
    End With
End Sub

Private Sub cmdFind_Click()
'condition for finding the bill
Dim strCondition As String
'recordset for cancelled bills
Dim rstBill As Recordset
    
    If Not mblnFind Then
        'setting find mode
        mblnFind = True
        'setting the caption to display
        cmdFind.Caption = "&Display"
        'disabling the spreadsheet view tab
        sstabRegister.TabEnabled(1) = False
        'disabling navigation buttons
        Call DisableNaviButtons
        'disabling the Transaction buttons
        Call EnableControls
        Call ProdEnableControls
        'resetting the text controls
        Call ClearingMainText
        Call ClearingProdText
        'unlocking the billno for entry
        txtBillNo.Locked = False
        txtBillNo.SetFocus
    Else
        'if in find mode
        If txtBillNo.Text = "" Then
            MsgBox "Enter the BillNo", vbOKOnly
            txtBillNo.SetFocus
        Else
            'if the billno is not empty
            'setting the condition for finding the bill
            strCondition = "BillNo = " & txtBillNo.Text
            mrstRegister.FindFirst strCondition
            If mrstRegister.NoMatch Then
                'checking for the bill in cancelledbills table
                strCondition = "Select * from Cancelledbills where BillNo = " & txtBillNo.Text
                Set rstBill = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
                If (rstBill.BOF And rstBill.EOF) Then
                    'if bill does not exists
                    MsgBox "BillNo does not exists", vbOKOnly
                Else
                    MsgBox "Bill " & txtBillNo.Text & " has been cancelled", vbOKOnly
                End If
                txtBillNo.Text = ""
                txtBillNo.SetFocus
            Else
                mintBillNo = mrstRegister!BillNo
                'resetting the find mode
                mblnFind = False
                cmdFind.Caption = "&Find"
                txtBillNo.Locked = True
                'displaying the details of the bill found
                Call DisplayingDetails
                Call DisplayingProdDetails
                'enabling the transaction buttons
                Call EnableControls
                Call ProdEnableControls
                'enabling the navigation buttons
                Call RecordPosition
                Call ProdRecordPosition
                'enabling the spreadsheetview tab
                sstabRegister.TabEnabled(1) = True
            End If
        End If
    End If
End Sub

Private Sub cmdNarrationSel_Click()
    gstrFormName = "Sales Narration"
    gstrDetName = "Narration"
    frmDetails.Show vbModal
End Sub

Private Sub cmdProdAccCodeSel_Click()
    gstrFormName = "Sales ProdAccCode"
    gstrDetName = "Account"
    frmDetails.Show vbModal
End Sub

Private Sub cmdProdNarrationSel_Click()
    gstrFormName = "Sales ProdNarration"
    gstrDetName = "Narration"
    frmDetails.Show vbModal
End Sub

Private Sub cmdReset_Click()
    If (mstrTransType = "Add") Then
        'setting the text field blank in case of Adding Bill
        Call ResettingControls
        'clearing the grid control
        msflxgrdRegister.Rows = 1
        mintProdNo = 1
        'disabling the product transaction buttons
        cmdProdModify.Enabled = False
        cmdProdDelete.Enabled = False
    ElseIf (mstrTransType = "Modify") Then
        'displaying the existing bill details in case of modifying the bill
        Call DisplayingDetails
        Call DisplayingProdDetails
        If (msflxgrdRegister.Rows = 1) Then
            'disabling product transaction buttons in case there is only one product
            cmdProdModify.Enabled = False
            cmdProdDelete.Enabled = False
        End If
        'enabling product navigation buttons
        Call ProdRecordPosition
    End If
    txtDate.SetFocus
End Sub

Private Sub cmdSave_Click()
    'checking the fields before saving
    If (FieldCheck) Then
        'if all the necessary informations are provided
        'checking if the product details are provided
        If msflxgrdRegister.Rows <= 1 Then
            MsgBox "Product details are missing", vbOKOnly
        Else
            If (mstrTransType = "Add") Then
                'in case of adding a bill
                Call Adding_Bill
            ElseIf mstrTransType = "Modify" Then
                'in case of modifying a bill
                'deleting the existing bill details
                Call Deleting_Bill
                'adding the new details
                Call Adding_Bill
                mintProdTotNo = 0
            End If
            'setting the information as the current bill
            mintBillNo = txtBillNo.Text
            'resetting the transaction mode
            mblnTrans = False
            mstrTransType = ""
            'locking the text controls
            Call LockTextControls
            'enabling the transaction buttons
            Call EnableControls
            Call ProdEnableControls
            'enabling navigation buttons
            Call RecordPosition
            Call ProdRecordPosition
        End If
        cmdAccCodeSel.Visible = False
        cmdNarrationSel.Visible = False
    End If
End Sub
Private Sub Adding_Bill()
'condition for opening the recordset
Dim strCondition As String
'recordset
Dim rstEntryNo As Recordset
'entryno in the cash table
Dim intEntryNo As Integer
'current row no in the grid
Dim intCurRow As Integer

    If gstrRegister = "Cash Sales" Then
        'getting the entry no from the cash table in case of Cash transaction
        strCondition = "Select EntryNo from Cash order by EntryNo"
        Set rstEntryNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    
        If (Not rstEntryNo.BOF And Not rstEntryNo.EOF) Then
            'if records exist getting the last entry no
            rstEntryNo.MoveLast
            intEntryNo = rstEntryNo!Entryno + 1
        Else
            'if no records
            intEntryNo = 1
        End If
    ElseIf gstrRegister = "Credit Sales" Then
        'getting the entry no from the Journal table in case of credit transaction
        strCondition = "Select EntryNo from Journal order by EntryNo"
        Set rstEntryNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    
        If (Not rstEntryNo.BOF And Not rstEntryNo.EOF) Then
            'if records exist getting the last entry no
            rstEntryNo.MoveLast
            intEntryNo = rstEntryNo!Entryno + 1
        Else
            'if no records
            intEntryNo = 1
        End If
    End If
    
    'saving the personal details in case of credit transaction
    If gstrRegister = "Credit Sales" Then
        With mrstRegister
            .AddNew
            !Entryno = intEntryNo
            !BillNo = Int(txtBillNo.Text)
            !Date = CDate(txtDate.Text)
            !Acc_Code = txtAccCode.Text
            If (txtNarration.Text <> "") Then
                mrstRegister!Narration = txtNarration.Text
            End If
            ![Debit/Credit] = "D"
            !Amount = txtAmount.Text
            .Update
        End With
        intEntryNo = intEntryNo + 1
    End If
            
    With msflxgrdRegister
        intCurRow = 1
        'looping until the current row no is less the no of rows in the grid
        While (intCurRow < .Rows)
            'setting the current row
            .Row = intCurRow
            'adding information to the table
            mrstRegister.AddNew
            mrstRegister!Entryno = intEntryNo
            mrstRegister!BillNo = Int(txtBillNo.Text)
            mrstRegister!Date = CDate(txtDate.Text)
            .Col = 0
            mrstRegister!Acc_Code = .Text
            .Col = 2
            If (.Text <> "") Then
                mrstRegister!Narration = .Text
            End If
            mrstRegister![Debit/Credit] = "C"
            .Col = 3
            mrstRegister!Quantity = .Text
            .Col = 4
            mrstRegister!Amount = .Text
            mrstRegister.Update
            'increasing the current row
            intCurRow = intCurRow + 1
            'increasing the entry no
            intEntryNo = intEntryNo + 1
        Wend
    End With
End Sub
Private Sub Deleting_Bill()
'current record in the table
Dim intCurRec As Integer
'condition for finding the bill no
Dim strFindCond As String

    intCurRec = 1
    strFindCond = "BillNo = " & mintBillNo
    'deleting personel details in case of credit transaction
    If gstrRegister = "Credit Sales" Then
        mrstRegister.FindFirst strFindCond
        mrstRegister.Delete
    End If
    'looping until the current records is equal to total products
    Do While (intCurRec <= mintProdTotNo)
        mrstRegister.FindFirst strFindCond
        If mrstRegister.NoMatch Then
            Exit Do
        End If
        'deleting the current record
        mrstRegister.Delete
        'increasing the current record no
        intCurRec = intCurRec + 1
    Loop
End Sub


Private Sub cmdAdd_Click()
    'storing previous date
    mstrDate = txtDate.Text
    'setting transaction mode
    mblnTrans = True
    mstrTransType = "Add"
    'clearing the product details in the grid
    msflxgrdRegister.Rows = 1
    'calculating the new bill no and the date
    Call CalBillNoDate
    'disabling navigation buttons
    Call DisableNaviButtons
    Call DisableProdNaviButtons
    'unlocking the text controls
    Call LockTextControls
    'disabling transaction buttons
    Call EnableControls
    Call ProdEnableControls
    'clearing text controls
    Call ResettingControls
    'setting the product no
    mintProdNo = 0
    'setting the total amount of the bill
    txtAmount.Text = "0"
    cmdAccCodeSel.Visible = True
    cmdNarrationSel.Visible = True
    
    txtBillNo.SetFocus
End Sub
Private Sub DisableNaviButtons()
    'disabling  main navigation buttons during transaction
    cmdFirst.Enabled = False
    cmdPrev.Enabled = False
    cmdNext.Enabled = False
    cmdLast.Enabled = False

    mnuFirst.Enabled = False
    mnuPrev.Enabled = False
    mnuNext.Enabled = False
    mnuLast.Enabled = False
End Sub
Private Sub DisableProdNaviButtons()
    'disabling product navigation buttons during transaction
    cmdProdFirst.Enabled = False
    cmdProdPrev.Enabled = False
    cmdProdNext.Enabled = False
    cmdProdLast.Enabled = False
End Sub
Private Sub ClearingMainText()
    'setting main text controls blank
    If (gstrRegister = "Credit Sales") Then
        txtAccCode.Text = ""
        txtAccDesc.Text = ""
        txtNarration.Text = ""
    End If
    txtBillNo.Text = ""
    txtDate.Text = ""
    txtAmount.Text = ""
End Sub
Private Sub ClearingProdText()
    'setting product text controls blank
    txtProdAccCode.Text = ""
    txtProdAccDesc.Text = ""
    txtProdNarration.Text = ""
    txtProdQuantity.Text = ""
    txtProdAmount.Text = ""
End Sub
Private Sub CalBillNoDate()
'condition for opeining a recordset
Dim strCondition As String
'recordset for Cash table
Dim rstBill As Recordset
'recordset for CancelledBills table
Dim rstCancelledBill As Recordset
'last billno in cash table
Dim intBillNo As Integer
'last billno in cancelledbills table
Dim intCancelledBillNo As Integer
'last date in cash table
Dim strDate As String
'last date in cancelledbills table
Dim strCancelledDate As String
 
'   setting billno to the next no of the current  record
    If mrstRegister.EOF And mrstRegister.BOF Then
        mintNewBillNo = 1
        strDate = Date
    Else
        mintNewBillNo = mrstRegister!BillNo + 1
        strDate = mrstRegister!Date
    End If
'    If gstrRegister = "Cash Sales" Then
'        'setting condition for getting billno and date from cash table in case of cash transaction
'        strCondition = "Select distinct BillNo,Date from Cash order by BillNo"
'        Set rstBill = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
'        If (Not rstBill.BOF) And (Not rstBill.EOF) Then
'            rstBill.MoveLast
'            'getting the last billno and date
'            intBillNo = rstBill!BillNo
'            strDate = rstBill!Date
'        Else
'            'in case of no bills
'            strDate = ""
'            intBillNo = 0
'        End If
'    ElseIf gstrRegister = "Credit Sales" Then
'        'setting condition for getting billno and date from journal table in case of credit transation
'        strCondition = "Select distinct BillNo,Date from Journal order by BillNo"
'        Set rstBill = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
'        If (Not rstBill.BOF) And (Not rstBill.EOF) Then
'            rstBill.MoveLast
'            'getting the last billno and date
'            intBillNo = rstBill!BillNo
'            strDate = rstBill!Date
'        Else
'            'in case of no bills
'            strDate = ""
'            intBillNo = 0
'        End If
'    End If
    
'    'setting condition for getting billno and date from Cancelledbills table
'    If gstrRegister = "Cash Sales" Then
'        strCondition = "Select distinct BillNo,Date from CancelledBills where Type = 'C' order by BillNo"
'    ElseIf gstrRegister = "Credit Sales" Then
'        strCondition = "Select distinct BillNo,Date from CancelledBills where Type = 'J' order by BillNo"
'    End If
'    Set rstCancelledBill = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
'    If (Not rstCancelledBill.BOF) And (Not rstCancelledBill.EOF) Then
'        rstCancelledBill.MoveLast
'        'getting the last billno and date
'        intCancelledBillNo = rstCancelledBill!BillNo
'        strCancelledDate = rstCancelledBill!Date
'    Else
'        'in case of no bills
'        intCancelledBillNo = 0
'        strCancelledDate = ""
'    End If
'
'    If (intBillNo >= intCancelledBillNo) Then
'        'in case cash billno is greater or equal to billno in cancelledbills
'        'getting new bill no and date from cash table
'        mintNewBillNo = intBillNo + 1
'        mstrNewDate = strDate
'    Else
'        'in case billno in cancelledbills table is greater
'        'getting new billno and date from Cancelledbills table
'        mintNewBillNo = intCancelledBillNo + 1
'        mstrNewDate = strCancelledDate
'    End If
    'in case of first bill setting the date as the current date
    If mintNewBillNo = 1 Then
        mstrNewDate = Date
    Else
        mstrNewDate = strDate
    End If
End Sub
Private Function CheckProdDet() As Boolean

'condition for opening the recordset
Dim strCondition As String
'recordset
Dim rstCode As Recordset
    
    'setting verification as false
    CheckProdDet = False
    'checking presence of product information
    If txtProdAccCode.Text = "" Then
        MsgBox "Account Code should not be blank"
        txtProdAccDesc.Text = ""
        txtProdAccCode.SetFocus
    Else
        'getting the description of the product from AccCode table
        strCondition = "Select Acc_Code from AccCode where Acc_Code = '" _
            & txtProdAccCode.Text & "'"
        Set rstCode = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
        If rstCode.EOF Then
            MsgBox "Invalid AccountCode"
            txtProdAccDesc.Text = ""
            txtProdAccCode.SetFocus
        ElseIf (Left(txtProdAccCode.Text, 1) = "I") And (txtProdQuantity.Text = "") Then
            MsgBox "Quantity should not be blank", vbOKOnly
            txtProdQuantity.SetFocus
        ElseIf txtProdAmount.Text = "" Then
            MsgBox "Amount should not be blank", vbOKOnly
            txtProdAmount.SetFocus
        Else
            'if all the information of the product is given
            'setting verification to true
            CheckProdDet = True
        End If
    End If
End Function

Private Sub cmdProdDelete_Click()
'current product no
Dim intCurProd As Integer

    If MsgBox("Do you want to Delete it?", vbOKCancel) = vbOK Then
        'deleting current product
        With msflxgrdRegister
            If (mintProdNo < (.Rows - 1)) Then
                'if the current product is the first on or is in middle of the grid
                intCurProd = mintProdNo
                'deleting the current product
                'moving the successive products one step before
                While (intCurProd < (.Rows - 1))
                    .Row = intCurProd
                    .Col = 0
                    .Text = .TextMatrix(intCurProd + 1, 0)
                    .Col = 1
                    .Text = .TextMatrix(intCurProd + 1, 1)
                    .Col = 2
                    If (.TextMatrix(intCurProd + 1, 2) <> "") Then
                        .Text = .TextMatrix(intCurProd + 1, 2)
                    Else
                        .Text = ""
                    End If
                    .Col = 3
                    .Text = .TextMatrix(intCurProd + 1, 3)
                    .Col = 4
                    .Text = .TextMatrix(intCurProd + 1, 4)
                    intCurProd = intCurProd + 1
                Wend
                .Rows = .Rows - 1
            ElseIf (mintProdNo = (.Rows - 1)) Then
                'if current product is the only record in the  grid
                'or if the current product is the last record in the grid
                .Rows = .Rows - 1
                mintProdNo = .Rows - 1
            End If
        End With
        'subtracting the amount of the product deleted from the total amount of the bill
        txtAmount.Text = Int(txtAmount.Text) - Int(txtProdAmount.Text)
        If (mintProdNo = 0) Then
            'if there is no product in the grid
            'disabling product transaction buttons
            cmdProdModify.Enabled = False
            cmdProdDelete.Enabled = False
            'clearing the product text controls
            Call ClearingProdText
        Else
            'if there is products in the grid
            'displaying the current product
            Call DisplayingProdDetails
        End If
        'enabling the product navigation buttons
        Call ProdRecordPosition
        'enabling product transaction buttons
        Call ProdEnableControls
    End If
End Sub

Private Sub cmdProdReset_Click()
    'resetting product text controls
    Call ResettingControls
    txtProdAccCode.SetFocus
End Sub
Private Sub cmdProdCancel_Click()
    If mintProdNo <> 0 Then
        'if there are products in the grid
        'displaying original product details
        Call DisplayingProdDetails
    Else
        'if there are no products in the grid
        'clearing product text controls
        Call ResettingControls
    End If
    'resetting product transaction flag
    mblnProd = False
    mstrProdTransType = ""
    'unlocking product text controls
    Call LockTextControls
    'enabling transaction buttons
    Call EnableControls
    Call ProdEnableControls
    'enabling product navigation buttons
    Call ProdRecordPosition
    'enabling spreadsheetview tab
    sstabRegister.TabEnabled(1) = True
    cmdProdAccCodeSel.Visible = False
    cmdProdNarrationSel.Visible = False
    
End Sub

Private Sub cmdProdSave_Click()
    'checking product text controls
    If (CheckProdDet) Then
        'if all the necessary details have been provided
        If (mstrProdTransType = "Add") Then
            'in case of adding a product
            'adding new row to the grid
            msflxgrdRegister.Rows = msflxgrdRegister.Rows + 1
            mintProdNo = msflxgrdRegister.Rows - 1
            If txtAmount.Text <> "" Then
                'if bill amount is not empty
                'adding the product amount to the bill amount
                txtAmount.Text = Int(txtAmount.Text) + Int(txtProdAmount.Text)
            Else
                'if bill amount is empty
                'setting the product amount as the bill amount
                txtAmount.Text = txtProdAmount.Text
            End If
        ElseIf (mstrProdTransType = "Modify") Then
            'in case of modifying the bill
            'subtracting the original amount from the bill amount
            txtAmount.Text = Int(txtAmount.Text) - Int(msflxgrdRegister.TextMatrix(mintProdNo, 4))
            'adding the new amount to the bill amount
            txtAmount.Text = Int(txtAmount.Text) + Int(txtProdAmount.Text)
        End If
        'displaying the current product details in the text fields
        With msflxgrdRegister
            .Row = mintProdNo
            .Col = 0
            .Text = txtProdAccCode.Text
            .Col = 1
            .Text = txtProdAccDesc.Text
            .Col = 2
            If (txtProdNarration.Text <> "") Then
                .Text = txtProdNarration.Text
            End If
            .Col = 3
            .Text = txtProdQuantity.Text
            .Col = 4
            .Text = txtProdAmount.Text
        End With
        'resetting the product transaction flag
        mstrProdTransType = ""
        mblnProd = False
        'enabling transaction buttons
        Call EnableControls
        Call ProdEnableControls
        'locking product text controls
        Call LockTextControls
        'enabling the product navigation buttons
        Call ProdRecordPosition
        'enabling spreadsheetview tab
        sstabRegister.TabEnabled(1) = True
        cmdProdAccCodeSel.Visible = False
        cmdProdNarrationSel.Visible = False
    End If
End Sub

Private Sub ResettingControls()
         
    If (Not mblnProd) Then
        'if in the bill transaction mode
        If (mstrTransType = "Add") Then
            'in case of adding bill
            'setting new billno and date to the text controls
            txtBillNo.Text = mintNewBillNo
            txtDate.Text = mstrNewDate
            txtAccCode.Text = ""
            txtAccDesc.Text = ""
            txtNarration.Text = ""
            txtAmount.Text = "0"
            'clearing the product text controls
            Call ClearingProdText
            'clearing the grid control
            msflxgrdRegister.Rows = 1
        ElseIf (mstrTransType = "Modify") Then
            'in case of modifying the bill
            'displaying original details
            Call DisplayingDetails
            Call DisplayingProdDetails
        End If
    Else
        'if in the product transaction mode
        If (mstrProdTransType = "Add") Then
            'in case of adding product
            'clearing product text controls
            Call ClearingProdText
        ElseIf (mstrProdTransType = "Modify") Then
            'in case of modifying product
            'displaying original product details
            Call DisplayingProdDetails
        End If
    End If
End Sub
Private Sub cmdProdAdd_Click()
    'checking bill details
    If (FieldCheck) Then
        'if all the necessary bill details are provided
        'setting product transaction flag
        mblnProd = True
        mstrProdTransType = "Add"
        'unlocking product text controls
        Call LockTextControls
        'clearing product text controls
        Call ResettingControls
        'disabling product transaction controls
        Call ProdEnableControls
        'disabling product navigation controls
        Call DisableProdNaviButtons
        'disabling spreadsheetview tab
        sstabRegister.TabEnabled(1) = False
        cmdProdAccCodeSel.Visible = True
        cmdProdNarrationSel.Visible = True
        txtProdAccCode.SetFocus
    End If
End Sub

Private Sub Form_Load()
'condition for opening the recordset
Dim strCondition As String

    If (gstrRegister = "Cash Sales") Then
        Call PreparingForm
    End If
    'disabling the grid
    msflxgrdRegister.Enabled = False
    'resetting all transaction flags
    mblnTrans = False
    mblnProd = False
    mblnFind = False
    mstrTransType = ""
    mstrProdTransType = ""
    
    'hiding account selection and narration selection buttons
    cmdAccCodeSel.Visible = False
    cmdProdAccCodeSel.Visible = False
    cmdNarrationSel.Visible = False
    cmdProdNarrationSel.Visible = False
    'locking bill amount and product account description text controls
    txtAmount.Locked = True
    txtAccDesc.Locked = True
    txtProdAccDesc.Locked = True
    
    'displaying grid heading
    Call DisplayingHeading
    'setting condition for opening recordset
    If gstrRegister = "Cash Sales" Then
        'frmCashRegister.Caption = "Cash Sales Register"
        strCondition = "Select * from Cash where (BillNo <> 0) Order by BillNo,EntryNo"
    ElseIf gstrRegister = "Credit Sales" Then
        'frmCashRegister.Caption = "Cash Purchase Register"
        strCondition = "Select * from Journal where (BillNo <> 0) Order by BillNo,EntryNo"
    End If
    'opening the recordset
    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    Set mrstRegister = mdbsAccounts.OpenRecordset(strCondition, dbOpenDynaset)
    If Not mrstRegister.BOF And Not mrstRegister.EOF Then
        'in case of bill already present
        mrstRegister.MoveLast
        mrstRegister.MoveFirst
        'displaying first bill details
        mintBillNo = mrstRegister!BillNo
        Call DisplayingDetails
        Call DisplayingProdDetails
    Else
        'in case if no bills exists
        mintBillNo = 0
        mintProdNo = 0
        'clearing grid control
        msflxgrdRegister.Rows = 1
    End If

    'enabling navigation buttons
    Call RecordPosition
    Call ProdRecordPosition
    'enabling transaction buttons
    Call EnableControls
    Call ProdEnableControls
    'locking text controls
    Call LockTextControls
    'showing form view tab
    sstabRegister.Tab = 0
    
End Sub
Private Sub DisplayingDetails()
'recordset for getting account description
Dim rstAccCode As Recordset
'condition for opening recordset
Dim strCondition As String
'condition for finding Account code
Dim strFindAccCode As String
'condition for finding Bill
Dim strFindBill As String
'for calculating bill amount
Dim curAmount As Currency

    'setting condition for opening recordset for obtaining account description
    strCondition = "Select Acc_Code,Acc_Desc from AccCode"
    'opening recordset
    Set rstAccCode = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not rstAccCode.BOF And Not rstAccCode.EOF) Then
        'if bills exists
        rstAccCode.MoveLast
        rstAccCode.MoveFirst
        If gstrRegister = "Credit Sales" Then
            'displaying personel details in case of credit sales
            strFindBill = "BillNo = " & mintBillNo & " and [Debit/Credit] = 'D'"
            mrstRegister.FindFirst strFindBill
            With mrstRegister
                txtAccCode.Text = !Acc_Code
                strFindAccCode = "Acc_code = '" & !Acc_Code & "'"
                rstAccCode.FindFirst strFindAccCode
                If rstAccCode.NoMatch Then
                    MsgBox "Account details doesnot exists"
                    txtAccDesc.Text = ""
                Else
                    txtAccDesc.Text = rstAccCode!Acc_Desc
                End If
                If IsNull(!Narration) Then
                    txtNarration.Text = ""
                Else
                    txtNarration.Text = !Narration
                End If
            End With
        ElseIf gstrRegister = "Cash Sales" Then
            strFindBill = "BillNo = " & mintBillNo
            mrstRegister.FindFirst strFindBill
        End If
        If Not mlnprod Then
            'in case not in product transaction mode
            'filling the Bill details

            txtBillNo.Text = mrstRegister!BillNo
            txtDate.Text = mrstRegister!Date
        End If
        curAmount = 0
        'filling product details in the grid
        strFindBill = "BillNo = " & mintBillNo & " and [Debit/Credit] = 'C'"
        mrstRegister.FindFirst strFindBill
        With msflxgrdRegister
            .Rows = 1
            Do While (Not mrstRegister.EOF)
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .Col = 0
                .Text = mrstRegister!Acc_Code
                strFindAccCode = "Acc_code = '" & mrstRegister!Acc_Code & "'"
                .Col = 1
                rstAccCode.FindFirst strFindAccCode
                If rstAccCode.NoMatch Then
                    MsgBox "Account details doesnot exists"
                    .Text = ""
                Else
                    .Text = rstAccCode!Acc_Desc
                End If
                .Col = 2
                If IsNull(mrstRegister!Narration) Then
                    .Text = ""
                Else
                    .Text = mrstRegister!Narration
                End If
                .Col = 3
                .Text = mrstRegister!Quantity
                .Col = 4
                .Text = mrstRegister!Amount
                curAmount = curAmount + mrstRegister!Amount
                mrstRegister.FindNext strFindBill
                If (mrstRegister.NoMatch) Then
                    Exit Do
                End If
            Loop
            'setting the calculated amount as the bill amount
            txtAmount.Text = curAmount
            mintProdNo = 1
        End With
    End If
End Sub
Private Sub DisplayingHeading()
    'setting the width and heading of the columns of the grid
    With msflxgrdRegister
        .FixedRows = 1
        .FixedCols = 0
        .Cols = 5
        .Row = 0
        .Col = 0
        .ColWidth(0) = 1300
        .Text = "Account Code"
        .Col = 1
        .ColWidth(1) = 3100
        .Text = "Description"
        .Col = 2
        .ColWidth(2) = 3500
        .Text = "Narration"
        .Col = 3
        .ColWidth(3) = 1000
        .Text = "Quantity"
        .Col = 4
        .ColWidth(4) = 1000
        .Text = "Amount"
    End With
End Sub

Private Sub DisplayingProdDetails()
    'displaying current product details in the product text controls
    With msflxgrdRegister
        .Row = mintProdNo
        .Col = 0
        txtProdAccCode.Text = .Text
        .Col = 1
        txtProdAccDesc.Text = .Text
        .Col = 2
        txtProdNarration.Text = .Text
        .Col = 3
        txtProdQuantity.Text = .Text
        .Col = 4
        txtProdAmount.Text = .Text
    End With
End Sub
Private Sub LockTextControls()

    If (mblnTrans) And (Not mblnProd) Then
        'in case of bill transaction mode and not in product transaction mode
        'locking billno and date text controls
        txtBillNo.Locked = False
        txtDate.Locked = False
        txtAccCode.Locked = False
        txtNarration.Locked = False
    Else
        'unlocking billno and date text controls
        txtBillNo.Locked = True
        txtDate.Locked = True
        txtAccCode.Locked = True
        txtNarration.Locked = True
    End If
    'locking or unlocking the product text controls according to transaction mode
    txtProdAccCode.Locked = Not mblnProd
    txtProdNarration.Locked = Not mblnProd
    txtProdQuantity.Locked = Not mblnProd
    txtProdAmount.Locked = Not mblnProd
End Sub
Private Sub RecordPosition()
'recordset for getting the unique billnos
Dim rstBillNo As Recordset
'condition for opening the recordset
Dim strCondition As String
'condition for finding the bill
Dim strFindCond As String

    'getting unique BillNos from cash table
    If gstrRegister = "Cash Sales" Then
        strCondition = "Select distinct BillNo from Cash where (BillNo <> 0) Order by BillNo"
    ElseIf gstrRegister = "Credit Sales" Then
        strCondition = "Select distinct BillNo from Journal where (BillNo <> 0)  Order by BillNo"
    End If
    'opening recordset
    Set rstBillNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not rstBillNo.BOF And Not rstBillNo.EOF) Then
        rstBillNo.MoveLast
        rstBillNo.MoveFirst
        'finding the current bill
        strFindCond = "BillNo = " & mintBillNo
        rstBillNo.FindFirst strFindCond
        If rstBillNo.NoMatch Then
            MsgBox "Bill No. does not exist", vbOKOnly
        Else
            'disabling the navigation buttons according to the position of the current bill
            With rstBillNo
                If .RecordCount = 0 Or .RecordCount = 1 Then
                    'in case of no bill or only one bill
                    cmdFirst.Enabled = False
                    cmdPrev.Enabled = False
                    cmdNext.Enabled = False
                    cmdLast.Enabled = False

                    mnuFirst.Enabled = False
                    mnuPrev.Enabled = False
                    mnuNext.Enabled = False
                    mnuLast.Enabled = False
                ElseIf .AbsolutePosition = 0 Then
                    'in case current bill is the first bill
                    cmdFirst.Enabled = False
                    cmdPrev.Enabled = False
                    cmdNext.Enabled = True
                    cmdLast.Enabled = True

                    mnuFirst.Enabled = False
                    mnuPrev.Enabled = False
                    mnuNext.Enabled = True
                    mnuLast.Enabled = True
                ElseIf .AbsolutePosition = (.RecordCount - 1) Then
                    'in case current bill is the last bill
                    cmdFirst.Enabled = True
                    cmdPrev.Enabled = True
                    cmdNext.Enabled = False
                    cmdLast.Enabled = False

                    mnuFirst.Enabled = True
                    mnuPrev.Enabled = True
                    mnuNext.Enabled = False
                    mnuLast.Enabled = False
                Else
                    'in case current bill is in the middle
                    cmdFirst.Enabled = True
                    cmdPrev.Enabled = True
                    cmdNext.Enabled = True
                    cmdLast.Enabled = True

                    mnuFirst.Enabled = True
                    mnuPrev.Enabled = True
                    mnuNext.Enabled = True
                    mnuLast.Enabled = True
                End If
            End With
        End If
    Else
        'in case no bills
        cmdFirst.Enabled = False
        cmdPrev.Enabled = False
        cmdNext.Enabled = False
        cmdLast.Enabled = False
        
        mnuFirst.Enabled = False
        mnuPrev.Enabled = False
        mnuNext.Enabled = False
        mnuLast.Enabled = False
    End If
End Sub
Private Sub ProdRecordPosition()
    
    With msflxgrdRegister
        If .Rows = 1 Or .Rows = 2 Then
            'if there are no products
            'or  there is only one product
            cmdProdFirst.Enabled = False
            cmdProdPrev.Enabled = False
            cmdProdNext.Enabled = False
            cmdProdLast.Enabled = False
        ElseIf mintProdNo = 1 Then
            'if the current product is the first one
            cmdProdFirst.Enabled = False
            cmdProdPrev.Enabled = False
            cmdProdNext.Enabled = True
            cmdProdLast.Enabled = True
        ElseIf mintProdNo = .Rows - 1 Then
            'if the current product is the last one
            cmdProdFirst.Enabled = True
            cmdProdPrev.Enabled = True
            cmdProdNext.Enabled = False
            cmdProdLast.Enabled = False
        Else
            'if the current product is in the middle
            cmdProdFirst.Enabled = True
            cmdProdPrev.Enabled = True
            cmdProdNext.Enabled = True
            cmdProdLast.Enabled = True
        End If
    End With
End Sub

Private Sub EnableControls()
'recordset for getting the unique billno
Dim rstBillNo As Recordset
'condition for opening the recordset
Dim strCondition As String

    If (Not mblnTrans) And (Not mblnFind) Then
        'in case not in the bill transaction mode or not in the find mode
        cmdAdd.Enabled = True
        cmdClose.Enabled = True
        mnuClose.Enabled = True
        'getting unique billno from cash
        If gstrRegister = "Cash Sales" Then
            strCondition = "select Distinct BillNo from Cash where (BillNo <> 0) order by BillNo"
        ElseIf gstrRegister = "Credit Sales" Then
            strCondition = "select Distinct BillNo from Journal where (BillNo <> 0) Order By BillNo"
        End If
        'opening recordset
        Set rstBillNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
        If (Not rstBillNo.BOF And Not rstBillNo.EOF) Then
            'if some data exist
            rstBillNo.MoveLast
            rstBillNo.MoveFirst
            cmdModify.Enabled = True
            cmdDelete.Enabled = True
            If rstBillNo.RecordCount = 1 Then
                'if there is only one bill
                cmdFind.Enabled = False
            ElseIf rstBillNo.RecordCount > 1 Then
                'if there are more bill
                cmdFind.Enabled = True
            End If
        Else
            'if no bill exists
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
            cmdFind.Enabled = False
        End If
        cmdSave.Enabled = False
        cmdReset.Enabled = False
        cmdCancel.Enabled = False
    ElseIf (mblnTrans) Or (mblnFind) Then
        'in case bill transation mode or in find mode
        cmdAdd.Enabled = False
        cmdModify.Enabled = False
        cmdDelete.Enabled = False
        cmdClose.Enabled = False
        If (mblnTrans) And (Not mblnProd) Then
            'in case bill transaction mode and not in product transaction mode
            cmdSave.Enabled = True
            cmdReset.Enabled = True
        Else
            cmdSave.Enabled = False
            cmdReset.Enabled = False
        End If
        If (mblnFind) Then
            'in case finding bill mode
            cmdFind.Enabled = True
        Else
            cmdFind.Enabled = False
        End If
        cmdCancel.Enabled = True
    End If
End Sub
Private Sub ProdEnableControls()
    
    If (mblnProd) Then
        'in case in product transaction mode
        cmdProdAdd.Enabled = False
        cmdProdModify.Enabled = False
        cmdProdDelete.Enabled = False
        cmdProdSave.Enabled = True
        cmdProdReset.Enabled = True
        cmdProdCancel.Enabled = True
        cmdSave.Enabled = False
        cmdReset.Enabled = False
        cmdCancel.Enabled = False
    ElseIf (mblnTrans) Then
        'in case in bill transaction mode
        cmdProdAdd.Enabled = True
        If (msflxgrdRegister.Rows > 1) Then
            'one or more product exist
            cmdProdModify.Enabled = True
            cmdProdDelete.Enabled = True
        Else
            'if no product exist
            cmdProdModify.Enabled = False
            cmdProdDelete.Enabled = False
        End If
        cmdProdSave.Enabled = False
        cmdProdReset.Enabled = False
        cmdProdCancel.Enabled = False
    Else
        'in case not in bill transaction mode or in product transaction mode
        cmdProdAdd.Enabled = False
        cmdProdModify.Enabled = False
        cmdProdDelete.Enabled = False
        cmdProdSave.Enabled = False
        cmdProdReset.Enabled = False
        cmdProdCancel.Enabled = False
    End If
End Sub

Private Sub cmdFirst_Click()
'condition for opening the recordset
Dim strCondition As String
'recordset for getting unique billno
Dim rstBillNo As Recordset
    
    'getting unique Bill details
    If gstrRegister = "Cash Sales" Then
        strCondition = "select distinct BillNo from Cash where (BillNo <> 0) order by BillNo"
    ElseIf gstrRegister = "Credit Sales" Then
        strCondition = "select distinct BillNo from Journal where (BillNo <> 0) order by BillNo"
    End If
    'opening recordset
    Set rstBillNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not rstBillNo.BOF And Not rstBillNo.EOF) Then
        rstBillNo.MoveFirst
        'getting the first bill
        mintBillNo = rstBillNo!BillNo
    End If
    'displaying current bill details
    Call DisplayingDetails
    Call DisplayingProdDetails
    'enabling navigation buttons
    Call RecordPosition
    Call ProdRecordPosition
End Sub

Private Sub cmdLast_Click()
'condition for opening recordset
Dim strCondition As String
'recordset for getting bill details
Dim rstBillNo As Recordset
    
    'getting unique Bill details
    If gstrRegister = "Cash Sales" Then
        strCondition = "select distinct BillNo from Cash where (BillNo <> 0) order by BillNo"
    ElseIf gstrRegister = "Credit Sales" Then
        strCondition = "select distinct BillNo from Journal where (BillNo <> 0) order by BillNo"
    End If
    'opening recordset
    Set rstBillNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not rstBillNo.BOF And Not rstBillNo.EOF) Then
        rstBillNo.MoveLast
        'getting the last bill
        mintBillNo = rstBillNo!BillNo
    End If
    'dispalying current bill details
    Call DisplayingDetails
    Call DisplayingProdDetails
    'enabling navigation buttons
    Call RecordPosition
    Call ProdRecordPosition
End Sub

Private Sub cmdModify_Click()
    'setting bill transaction flag
    mblnTrans = True
    mstrTransType = "Modify"
    'getting the total no of products
    mintProdTotNo = msflxgrdRegister.Rows - 1
    'unlocking text controls
    Call LockTextControls
    'enabling transaction buttons
    Call EnableControls
    Call ProdEnableControls
    'disabling bill navigation buttons
    Call DisableNaviButtons
    txtBillNo.Locked = True
    cmdAccCodeSel.Visible = True
    cmdNarrationSel.Visible = True
    txtDate.SetFocus
End Sub

Private Sub cmdNext_Click()
'condition for opening the recordset
Dim strCondition As String
'condition for finding the bill
Dim strFindCond As String
'recordset for getting the bill
Dim rstBillNo As Recordset
    
    'getting unique Bill details
    If gstrRegister = "Cash Sales" Then
        strCondition = "select distinct BillNo from Cash where (BillNo <> 0) order by BillNo"
    ElseIf gstrRegister = "Credit Sales" Then
        strCondition = "select distinct BillNo from Journal where (BillNo <> 0) order by BillNo"
    End If
    'opening recordset
    Set rstBillNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not rstBillNo.BOF And Not rstBillNo.EOF) Then
        'finding the current bill
        strFindCond = "BillNo= " & mintBillNo
        rstBillNo.FindFirst strFindCond
        If rstBillNo.NoMatch Then
            MsgBox "No Bill Details", vbOKOnly
        Else
            'moving to the next bill
            rstBillNo.MoveNext
            mintBillNo = rstBillNo!BillNo
        End If
    End If
    'displaying current bill details
    Call DisplayingDetails
    Call DisplayingProdDetails
    'enabling navigation buttons
    Call RecordPosition
    Call ProdRecordPosition
End Sub

Private Sub cmdPrev_Click()
'condition for opening the recordset
Dim strCondition As String
'condition for finding the bill
Dim strFindCond As String
'recordset for getting the bill details
Dim rstBillNo As Recordset
    
    'getting unique Bill details
    If gstrRegister = "Cash Sales" Then
        strCondition = "select distinct BillNo from Cash where (BillNo <> 0) order by BillNo"
    ElseIf gstrRegister = "Credit Sales" Then
        strCondition = "select distinct BillNo from Journal where (BillNo <> 0) order by BillNo"
    End If
    'opening recordset
    Set rstBillNo = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not rstBillNo.BOF And Not rstBillNo.EOF) Then
        'finding the current bill
        strFindCond = "BillNo= " & mintBillNo
        rstBillNo.FindFirst strFindCond
        If rstBillNo.NoMatch Then
            MsgBox "No Bill Details", vbOKOnly
        Else
            'moving to the previous bill
            rstBillNo.MovePrevious
            mintBillNo = rstBillNo!BillNo
        End If
    End If
    'disaplying the current bill details
    Call DisplayingDetails
    Call DisplayingProdDetails
    'enabling navigation buttons
    Call RecordPosition
    Call ProdRecordPosition
End Sub
Private Sub cmdProdFirst_Click()
    'setting first product as the current product
    mintProdNo = 1
    'displaying current product details in product text controls
    Call DisplayingProdDetails
    'enabling product navigation buttons
    Call ProdRecordPosition
End Sub

Private Sub cmdProdLast_Click()
    'setting last product as current product
    mintProdNo = msflxgrdRegister.Rows - 1
    'displaying current product details in product text controls
    Call DisplayingProdDetails
    'enabling product navigation buttons
    Call ProdRecordPosition
End Sub

Private Sub cmdProdModify_Click()
    'setting product transaction flag
    mblnProd = True
    mstrProdTransType = "Modify"
    'unlocking product text controls for modification
    Call LockTextControls
    'disabliing product transaction buttons
    Call EnableControls
    Call ProdEnableControls
    'disabling product navigation buttons
    Call DisableProdNaviButtons
    'disabling spreadsheet view tab
    sstabRegister.TabEnabled(1) = False
    cmdProdAccCodeSel.Visible = True
    cmdProdNarrationSel.Visible = True
    txtProdAccCode.SetFocus
End Sub

Private Sub cmdProdNext_Click()
    'setting next product as current product
    mintProdNo = mintProdNo + 1
    'displaying current product details in product text controls
    Call DisplayingProdDetails
    'enabling product navigation buttons
    Call ProdRecordPosition
End Sub

Private Sub cmdProdPrev_Click()
    'setting previous product as current product
    mintProdNo = mintProdNo - 1
    'displaying current product details in product text controls
    Call DisplayingProdDetails
    'enabling product navigation buttons
    Call ProdRecordPosition
End Sub

Private Function FieldCheck() As Boolean
    'setting checking flag false
    FieldCheck = False
    If (Not mblnProd) Then
        'in case not in product transaction mode
        If txtBillNo.Text = "" Then
            MsgBox "Bill No should not be blank", vbOKOnly
            txtBillNo.SetFocus
        ElseIf txtDate.Text = "" Then
            MsgBox "Date should not blank"
            txtDate.SetFocus
        ElseIf txtAccCode.Text = "" And (gstrRegister = "Credit Sales") Then
            MsgBox "Personal Account code should not be blank", vbOKOnly
            txtAccCode.SetFocus
        Else
            'if all necessary details have been provided
            'setting checking flag true
            FieldCheck = True
        End If
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    'selecting close button
    cmdClose.Value = True
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
        'in case in product transaction mode
        'allowing alphabets, numbers and backspace
        If (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            Beep
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Len(txtAccCode.Text) = 0 Then
                If KeyAscii <> Asc("A") And KeyAscii <> Asc("I") Then
                    MsgBox "First letter must begin with E,I"
                    KeyAscii = 0
                End If
            End If
        End If
    End If
End Sub

Private Sub txtAccCode_LostFocus()
'condition for opening recordset
Dim strCondition As String
'recordset for getting the description of the product code
Dim rstCode As Recordset
    
    If mblnTrans Then
        'if in product transation mode
        If txtAccCode.Text <> "" Then
            'getting the description of the entered product
            strCondition = "Select Acc_Desc from AccCode where Acc_Code = '" & txtAccCode.Text & "'"
            Set rstCode = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
            If rstCode.EOF Then
                MsgBox "Invalid Account Code"
                txtAccDesc.Text = ""
                txtAccCode.SetFocus
            Else
                'displaying the description
                txtAccDesc.Text = rstCode!Acc_Desc
            End If
        ElseIf txtAccCode.Text = "" Then
            'setting description empty if acccode is empty
            txtAccDesc.Text = ""
        End If
    End If

End Sub

Private Sub txtBillNo_KeyPress(KeyAscii As Integer)
    If mblnTrans Then
        'in case in bill transaction mode
        'allowing only numbers and backspace
        If (KeyAscii <> 8) And (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
           Beep
           KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If mblnTrans Then
        'in case in bill transaction mode
        'allowing only numbers ,/, and backspace
        If (KeyAscii <> 8) And (KeyAscii <> Asc("/")) And _
        (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
           Beep
           KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtNarration_KeyPress(KeyAscii As Integer)
'condition for opening the recordset
Dim strCondition As String
'recordset for getting the narration
Dim rstNarr As Recordset

    If mblnTrans Then
        'in case in product transaction mode
        'restricting datas
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
            'getting the narration
            strCondition = "Select Narr_Desc from NarrCode where Narr_Code  = '" & txtNarration.Text & " '"
            Set rstNarr = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
            If rstNarr.EOF Then
                MsgBox "Invalid Narration"
                txtNarration.SetFocus
            Else
                txtNarration.Text = rstNarr!Narr_Desc
                txtNarration.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtProdAccCode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 And mblnProd Then
        'in case in product transaction mode
        'allowing alphabets, numbers and backspace
        If (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            Beep
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Len(txtProdAccCode.Text) = 0 Then
                If KeyAscii <> Asc("E") And KeyAscii <> Asc("I") Then
                    MsgBox "First letter must begin with E,I"
                    KeyAscii = 0
                End If
            End If
        End If
    End If
End Sub
Private Sub txtProdAccCode_LostFocus()
'condition for opening recordset
Dim strCondition As String
'recordset for getting the description of the product code
Dim rstCode As Recordset
    
    If mblnProd Then
        'if in product transation mode
        If txtProdAccCode.Text <> "" Then
            'getting the description of the entered product
            strCondition = "Select Acc_Desc from AccCode where Acc_Code = '" & txtProdAccCode.Text & "'"
            Set rstCode = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
            If rstCode.EOF Then
                MsgBox "Invalid Account Code"
                txtProdAccDesc.Text = ""
                txtProdAccCode.SetFocus
            Else
                'displaying the description
                txtProdAccDesc.Text = rstCode!Acc_Desc
            End If
        ElseIf txtProdAccCode.Text = "" Then
            'setting description empty if acccode is empty
            txtProdAccDesc.Text = ""
        End If
    End If
End Sub


Private Sub txtProdAmount_KeyPress(KeyAscii As Integer)
    If mblnProd Then
        'in case in product transaction mode
        'allowing only numbers,., and backspace
        If (KeyAscii <> 8) And (KeyAscii <> Asc(".")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtProdNarration_KeyPress(KeyAscii As Integer)
'condition for opening the recordset
Dim strCondition As String
'recordset for getting the narration
Dim rstNarr As Recordset

    If mblnProd Then
        'in case in product transaction mode
        'restricting datas
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
            'getting the narration
            strCondition = "Select Narr_Desc from NarrCode where Narr_Code  = '" & txtProdNarration.Text & " '"
            Set rstNarr = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
            If rstNarr.EOF Then
                MsgBox "Invalid Narration"
                txtProdNarration.SetFocus
            Else
                txtProdNarration.Text = rstNarr!Narr_Desc
                txtProdNarration.SetFocus
            End If
        End If
    End If
End Sub
Private Sub txtProdQuantity_KeyPress(KeyAscii As Integer)
    If mblnProd Then
        'in case in product transaction mode
        'allowing only numbers,.,-, and backspace
        If (KeyAscii <> 8) And (KeyAscii <> Asc(".")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
            (KeyAscii <> Asc("-")) Then
            KeyAscii = 0
            Beep
        End If
    End If
End Sub
