VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6105
   ClientLeft      =   105
   ClientTop       =   780
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9285
   Begin MSFlexGridLib.MSFlexGrid msflxgrdLedger 
      Height          =   4455
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7858
      _Version        =   393216
   End
   Begin VB.TextBox txtAccCode 
      Height          =   300
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   960
   End
   Begin VB.TextBox txtAccDesc 
      Height          =   300
      Left            =   3660
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   3330
   End
   Begin VB.TextBox txtDate 
      Height          =   300
      Left            =   7920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   1170
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F&ind"
      Height          =   375
      Left            =   6405
      TabIndex        =   2
      Tag             =   "Find"
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7485
      TabIndex        =   1
      Top             =   780
      Width           =   975
   End
   Begin VB.ComboBox cboAccCode 
      Height          =   315
      Left            =   1365
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   975
   End
   Begin VB.Label lblLedger 
      Caption         =   "Account Code"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   390
      Width           =   1140
   End
   Begin VB.Label lblLedger 
      Caption         =   "Description"
      Height          =   255
      Index           =   1
      Left            =   2610
      TabIndex        =   7
      Top             =   390
      Width           =   885
   End
   Begin VB.Label lblLedger 
      Caption         =   "Date"
      Height          =   255
      Index           =   2
      Left            =   7215
      TabIndex        =   6
      Top             =   390
      Width           =   525
   End
   Begin VB.Menu mnuForm 
      Caption         =   "&Form"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
         Shortcut        =   ^C
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
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    Dim mdbsAccounts As Database
    Dim mrstAccCode As Recordset
    Dim mstrStLtAcccode As String
    Dim mstrCurDate As String

    Dim mblnDiffCodes As Boolean
    Dim mblnDiffDates As Boolean
    Dim mvntBkMark As Variant
    Dim mcurDb As Currency
    Dim mcurCr As Currency
    Dim mcurQty As Currency
Private Sub cmdCancel_Click()
    If (Not mrstAccCode.EOF) And (Not mrstAccCode.BOF) Then
        mrstAccCode.Bookmark = mvntBkMark
    End If
    cmdFind.Caption = "F&ind"
    cmdFind.Tag = "Find"
    cmdCancel.Enabled = False
    mnuPrint.Enabled = True
    mnuClose.Enabled = True
    Call RecordPosition
    cboAccCode.Visible = False
    txtAccCode.Visible = True
    msflxgrdLedger.Enabled = True
End Sub
Private Sub cmdFind_Click()
    If (cmdFind.Tag = "Find") Then
        cmdFind.Caption = "&Display"
        cmdFind.Tag = "Display"
        cmdCancel.Enabled = True
        mnuPrint.Enabled = False
        mnuClose.Enabled = False
        mnuFirst.Enabled = False
        mnuPrev.Enabled = False
        mnuNext.Enabled = False
        mnuLast.Enabled = False
        msflxgrdLedger.Enabled = False
        
        txtAccCode.Visible = False
        cboAccCode.Visible = True
        cboAccCode.Clear
        If (Not mrstAccCode.EOF) And (Not mrstAccCode.BOF) Then
            mvntBkMark = mrstAccCode.Bookmark
            mrstAccCode.MoveFirst
        End If
        While Not mrstAccCode.EOF
            cboAccCode.AddItem mrstAccCode!Acc_Code
            mrstAccCode.MoveNext
        Wend
        mrstAccCode.MoveFirst
        cboAccCode.SetFocus
    Else
        If (cboAccCode.ListIndex = -1) Then
            MsgBox "Specify the Account Code", vbOKOnly
            cboAccCode.SetFocus
        Else
            mrstAccCode.FindFirst "Acc_Code='" & cboAccCode.Text & "'"
            Call PreparingDetails
            Call RecordPosition
            cmdFind.Caption = "F&ind"
            cmdFind.Tag = "Find"
            cmdCancel.Enabled = False
            cboAccCode.Visible = False
            txtAccCode.Visible = True
            mnuPrint.Enabled = True
            mnuClose.Enabled = True
            msflxgrdLedger.Enabled = True
        End If
    End If
End Sub
Private Sub Form_Load()

    'condition for opening Account Details
    Dim strCodeCond As String
    'Starting AccCode
    Dim strStCode As String
    'End AccCode
    Dim strEndCode As String
    
    'openning the database
    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    'Adding * in the Account code to search for the specified pattern
    strStCode = gstrStCode & "*"
    strEndCode = gstrEndCode & "*"
    'getting the account codes lying between the specified pattern
    strCodeCond = "Select Acc_Code, Acc_Desc, Open_Bal, Bal_Type, YTop_Qty  from AccCode Where " _
            & "(Acc_Code Like '" & strStCode & "') or (Acc_Code Like '" _
            & strEndCode & "') or (Acc_Code >= '" & gstrStCode _
            & "' and Acc_Code <= '" & gstrEndCode & "') Order by Acc_Code "
    Set mrstAccCode = mdbsAccounts.OpenRecordset(strCodeCond, dbOpenSnapshot)
    If (Not mrstAccCode.BOF And Not mrstAccCode.EOF) Then
        'if there are Accounts for the specified pattern
        mrstAccCode.MoveLast
        mrstAccCode.MoveFirst
        'checking whether ledger of more accounts is requisted
        If mrstAccCode.RecordCount > 1 Then
            'if there are more accounts setting navigation accordingly
            cmdFind.Enabled = True
            mnuNavigation.Visible = True
            Call RecordPosition
        Else
            'if ther is only one account hidding navigation
            cmdFind.Enabled = False
            mnuNavigation.Visible = False
        End If
        
        'checking whether the ledger details of more than 1 date is requiested
        mblnDiffDates = False
        If gstrStDate <> gstrEndDate Then
            'if transactions of more dates are required
            'setting different date flag true
            mblnDiffDates = True
            'hiding the date field
            lblLedger(2).Visible = False
            txtDate.Visible = False
            'adjusting the position of the AccCode and AccDesc field
            txtAccDesc.Width = txtAccDesc.Width + 1000
            lblLedger(0).Left = lblLedger(0).Left + 250
            txtAccCode.Left = txtAccCode.Left + 250
            lblLedger(1).Left = lblLedger(1).Left + 350
            txtAccDesc.Left = txtAccDesc.Left + 350
        End If
        'getting necessary information and displaying the details
        Call PreparingDetails
    Else
        'if no accounts on the specified patterns exist
        mblnDiffDates = False
        Call Heading
        Call CalPrevBal
        MsgBox "No Accounts in specified pattern", vbOKOnly
    End If
    
    'hiding the combo box used in find mode
    cboAccCode.Visible = False
    'disabling the cancel button
    cmdCancel.Enabled = False
End Sub
Private Sub CalPrevBal()
    
    'recordset for cash
    Dim rstCash As Recordset
    'recordset for journal
    Dim rstJournal As Recordset
    'condition for opening recordset
    Dim strCond As String
    'current column no in grid
    Dim intCol As Integer
    'difference between Debit and Credit amount
    Dim curBal As Currency
    
    'getting the openbalance and year top quantity from AccCode
    mcurDb = 0
    mcurCr = 0
    mcurQty = 0
    If Not IsNull(mrstAccCode!Open_Bal) Then
        If mrstAccCode!Bal_Type = "D" Then
            mcurDb = Format(mcurDb, "######.00") + mrstAccCode!Open_Bal
        ElseIf mrstAccCode!Bal_Type = "C" Then
            mcurCr = Format(mcurCr, "######.00") + mrstAccCode!Open_Bal
        End If
    End If
    If Not IsNull(mrstAccCode!YTop_Qty) Then
        mcurQty = Format(mcurQty, "####.00") + mrstAccCode!YTop_Qty
    End If
    
    'getting the details of transaction of the account before the starting date from Cash
    strCond = "Select [Debit/Credit], Quantity, Amount from Cash Where " _
                & "Acc_Code = '" & mrstAccCode!Acc_Code & "' and Date < CDate('" _
                & gstrStDate & "') Order by Date "
    Set rstCash = mdbsAccounts.OpenRecordset(strCond, dbOpenSnapshot)
    'calculating the openbalance from Cash
    While Not rstCash.EOF
        If rstCash![Debit/Credit] = "D" Then
            mcurDb = Format(mcurDb, "######.00") + rstCash!Amount
        ElseIf rstCash![Debit/Credit] = "C" Then
            mcurCr = Format(mcurCr, "######.00") + rstCash!Amount
        End If
        If Not IsNull(rstCash!Quantity) Then
            mcurQty = Format(mcurQty, "####.00") + rstCash!Quantity
        End If
        rstCash.MoveNext
    Wend
    
    'getting the details of transaction of the account before the starting date from Journal
    strCond = "Select [Debit/Credit], Quantity, Amount from journal Where " _
                & "Acc_Code = '" & mrstAccCode!Acc_Code & "' and Date < CDate('" _
                & gstrStDate & "') Order by Date "
    Set rstJournal = mdbsAccounts.OpenRecordset(strCond, dbOpenSnapshot)
    'calculating the openbalance from journal
    While Not rstJournal.EOF
        If rstJournal![Debit/Credit] = "D" Then
            mcurDb = Format(mcurDb, "######.00") + rstJournal!Amount
        ElseIf rstJournal![Debit/Credit] = "C" Then
            mcurCr = Format(mcurCr, "######.00") + rstJournal!Amount
        End If
        If Not IsNull(rstJournal!Quantity) Then
            mcurQty = Format(mcurQty, "####.00") + rstJournal!Quantity
        End If
        rstJournal.MoveNext
    Wend
    
    'displaying openbalance in the grid
    With msflxgrdLedger
        .Rows = 2
        If mblnDiffDates Then
            intCol = 1
        Else
            intCol = 0
        End If
        .TextMatrix(1, intCol) = "Opening Balance"
        intCol = intCol + 1
        If mcurQty <> 0# Then
            .TextMatrix(1, intCol) = Format(mcurQty, "#####.00")
        End If
        If mcurDb > mcurCr Then
            'if debit is more than credit
            'displaying the difference
            'setting the difference as the debit
            intCol = intCol + 1
            curBal = Format(mcurDb, "######.00") - Format(mcurCr, "######.00")
            .TextMatrix(1, intCol) = Format(curBal, "######.00")
            mcurDb = Format(curBal, "######.00")
            mcurCr = 0#
        ElseIf mcurCr > mcurDb Then
            'if credit is more than debit
            'dispalying the difference
            'setting the difference as the credit
            intCol = intCol + 2
            curBal = Format(mcurCr, "######.00") - Format(mcurDb, "######.00")
            .TextMatrix(1, intCol) = Format(curBal, "######.00")
            mcurCr = Format(curBal, "######.00")
            mcurDb = 0#
        Else
            'if credit is equal to debit
            'dispalying nill balance
            'setting the debit and credit as 0
            mcurDb = 0#
            mcurCr = 0#
            intCol = intCol - 1
            .TextMatrix(1, intCol) = "Opening Balance (Nill)"
        End If
    End With
End Sub
Private Sub DisplayingDetails(rstTrans As Recordset)
Dim curQty As Currency
Dim curAmount As Currency
Dim intCol As Integer
Dim intRow As Integer
Dim blnCreditFlag As Boolean
          
blnCreditFlag = False

    If (mstrStLtAcccode = "I") And (rstTrans![Debit/Credit] = "C") And (mrstAccCode!Acc_Code <> "IPAL") Then
         curQty = 0
         curAmount = 0
         Do While (CDate(mstrCurDate) = CDate(rstTrans!Date) And (mstrStLtAcccode = "I") And (rstTrans![Debit/Credit] = "C"))
            blnCreditFlag = True
            If rstTrans!Quantity <> 0# Then
                curQty = Format(curQty, "####.00") + rstTrans!Quantity
            End If
            curAmount = Format(curAmount, "######.00") + rstTrans!Amount
             rstTrans.MoveNext
             If (rstTrans.EOF) Then
                 Exit Do
             End If
         Loop
         rstTrans.MovePrevious
     End If
     With msflxgrdLedger
         .Rows = .Rows + 1
         intRow = .Rows - 1
         intCol = 0
         If mblnDiffDates Then
             .TextMatrix(intRow, intCol) = Format(rstTrans!Date, "dd-mm-yyyy")
             intCol = intCol + 1
         End If
         If blnCreditFlag Then
            intCol = intCol + 1
            If curQty <> 0# Then
                .TextMatrix(intRow, intCol) = Format(curQty, "####.00")
                mcurQty = Format(mcurQty, "####.00") + curQty
            End If
            intCol = intCol + 2
            .TextMatrix(intRow, intCol) = Format(curAmount, "######.00")
            mcurCr = mcurCr + curAmount
            'rstTrans.MovePrevious
         Else
            If Not IsNull(rstTrans!Narration) Then
             .TextMatrix(intRow, intCol) = rstTrans!Narration
            End If
            intCol = intCol + 1
             If Not IsNull(rstTrans!Quantity) Then
                 .TextMatrix(intRow, intCol) = Format(rstTrans!Quantity, "####.00")
                 mcurQty = Format(mcurQty, "####.00") + rstTrans!Quantity
             End If
             intCol = intCol + 1
             If rstTrans![Debit/Credit] = "D" Then
                 .TextMatrix(intRow, intCol) = Format(rstTrans!Amount, "######.00")
                 mcurDb = mcurDb + rstTrans!Amount
             End If
             intCol = intCol + 1
             If rstTrans![Debit/Credit] = "C" Then
                 .TextMatrix(intRow, intCol) = Format(rstTrans!Amount, "######.00")
                 mcurCr = mcurCr + rstTrans!Amount
            End If
         End If
     End With
End Sub
Private Sub Heading()
    
    'setting the heading of the grid control
    With msflxgrdLedger
        .Clear
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        If mblnDiffDates Then
            'in case of different dates displaying the date in the grid
            .Cols = 5
            .TextMatrix(0, 0) = "Date"
            .ColWidth(0) = 1200
            .TextMatrix(0, 1) = "Narration"
            .ColWidth(1) = 3450
            .TextMatrix(0, 2) = "Quantity"
            .ColWidth(2) = 1350
            .TextMatrix(0, 3) = "Debit"
            .ColWidth(3) = 1350
            .TextMatrix(0, 4) = "Credit"
            .ColWidth(4) = 1350
        Else
            'in case of single date displaying only the other details
            .Cols = 4
            .TextMatrix(0, 0) = "Narration"
            .ColWidth(0) = 4400
            .TextMatrix(0, 1) = "Quantity"
            .ColWidth(1) = 1500
            .TextMatrix(0, 2) = "Debit"
            .ColWidth(2) = 1500
            .TextMatrix(0, 3) = "Credit"
            .ColWidth(3) = 1500
        End If
    End With
End Sub
Private Sub PreparingDetails()
    
    Dim strCashCond As String
    Dim rstCash As Recordset
    Dim strJourCond As String
    Dim rstJournal As Recordset
    Dim blnCashTrans As Boolean
    Dim blnJournalTrans As Boolean
    Dim strFindCond As String
    
    'displaying Account details
    txtAccCode.Text = mrstAccCode!Acc_Code
    txtAccDesc.Text = mrstAccCode!Acc_Desc
    If Not mblnDiffDates Then
        txtDate.Text = gstrStDate
    End If
    'clearing grid control
    msflxgrdLedger.Clear
    'displaying Heading
    Call Heading
    'calculaing and displaying openbalance
    Call CalPrevBal
    
    'getting transaction details from cash table
    strCashCond = "Select Date, Narration, Quantity, [Debit/Credit], Amount " _
        & " from Cash where Acc_Code = '" & mrstAccCode!Acc_Code & "'" _
        & " and Date >= CDate('" & gstrStDate & "') and Date <= " _
        & " CDate('" & gstrEndDate & "') Order By Date,[Debit/Credit],EntryNo"
    Set rstCash = mdbsAccounts.OpenRecordset(strCashCond, dbOpenSnapshot)
    If Not rstCash.EOF Then
        rstCash.MoveLast
        rstCash.MoveFirst
        blnCashTrans = True
    End If

    'getting transaction details from journal table
    strJourCond = "Select Date, Narration, Quantity, [Debit/Credit], Amount " _
        & " from Journal Where Acc_Code = '" & mrstAccCode!Acc_Code & "'" _
        & " and Date >= CDate('" & gstrStDate & "') and Date <= CDate('" _
        & gstrEndDate & "') Order By Date,[Debit/Credit],EntryNo"
    Set rstJournal = mdbsAccounts.OpenRecordset(strJourCond, dbOpenSnapshot)
    If Not rstJournal.EOF Then
        rstJournal.MoveLast
        rstJournal.MoveFirst
        blnJournalTrans = True
    End If
    
    If (blnCashTrans Or blnJournalTrans) Then
        mstrCurDate = gstrStDate
        mstrStLtAcccode = Left(mrstAccCode!Acc_Code, 1)
        Do While (CDate(mstrCurDate) <= CDate(gstrEndDate))
            strFindCond = "Date = cdate('" & mstrCurDate & "')"
            If (blnJournalTrans) Then
                rstJournal.FindFirst strFindCond
                If (Not rstJournal.NoMatch) Then
                    Do While (CDate(mstrCurDate) = CDate(rstJournal!Date))
                        Call DisplayingDetails(rstJournal)
                        rstJournal.MoveNext
                        If rstJournal.EOF Then
                            blnJournalTrans = False
                            Exit Do
                        End If
                    Loop
                End If
            End If
            If (blnCashTrans) Then
                rstCash.FindFirst strFindCond
                If (Not rstCash.NoMatch) Then
                    Do While (CDate(mstrCurDate) = CDate(rstCash!Date))
                        Call DisplayingDetails(rstCash)
                        rstCash.MoveNext
                        If rstCash.EOF Then
                            blnCashTrans = False
                            Exit Do
                        End If
                    Loop
                End If
            End If
            mstrCurDate = DateAdd("d", 1, mstrCurDate)
        Loop
        Call DisplayingTotal
    ElseIf (Not blnCashTrans And Not blnJournalTrans) Then
        Call Heading
        Call CalPrevBal
        Call DisplayingTotal
        MsgBox "Transaction for the Account for the specified dates does not exist", vbOKOnly
    End If
End Sub
Private Sub RecordPosition()
    
    With mrstAccCode
        If (.EOF And .BOF) Or .RecordCount = 1 Then
            mnuFirst.Enabled = False
            mnuPrev.Enabled = False
            mnuNext.Enabled = False
            mnuLast.Enabled = False
        ElseIf .AbsolutePosition = 0 Then
            mnuFirst.Enabled = False
            mnuPrev.Enabled = False
            mnuNext.Enabled = True
            mnuLast.Enabled = True
        ElseIf .AbsolutePosition = (.RecordCount - 1) Then
            mnuFirst.Enabled = True
            mnuPrev.Enabled = True
            mnuNext.Enabled = False
            mnuLast.Enabled = False
        Else
            mnuFirst.Enabled = True
            mnuPrev.Enabled = True
            mnuNext.Enabled = True
            mnuLast.Enabled = True
        End If
    End With
End Sub
Private Sub FillingLedger()
    Dim dbsLedger As Database
    Dim rstLedger As Recordset
    Dim intCurRow As Integer
    Dim intCurCol As Integer
    Dim intTotRow As Integer
    
    Set dbsLedger = OpenDatabase(gstrDatabase)
    dbsLedger.Execute "Delete * from Ledger"
    Set rstLedger = dbsLedger.OpenRecordset("select * from Ledger", dbOpenDynaset)
    intCurRow = 1
    With msflxgrdLedger
        intTotRow = .Rows - 2
        While intCurRow < intTotRow
            rstLedger.AddNew
            intCurCol = 0
            If mblnDiffDates Then
                If (.TextMatrix(intCurRow, intCurCol) <> "") Then
                    rstLedger!Date = Format(CDate(.TextMatrix(intCurRow, intCurCol)), "dd/mm/yyyy")
                End If
                intCurCol = intCurCol + 1
            End If
            If (.TextMatrix(intCurRow, intCurCol) <> "") Then
                rstLedger!Narration = .TextMatrix(intCurRow, intCurCol)
            End If
            intCurCol = intCurCol + 1
            If (.TextMatrix(intCurRow, intCurCol) <> "") Then
                rstLedger!Quantity = CSng(.TextMatrix(intCurRow, intCurCol))
            End If
            intCurCol = intCurCol + 1
            If (.TextMatrix(intCurRow, intCurCol) <> "") Then
                rstLedger!Debit = CSng(.TextMatrix(intCurRow, intCurCol))
            End If
            intCurCol = intCurCol + 1
            If (.TextMatrix(intCurRow, intCurCol) <> "") Then
                rstLedger!Credit = CSng(.TextMatrix(intCurRow, intCurCol))
            End If
            rstLedger.Update
            intCurRow = intCurRow + 1
        Wend
    End With
    rstLedger.Close
    dbsLedger.Close
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub
Private Sub mnuFirst_Click()
    mrstAccCode.MoveFirst
    Call RecordPosition
    Call PreparingDetails
End Sub
Private Sub mnuLast_Click()
    mrstAccCode.MoveLast
    Call RecordPosition
    Call PreparingDetails
End Sub
Private Sub mnuNext_Click()
    mrstAccCode.MoveNext
    Call RecordPosition
    Call PreparingDetails
End Sub
Private Sub mnuPrev_Click()
    mrstAccCode.MovePrevious
    Call RecordPosition
    Call PreparingDetails
End Sub
Private Sub mnuPrint_Click()
    Dim intCurRow As Integer
    Call FillingLedger
    With msflxgrdLedger
        intCurRow = .Rows - 2
        If mblnDiffDates Then
            DataEnvironment1.DiffDateLedger
            If (gstrCompanyCode = "SNTC") Then
                rptDiffDateLedger.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI NARAYANA TRADING COMPANY"
            ElseIf (gstrCompanyCode = "VPPS") Then
                rptDiffDateLedger.Sections("PageHeader").Controls.Item("lblCompany").Caption = "V.P.PACHAIYAPPA MUDALIAR & SONS"
            Else
                rptDiffDateLedger.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI PACHAIYAPPA MARKETING AGENCIES"
            End If
            rptDiffDateLedger.Sections("PageHeader").Controls.Item("lblYear").Caption = "Account Year " & gstrAccYear
            rptDiffDateLedger.Sections("PageHeader").Controls.Item("lblYearDate").Caption = "Date: " _
                & Format(gstrStDate, "dd/mm/yyyy") & " to " & Format(gstrEndDate, "dd/mm/yyyy")
            rptDiffDateLedger.Sections("PageHeader").Controls.Item("lblAccount").Caption = "Account : " & txtAccDesc.Text
            If (Not IsNull(.TextMatrix(intCurRow, 2))) Then
                rptDiffDateLedger.Sections("ReportFooter").Controls.Item("lblTotalQuantity").Caption = .TextMatrix(intCurRow, 2)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 3))) Then
                rptDiffDateLedger.Sections("ReportFooter").Controls.Item("lblTotalDebit").Caption = .TextMatrix(intCurRow, 3)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 4))) Then
                rptDiffDateLedger.Sections("ReportFooter").Controls.Item("lblTotalCredit").Caption = .TextMatrix(intCurRow, 4)
            End If
            intCurRow = intCurRow + 1
            rptDiffDateLedger.Sections("ReportFooter").Controls.Item("lblBalance").Caption = .TextMatrix(intCurRow, 1)
            If (Not IsNull(.TextMatrix(intCurRow, 2))) Then
                rptDiffDateLedger.Sections("ReportFooter").Controls.Item("lblBalanceQuantity").Caption = .TextMatrix(intCurRow, 2)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 3))) Then
                rptDiffDateLedger.Sections("ReportFooter").Controls.Item("lblBalanceDebit").Caption = .TextMatrix(intCurRow, 3)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 4))) Then
                rptDiffDateLedger.Sections("ReportFooter").Controls.Item("lblBalanceCredit").Caption = .TextMatrix(intCurRow, 4)
            End If
        Else
            DataEnvironment1.SameDateLedger
            If (gstrCompanyCode = "SNTC") Then
                rptSameDateLedger.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI NARAYANA TRADING COMPANY"
            ElseIf (gstrCompanyCode = "VPPS") Then
                rptSameDateLedger.Sections("PageHeader").Controls.Item("lblCompany").Caption = "V.P.PACHAIYAPPA MUDHALIAR & SONS"
            Else
                rptSameDateLedger.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI PACHAIYAPPA MARKETING AGENCIES"
            End If
            rptSameDateLedger.Sections("PageHeader").Controls.Item("lblYear").Caption = "Account Year " & gstrAccYear
            rptSameDateLedger.Sections("PageHeader").Controls.Item("lblYearDate").Caption = "Date: " & Format(gstrStDate, "dd/mm/yyyy") _
                    & " to " & Format(gstrStDate, "dd/mm/yyyy")
            rptSameDateLedger.Sections("PageHeader").Controls.Item("lblAccount").Caption = "Account : " & txtAccDesc.Text
            If (Not IsNull(.TextMatrix(intCurRow, 1))) Then
                rptSameDateLedger.Sections("ReportFooter").Controls.Item("lblTotalQuantity").Caption = .TextMatrix(intCurRow, 1)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 2))) Then
                rptSameDateLedger.Sections("ReportFooter").Controls.Item("lblTotalDebit").Caption = .TextMatrix(intCurRow, 2)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 3))) Then
                rptSameDateLedger.Sections("ReportFooter").Controls.Item("lblTotalCredit").Caption = .TextMatrix(intCurRow, 3)
            End If
            intCurRow = intCurRow + 1
            rptSameDateLedger.Sections("ReportFooter").Controls.Item("lblBalance").Caption = .TextMatrix(intCurRow, 0)
            If (Not IsNull(.TextMatrix(intCurRow, 1))) Then
                rptSameDateLedger.Sections("ReportFooter").Controls.Item("lblBalanceQuantity").Caption = .TextMatrix(intCurRow, 1)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 2))) Then
                rptSameDateLedger.Sections("ReportFooter").Controls.Item("lblBalanceDebit").Caption = .TextMatrix(intCurRow, 2)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 3))) Then
                rptSameDateLedger.Sections("ReportFooter").Controls.Item("lblBalanceCredit").Caption = .TextMatrix(intCurRow, 3)
            End If
        End If
    End With
       
    If mblnDiffDates Then
        rptDiffDateLedger.Show vbModal
        DataEnvironment1.rsDiffDateLedger.Close
    Else
        rptSameDateLedger.Show vbModal
        DataEnvironment1.rsSameDateLedger.Close
    End If
End Sub
Private Sub DisplayingTotal()

    Dim intStCol As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim curBal As Currency
    If mblnDiffDates Then
        intStCol = 1
    Else
        intStCol = 0
    End If
    intCol = intStCol
    With msflxgrdLedger
        .Rows = .Rows + 1
        intRow = .Rows - 1
        .TextMatrix(intRow, intCol) = "Total"
        intCol = intCol + 1
        If mcurQty <> 0 Then
            .TextMatrix(intRow, intCol) = Format(mcurQty, "#####.00")
        End If
        intCol = intCol + 1
        If mcurDb <> 0 Then
            .TextMatrix(intRow, intCol) = Format(mcurDb, "#####.00")
        End If
        intCol = intCol + 1
        If mcurCr <> 0 Then
            .TextMatrix(intRow, intCol) = Format(mcurCr, "#####.00")
        End If
        .Rows = .Rows + 1
        intRow = .Rows - 1
        intCol = intStCol
        If mcurDb > mcurCr Then
            .TextMatrix(intRow, intCol) = "Balance"
             curBal = Format(mcurDb, "#####.00") - Format(mcurCr, "#####.00")
            .TextMatrix(intRow, intCol + 2) = Format(curBal, "#####.00")
        ElseIf mcurCr > mcurDb Then
            .TextMatrix(intRow, intCol) = "Balance"
            curBal = Format(mcurCr, "#####.00") - Format(mcurDb, "#####.00")
            .TextMatrix(intRow, intCol + 3) = Format(curBal, "#####.00")
        ElseIf mcurCr = mcurDb Then
            .TextMatrix(intRow, intCol) = "Balance (Nill)"
        End If
    End With

End Sub
