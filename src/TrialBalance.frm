VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmTrialBalance 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6030
   ClientLeft      =   240
   ClientTop       =   840
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8985
   Begin MSFlexGridLib.MSFlexGrid msflxgrdTrialBalance 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9340
      _Version        =   393216
   End
   Begin VB.Menu mnuForm 
      Caption         =   "&Form"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuWNB 
            Caption         =   "With Null Balance"
         End
         Begin VB.Menu mnWONB 
            Caption         =   "Without Null Balance"
         End
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmTrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mdbsAccounts As Database    'to open database
    Dim mrstAccCode As Recordset    'to open Account table
    Dim mintCurRow As Integer       'to store current row
    Dim mcurDBTot As Currency       'to store current debit total
    Dim mcurCRTot As Currency       'to store current credit total
    Dim mstrPrint As String         'to store the type of printing required
Private Sub Form_Load()
    Dim strCodeCond As String       'to store the condition for the AccCode table
    
    'openning the database
    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    'setting the variables to the required initial value
    mintCurRow = 1
    mcurDBTot = 0#
    mcurCRTot = 0#
    'framing the condition according to the trial balance required
    If frmMain.mnuAssLiab.Tag = "Prepare" Then
        'to list all assets and liablities
        strCodeCond = "Select Acc_Code, Acc_Desc, Open_Bal, Bal_Type, " _
                    & " Qty_ToRec, YTop_Qty from AccCode" _
                    & " Where Acc_Code Like '[A,L]*' Order By Acc_Desc"
    ElseIf frmMain.mnuNomAcc.Tag = "Prepare" Then
        'to list all except assets and liablities
        strCodeCond = "Select Acc_Code, Acc_Desc, Open_Bal, Bal_Type, " _
                    & " Qty_ToRec, YTop_Qty from AccCode" _
                    & " Where Acc_Code Not Like '[A,L]*' Order By Acc_Desc"
    End If
    'openning the table
    Set mrstAccCode = mdbsAccounts.OpenRecordset(strCodeCond, dbOpenSnapshot)
    If Not mrstAccCode.EOF Then
        'if record exist
        mrstAccCode.MoveLast
        mrstAccCode.MoveFirst
        'invoking function to display the heading
        Call DisplayHeading
        'performing operation on all accounts
        While Not mrstAccCode.EOF
            'invoking function to calculate the total for the current account
            Call CalCodeBal
            mrstAccCode.MoveNext
        Wend
        'displaying the trial balance total
        Call DisplayTotal
    Else
        'if record does not exist
        'disabling the print menu option
        mnuPrint.Enabled = False
    End If
End Sub
Private Sub CalCodeBal()
    Dim strJourCond As String   'to store the condition for the Journal table
    Dim strCashCond As String   'to store the condition for the Cash table
    Dim rstJournal As Recordset 'to open the Journal table
    Dim rstCash As Recordset    'to open the Cash table
    
    Dim curDB As Currency       'to store the Debit amount
    Dim curCR As Currency       'to store the Credit amount
    Dim curBal As Currency      'to store the Balance amount
    Dim sngQty As Single        'to store the Quantity
    Dim intAmtCol As Integer    'to store the Amount Column Number
   
    curBal = 0#
    sngQty = 0#
    'setting the openning balance according to the Balance type
    If Not IsNull(mrstAccCode!Open_Bal) Then
        'if opening balance exist
        If mrstAccCode!Bal_Type = "D" Then
            'setting the balance as debit
            curDB = mrstAccCode!Open_Bal
            curCR = 0#
        ElseIf mrstAccCode!Bal_Type = "C" Then
            'setting the balance as credit
            curDB = 0#
            curCR = mrstAccCode!Open_Bal
        Else
            'setting both amount to 0 since no balance exist
            curDB = 0#
            curCR = 0#
        End If
    Else
        'setting both amount to 0 since no balance exist
        curDB = 0#
        curCR = 0#
    End If
    'setting Quantity to the Year Top Quantity
    If Not IsNull(mrstAccCode!YTop_Qty) Then
        sngQty = mrstAccCode!YTop_Qty
    End If
    
    'Searching for the Account's transactions in the Journal table
    strJourCond = "Select Quantity, [Debit/Credit], Amount from " _
        & "Journal Where Acc_Code = '" & mrstAccCode!Acc_Code _
        & "' Order by Date"
    Set rstJournal = mdbsAccounts.OpenRecordset(strJourCond, dbOpenSnapshot)
    With rstJournal
        If Not .EOF Then
            .MoveLast
            .MoveFirst
        End If
        'calculating the total debit and credit amount and the quantity for that account
        While Not .EOF
            If ![Debit/Credit] = "D" Then
                curDB = Format(curDB, "#####.00") + !Amount
            ElseIf ![Debit/Credit] = "C" Then
                curCR = Format(curCR, "#####.00") + !Amount
            End If
            If Not IsNull(!Quantity) Then
                sngQty = Format(sngQty, "#####.00") + !Quantity
            End If
            .MoveNext
        Wend
    End With
    'Searching for the Account's transactions in the Cash table
    strCashCond = "Select Quantity, [Debit/Credit], Amount from " _
        & " Cash Where Acc_Code = '" & mrstAccCode!Acc_Code _
        & "' Order By Date"
    Set rstCash = mdbsAccounts.OpenRecordset(strCashCond, dbOpenSnapshot)
    With rstCash
        If Not .EOF Then
            .MoveLast
            .MoveFirst
        End If
        'calculating the total debit and credit amount and the quantity for that account
        While Not .EOF
            If ![Debit/Credit] = "D" Then
                curDB = Format(curDB, "#####.00") + !Amount
            ElseIf ![Debit/Credit] = "C" Then
                curCR = Format(curCR, "#####.00") + !Amount
            End If
            If Not IsNull(!Quantity) Then
                sngQty = Format(sngQty, "#####.00") + !Quantity
            End If
            .MoveNext
        Wend
    End With
    'calculating the Balance amount
    If curDB > curCR Then
        curBal = Format(curDB, "#####.00") - Format(curCR, "#####.00")
        intAmtCol = 2
        mcurDBTot = Format(mcurDBTot, "#####.00") + curBal
    ElseIf curCR > curDB Then
        curBal = Format(curCR, "#####.00") - Format(curDB, "#####.00")
        intAmtCol = 3
        mcurCRTot = Format(mcurCRTot, "#####.00") + curBal
    End If
    'displaying the informations about that account
    msflxgrdTrialBalance.TextMatrix(mintCurRow, 0) = mrstAccCode!Acc_Desc
    If sngQty <> 0# Then
        msflxgrdTrialBalance.TextMatrix(mintCurRow, 1) = Format(sngQty, "####0.00")
    Else
        'if quantity is zero
        If mrstAccCode!Qty_ToRec = "Y" Then
            'if quantity to record is true and to display zero for the quantity
            msflxgrdTrialBalance.TextMatrix(mintCurRow, 1) = Format(sngQty, "####0.00")
        End If
    End If
    If curBal <> 0# Then
        msflxgrdTrialBalance.TextMatrix(mintCurRow, intAmtCol) = Format(curBal, "####0.00")
    End If
    mintCurRow = mintCurRow + 1
End Sub
Private Sub DisplayHeading()
    'displaying the Heading
    With msflxgrdTrialBalance
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Rows = mrstAccCode.RecordCount + 2
        .Cols = 4
        .TextMatrix(0, 0) = "Account Name"
        .ColWidth(0) = 4500
        .TextMatrix(0, 1) = "Quantity"
        .ColWidth(1) = 1250
        .TextMatrix(0, 2) = "Amount  Db"
        .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "Amount  Cr"
        .ColWidth(3) = 1500
    End With
End Sub
Private Sub DisplayTotal()
    'displaying Total for the current account
    With msflxgrdTrialBalance
        .TextMatrix(mintCurRow, 0) = "Total"
        If mcurDBTot <> 0# Then
            .TextMatrix(mintCurRow, 2) = Format(mcurDBTot, "#####.00")
        End If
        If mcurCRTot <> 0# Then
            .TextMatrix(mintCurRow, 3) = Format(mcurCRTot, "#####.00")
        End If
    End With
End Sub
Private Sub FillingTrialBalance()
Dim dbsTrialBalance As Database
Dim rstTrialBalance As Recordset
Dim intCurRow As Integer
Dim intTotRow As Integer
    'openning the database
    Set dbsTrialBalance = OpenDatabase(gstrDatabase)
    'deleting existing record in the TrialBalance table
    dbsTrialBalance.Execute "Delete * from TrialBalance"
    'openning the TrailBalance table for filling it with records to be printed
    Set rstTrialBalance = dbsTrialBalance.OpenRecordset("Select * from TrialBalance", dbOpenDynaset)
    'filling the details into the table
    With msflxgrdTrialBalance
        intTotRow = .Rows - 1
        intCurRow = 1
        While intCurRow < intTotRow
            'checking if user has selected without Null Balance Account and if null balance exist
            If Not ((.TextMatrix(intCurRow, 2) = "") And (.TextMatrix(intCurRow, 3) = "") And _
                mstrPrint = "WONB") Then
                'if the condition is false adding account information to the table
                rstTrialBalance.AddNew
                If (.TextMatrix(intCurRow, 0) <> "") Then
                    rstTrialBalance!AccDesc = .TextMatrix(intCurRow, 0)
                End If
                If (.TextMatrix(intCurRow, 1) <> "") Then
                    rstTrialBalance!Quantity = .TextMatrix(intCurRow, 1)
                End If
                If (.TextMatrix(intCurRow, 2) <> "") Then
                    rstTrialBalance!Debit = .TextMatrix(intCurRow, 2)
                End If
                If (.TextMatrix(intCurRow, 3) <> "") Then
                    rstTrialBalance!Credit = .TextMatrix(intCurRow, 3)
                End If
                rstTrialBalance.Update
            End If
            intCurRow = intCurRow + 1
        Wend
    End With
    rstTrialBalance.Close
    dbsTrialBalance.Close
End Sub
Private Sub mnuWNB_Click()
    'setting to print all the accounts including one with Null Balance
    mstrPrint = "WNB"
    Call PrintTB
End Sub
Private Sub mnWONB_Click()
    'setting to print only accounts without Null Balance
    mstrPrint = "WONB"
    Call PrintTB
End Sub
Private Sub PrintTB()
Dim rstDate As Recordset    'to open the Jounal and Cash table
Dim strDate As String
Dim intCurRow As Integer
    
    'invoking function to fill the TrialBalance table with necessary records
    Call FillingTrialBalance
    'openning the Journal table to obtain the last transaction date
    Set rstDate = mdbsAccounts.OpenRecordset("Select distinct Date from Journal order by date", dbOpenSnapshot)
    'setting strdate to the Year beginning
    strDate = "1/4/" & Left(gstrAccYear, 4)
    If (Not rstDate.EOF) Then
        rstDate.MoveLast
        'checking if strdate is less than the last transaction date
        If (CDate(strDate) < Format(rstDate!Date, "dd/mm/yyyy")) Then
            'if yes setting last transaction date to strdate
            strDate = Format(rstDate!Date, "dd/mm/yyyy")
        End If
    End If
    rstDate.Close
    'openning the Cash table to obtain the last transaction date
    Set rstDate = mdbsAccounts.OpenRecordset("Select distinct Date from Cash order by date", dbOpenSnapshot)
    If (Not rstDate.EOF) Then
        rstDate.MoveLast
        'checking if strdate is less than the last transaction date
        If (CDate(strDate) < Format(rstDate!Date, "dd/mm/yyyy")) Then
            'if yes setting last transaction date to strdate
            strDate = Format(rstDate!Date, "dd/mm/yyyy")
        End If
    End If
    rstDate.Close
    'displaying the Header and Footer informations of the Report
    With msflxgrdTrialBalance
        intCurRow = .Rows - 1
        DataEnvironment1.TrialBalance
        If (gstrCompanyCode = "SNTC") Then
            rptTrialBalance.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI NARAYANA TRADING COMPANY"
        ElseIf (gstrCompanyCode = "VPPS") Then
            rptTrialBalance.Sections("PageHeader").Controls.Item("lblCompany").Caption = "V.P.PACHAIYAPPA MUDALIAR & SONS"
        Else
            rptTrialBalance.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI PACHAIYAPPA MARKETING AGENCIES"
        End If
        rptTrialBalance.Sections("PageHeader").Controls.Item("lblAccYear").Caption = "Account Year " & gstrAccYear
        rptTrialBalance.Sections("PageHeader").Controls.Item("lblYear").Caption = "Trial Balance as on " & strDate
        If (Not IsNull(.TextMatrix(intCurRow, 2))) Then
            rptTrialBalance.Sections("ReportFooter").Controls.Item("lblTotalDebit").Caption = .TextMatrix(intCurRow, 2)
        End If
        If (Not IsNull(.TextMatrix(intCurRow, 3))) Then
            rptTrialBalance.Sections("ReportFooter").Controls.Item("lblTotalCredit").Caption = .TextMatrix(intCurRow, 3)
        End If
    End With
    rptTrialBalance.Show vbModal
    DataEnvironment1.rsTrialBalance.Close
End Sub

