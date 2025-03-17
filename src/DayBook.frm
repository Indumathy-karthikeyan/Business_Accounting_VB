VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDayBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DayBook"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   11580
   Begin VB.TextBox txtDate 
      Height          =   300
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   1470
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   2130
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   435
      Width           =   1470
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   6255
      TabIndex        =   1
      Tag             =   "Find"
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Height          =   375
      Left            =   7335
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid msflexgrdDayBook 
      Height          =   5145
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   11220
      _ExtentX        =   19791
      _ExtentY        =   9075
      _Version        =   393216
      Cols            =   6
      AllowBigSelection=   0   'False
   End
   Begin VB.Label lblDayBook 
      Caption         =   "Day Book for "
      Height          =   210
      Left            =   975
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Menu mnuForm 
      Caption         =   "&Form"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu mnuNavigation 
      Caption         =   "Na&vigation"
      Begin VB.Menu mnuFirst 
         Caption         =   "&First"
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "&Previous"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last"
      End
   End
End
Attribute VB_Name = "frmDayBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
    Dim mdbsAccounts As Database    'To open the database
    Dim mrstJournal As Recordset    'To open the Journal table
    Dim mrstCash As Recordset       'To open the Cash table
    Dim mrstAccCode As Recordset
    
    Dim mblnAccCode As Boolean        'Flag to close the form
    Dim mblnTrans As Boolean
    Dim mstrDate As String          'To store the specified date
    
    Dim mcurJrnlDb As Currency
    Dim mcurJrnlCr As Currency
    Dim mcurCashDb As Currency
    Dim mcurCashCr As Currency
    Dim mstrType As String
Private Sub cmdCancel_Click()
    'Cancelation of the Find operation
    'To check if current date is startdate or enddate or middle date
    'and to set the Navigation buttons according to it
    If (DateDiff("d", mstrDate, gstrStDate) = 0) Then
        Call RecordPosition("F")
    ElseIf (DateDiff("d", mstrDate, gstrStDate) = 0) Then
        Call RecordPosition("L")
    Else
        Call RecordPosition("M")
    End If
    'Resetting to Display mode
    cmdFind.Caption = "F&ind"
    cmdFind.Tag = "Find"
    'disabling the cancel button
    cmdCancel.Enabled = False
    'hidding the date combo box
    cboDate.Visible = False
    'showing the date text box
    txtDate.Visible = True
    'enabling the print and close menu options
    mnuPrint.Enabled = True
    mnuClose.Enabled = True
    'enabling the grid control
    msflexgrdDayBook.Enabled = True
End Sub
Private Sub cmdFind_Click()
Dim strDate As String           'To store the intermediate date
    If (cmdFind.Tag = "Find") Then
        'in the Find mode
        'preparing the Display button
        cmdFind.Caption = "&Display"
        cmdFind.Tag = "Display"
        'enabling the Cancel button
        cmdCancel.Enabled = True
        'disabling the grid control
        msflexgrdDayBook.Enabled = False
        'hidding the Date textbox
        txtDate.Visible = False
        'showing the Date combobox
        cboDate.Visible = True
        'disabling the Navigation buttons
        Call RecordPosition("")
        'disabling Print and Close menu options
        mnuClose.Enabled = False
        mnuPrint.Enabled = False
        'adding all the dates between start date and end date
        cboDate.Clear
        strDate = gstrStDate
        While (DateDiff("d", strDate, gstrEndDate) <> 0)
            cboDate.AddItem strDate
            strDate = DateAdd("d", 1, strDate)
        Wend
        cboDate.AddItem strDate
        cboDate.SetFocus
    Else
        'checking if the user has selected one
        If (cboDate.ListIndex = -1) Then
            MsgBox "Specify the Date", vbOKOnly
            cboDate.SetFocus
        Else
            'if yes
            mstrDate = cboDate.Text
            'invoking function to display the transaction of the selected date
            Call PreparingDetails
            'positioning the Navigation menu option according to the current date
            If (DateDiff("d", cboDate.Text, gstrStDate) = 0) Then
                Call RecordPosition("F")
            ElseIf (DateDiff("d", cboDate.Text, gstrEndDate) = 0) Then
                Call RecordPosition("L")
            Else
                Call RecordPosition("M")
            End If
            'resetting the Find mode
            cmdFind.Caption = "F&ind"
            cmdFind.Tag = "Find"
            'disabling the Cancel button
            cmdCancel.Enabled = False
            'enabling the Print and Close menu options
            mnuPrint.Enabled = True
            mnuClose.Enabled = True
            'hidding the Date combo box
            cboDate.Visible = False
            'showing the Date text box
            txtDate.Visible = True
            'enabling the grid control
            msflexgrdDayBook.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnAccCode Then
        MsgBox "Account details does not exists", vbOKOnly
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim strAccCode As String
    Dim strJourCond As String
    Dim strCashCond As String
    
    'hidding the Date combo box
    cboDate.Visible = False
    'opening the database
    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    
    'checking whether the start and enddate are same (or) different
    If gstrStDate = gstrEndDate Then
        'if both dates are same hiding the Navigation menu options
        mnuNavigation.Visible = False
        'disabling the Find button
        cmdFind.Enabled = False
    Else
        'if both dates are same
        'showing the Navigation menu options
        mnuNavigation.Visible = True
    End If
    'disabling the Cancel button
    cmdCancel.Enabled = False
    
    'opening recordset for Accounts to get the description
    strAccCode = "Select Acc_Code, Acc_Desc from AccCode"
    Set mrstAccCode = mdbsAccounts.OpenRecordset(strAccCode, dbOpenSnapshot)
    If (mrstAccCode.BOF And mrstAccCode.EOF) Then
        mblnAccCode = True
    Else
        mstrDate = gstrStDate
        Call PreparingDetails
        Call RecordPosition("F")
    End If
End Sub
Private Sub Heading()

    'displaying the column heading
    With msflexgrdDayBook
        .Clear
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 0
        'setting total no. of columns
        .Cols = 7
        .TextMatrix(0, 0) = "Account Description"
        .ColWidth(0) = 3000
        .TextMatrix(0, 1) = "Narration"
        .ColWidth(1) = 2500
        .TextMatrix(0, 2) = "Quantity"
        .ColWidth(2) = 700
        .TextMatrix(0, 3) = "Debit"
        .ColWidth(3) = 1125
        .TextMatrix(0, 4) = "Credit"
        .ColWidth(4) = 1125
        .TextMatrix(0, 5) = "Cash Payment"
        .ColWidth(5) = 1125
        .TextMatrix(0, 6) = "Cash Receipt"
        .ColWidth(6) = 1125
    End With
End Sub
Private Sub DisplayingTotals()
    Dim intRow As Integer
    Dim intCol As Integer           'to store the current column
    Dim curJrnlDiff As Currency     'to store the Journal amount difference
    Dim curCashDiff As Currency     'to store the cash amount difference
    
    curJrnlDiff = 0#
    curCashDiff = 0#
    With msflexgrdDayBook
        'to display the total in the grid
        .Rows = .Rows + 1
        intRow = .Rows - 1
        .TextMatrix(intRow, 1) = "Total"
        If mcurJrnlDb <> 0# Then
            .TextMatrix(intRow, 3) = Format(mcurJrnlDb, "######.00")
        End If
        If mcurJrnlCr <> 0# Then
            .TextMatrix(intRow, 4) = Format(mcurJrnlCr, "######.00")
        End If
        If mcurCashDb <> 0# Then
            .TextMatrix(intRow, 5) = Format(mcurCashDb, "######.00")
        End If
        If mcurCashCr <> 0# Then
            .TextMatrix(intRow, 6) = Format(mcurCashCr, "######.00")
        End If
        intRow = intRow + 1
        .Rows = .Rows + 1
        'to display the Balance in the grid
        .TextMatrix(intRow, 1) = "Balance for the day"
        If mcurJrnlDb > mcurJrnlCr Then
            curJrnlDiff = mcurJrnlDb - mcurJrnlCr
            intCol = 3
        ElseIf mcurJrnlCr > mcurJrnlDb Then
            curJrnlDiff = mcurJrnlCr - mcurJrnlDb
            intCol = 4
        ElseIf mcurJrnlDb = mcurJrnlCr Then
            curJrnlDiff = 0#
        End If
        If curJrnlDiff <> 0# Then
            .TextMatrix(intRow, intCol) = Format(curJrnlDiff, "#####.00")
        End If
        
        If mcurCashDb > mcurCashCr Then
            curCashDiff = mcurCashDb - mcurCashCr
            intCol = 5
        ElseIf mcurCashCr > mcurCashDb Then
            curCashDiff = mcurCashCr - mcurCashDb
            intCol = 6
        ElseIf mcurCashDb = mcurCashCr Then
            curCashDiff = 0#
        End If
        If curCashDiff <> 0# Then
            .TextMatrix(intRow, intCol) = Format(curCashDiff, "#####.00")
        End If
    End With
End Sub
Private Sub CalPrevBal()
'to calculate the Balance before the current date
    Dim strCondition As String      'to store the condition for data retrieval
    Dim rstCashTemp As Recordset    'to open the cash table
    Dim curDB As Currency           'to store the debit amount
    Dim curCR As Currency           'to store the credit amount
    Dim curBal As Currency
    Dim strBalType As String
    
    'openning the Cash table with transaction upto the current date
    strCondition = "Select Date, [Debit/Credit], Amount from Cash " & _
                "Where date < cdate('" & mstrDate & "') Order By Date "
    Set rstCashTemp = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    With rstCashTemp
        If Not .EOF Then
            curDB = 0#
            curCR = 0#
            'calculating the total credit and debit amount
            While Not .EOF
                If ![Debit/Credit] = "D" Then
                    curDB = Format(curDB, "#####.00") + !Amount
                ElseIf ![Debit/Credit] = "C" Then
                    curCR = Format(curCR, "#####.00") + !Amount
                End If
                .MoveNext
            Wend
            curBal = 0
            strBalType = ""
            'calculating the amount balance
            If curDB > curCR Then
                curBal = Format(curDB, "#####.00") - Format(curCR, "#####.00")
                strBalType = "D"
            ElseIf curCR > curDB Then
                curBal = Format(curCR, "#####.000") - Format(curDB, "#####.00")
                strBalType = "C"
            End If
        End If
    End With
    
    With msflexgrdDayBook
        'displaying the openning balance in the grid
        If strBalType = "" Then
            .TextMatrix(1, 1) = "Opening Balance  (Nill)"
        ElseIf strBalType = "D" Then
            .TextMatrix(1, 1) = "Opening Balance for the day"
            .TextMatrix(1, 5) = Format(curBal, "#####.00")
        ElseIf strBalType = "C" Then
            .TextMatrix(1, 1) = "Opening Balance for the day"
            .TextMatrix(1, 6) = Format(curBal, "#####.00")
        End If
    End With
    
    mcurJrnlDb = 0#
    mcurJrnlCr = 0#
    'setting the Cash debit or credit according to the balance type
    If strBalType = "D" Then
        mcurCashDb = curBal
        mcurCashCr = 0#
    ElseIf strBalType = "C" Then
        mcurCashCr = curBal
        mcurCashDb = 0#
    ElseIf strBalType = "" Then
        mcurCashDb = 0#
        mcurCashCr = 0#
    End If
End Sub
Private Sub PreparingDetails()
Dim strCondition  As String
Dim rstCash As Recordset
Dim rstJournal As Recordset
 
    Call Heading
    'calculating and displaying opening balance for the day
    Call CalPrevBal

    txtDate.Text = mstrDate
    mblnTrans = False
    'getting transaction from Journal on the specified date
    'transactions other than product and expenses
    'to display them in the order of their entry
    strCondition = "Select Date, Acc_Code, Narration, Quantity, " & _
                "Amount, [Debit/Credit] from Journal where " & _
                "Date = cdate('" & mstrDate & "') and " & _
                "(Acc_Code not like 'I*') Order By EntryNo"
    Set mrstJournal = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not mrstJournal.BOF And Not mrstJournal.EOF) Then
        mblnTrans = True
        Call DisplayingJournal("Others")
    End If
    mrstJournal.Close
    
    'getting transaction from Journal on the specified date
    'product and expense transaction
    'to display the consolidated entry of the product or expense
    strCondition = "Select Date, Acc_Code, Narration, Quantity, " & _
                "Amount, [Debit/Credit] from Journal where " & _
                "Date = cdate('" & mstrDate & "') and " & _
                "(Acc_Code like 'I*') Order By Acc_Code, EntryNo"
    Set mrstJournal = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not mrstJournal.BOF And Not mrstJournal.EOF) Then
        mblnTrans = True
        Call DisplayingJournal("Sales")
    End If
    mrstJournal.Close
    
    'getting transaction from Cash on the specified date
    'transactions other than product Sales and expenses
    'to display them in the order of their entry
    strCondition = "Select Date, Acc_Code, Narration, Quantity, " & _
                "Amount, [Debit/Credit] from Cash where " & _
                "Date = cdate('" & mstrDate & "') and " & _
                "(Acc_Code not like 'I*') Order By EntryNo"
    Set mrstCash = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not mrstCash.BOF And Not mrstCash.EOF) Then
        mblnTrans = True
        Call DisplayingCash("Others")
    End If
    mrstCash.Close
    
    'getting transaction from Cash on the specified date
    'product and expense transaction
    'to display the consolidated entry of the product or expense
    strCondition = "Select Date, Acc_Code, Narration, Quantity, " & _
                "Amount, [Debit/Credit] from Cash where " & _
                "Date = cdate('" & mstrDate & "') and " & _
                "(Acc_Code like 'I*') Order By Acc_Code,EntryNo"
    Set mrstCash = mdbsAccounts.OpenRecordset(strCondition, dbOpenSnapshot)
    If (Not mrstCash.BOF And Not mrstCash.EOF) Then
        mblnTrans = True
        Call DisplayingCash("Sales")
    End If
    mrstCash.Close
    If mblnTrans Then
        Call DisplayingTotals
    Else
        'if transaction does not exist
        'clearing the grid
        msflexgrdDayBook.Clear
        msflexgrdDayBook.Rows = 1
        Call Heading
        'displaying message
        MsgBox "Transaction for " & mstrDate & "date does not exist", vbOKOnly
        'disabling the Print menu option
        mnuPrint.Enabled = False
    End If
End Sub
Private Sub DisplayingJournal(strType As String)
'storing product Sales or expense code to calculate consolidated information
Dim strAccCode As String
Dim curQuantity As Currency
Dim curAmount As Currency
Dim strCondition As String
Dim intCurRow As Integer
Dim blnCreditFlag As Boolean

    mrstJournal.MoveLast
    mrstJournal.MoveFirst
    Do While (Not mrstJournal.EOF)
    
        msflexgrdDayBook.Rows = msflexgrdDayBook.Rows + 1
        
        intCurRow = msflexgrdDayBook.Rows - 1
        msflexgrdDayBook.Row = intCurRow
        blnCreditFlag = False
        If (strType = "Sales") Then
            If mrstJournal!Acc_Code <> "IPAL" Then
                strAccCode = mrstJournal!Acc_Code
                curQuantity = 0
                curAmount = 0
                Do While ((mrstJournal!Acc_Code = strAccCode) And (mrstJournal![Debit/Credit] = "C"))
                    blnCreditFlag = True
                    If Not IsNull(mrstJournal!Quantity) Then
                        curQuantity = curQuantity + mrstJournal!Quantity
                    End If
                    curAmount = curAmount + mrstJournal!Amount
                    mrstJournal.MoveNext
                    If (mrstJournal.EOF) Then
                        Exit Do
                    End If
                Loop
                If blnCreditFlag Then
                    mrstJournal.MovePrevious
                End If
            End If
        End If
        With msflexgrdDayBook
            strCondition = "Acc_Code = '" & mrstJournal!Acc_Code & "'"
            mrstAccCode.FindFirst strCondition
            If (mrstAccCode.NoMatch) Then
                MsgBox "Details for the Account Code " & mrstJournal!Acc_Code & "is missing", vbOKOnly
            Else
                .TextMatrix(intCurRow, 0) = mrstAccCode!Acc_Desc
            End If
            If (Not blnCreditFlag) And (Not IsNull(mrstJournal!Narration)) Then
                .TextMatrix(intCurRow, 1) = mrstJournal!Narration
            End If
            If (blnCreditFlag) Then
                .TextMatrix(intCurRow, 2) = Format(curQuantity, "####.00")
            Else
                If (Not IsNull(mrstJournal!Quantity)) Then
                    .TextMatrix(intCurRow, 2) = Format(mrstJournal!Quantity, "####.00")
                End If
            End If
            If Not blnCreditFlag Then
                If (mrstJournal![Debit/Credit] = "D") Then
                    .TextMatrix(intCurRow, 3) = Format(mrstJournal!Amount, "######.00")
                    mcurJrnlDb = mcurJrnlDb + mrstJournal!Amount
                ElseIf (mrstJournal![Debit/Credit] = "C") Then
                    .TextMatrix(intCurRow, 4) = Format(mrstJournal!Amount, "######.00")
                    mcurJrnlCr = mcurJrnlCr + mrstJournal!Amount
                End If
            Else
                .TextMatrix(intCurRow, 4) = Format(curAmount, "######.00")
                mcurJrnlCr = mcurJrnlCr + curAmount
            End If
        End With
        mrstJournal.MoveNext
        If mrstJournal.EOF Then
            Exit Do
        End If
    Loop
End Sub
Private Sub DisplayingCash(strType As String)
'storing product Sales or expense code to calculate consolidated information
Dim strAccCode As String
'to consolidate quantity in case of product Sales transaction
Dim curQuantity As Currency
'to cosolidate amount in case of product Sales transaction
Dim curAmount As Currency
'to condition to find the Description for the account
Dim strCondition As String
'to display the data in the particular row in the grid
Dim intCurRow As Integer
'indicates the credit transaction of product sales
Dim blnCreditFlag As Boolean

    'to enable the recordset for travelling
    mrstCash.MoveLast
    mrstCash.MoveFirst
    Do While (Not mrstCash.EOF)
        'increasing the rows in the grid to display another transaction
        msflexgrdDayBook.Rows = msflexgrdDayBook.Rows + 1
        'getting last row no in the grid
        intCurRow = msflexgrdDayBook.Rows - 1
        'setting that row as the current row
        msflexgrdDayBook.Row = intCurRow
        'resetting the credit product sales
        blnCreditFlag = False
        If (strType = "Sales") Then
            'in case of displaying product Sales transaction
            'to get the consolidated information
            If mrstCash!Acc_Code <> "IPAL" Then
            'in case the transaction is not IPAL
                strAccCode = mrstCash!Acc_Code
                curQuantity = 0
                curAmount = 0
                Do While (mrstCash!Acc_Code = strAccCode) And (mrstCash![Debit/Credit] = "C")
                'looping for the same acccode and credit transaction
                'getting the consolidated quantity and amount
                    blnCreditFlag = True
                    If Not IsNull(mrstCash!Quantity) Then
                        curQuantity = curQuantity + mrstCash!Quantity
                    End If
                    curAmount = curAmount + mrstCash!Amount
                    mrstCash.MoveNext
                    If (mrstCash.EOF) Then
                        'in case of end of recordset exit do loop
                        Exit Do
                    End If
                Loop
                If blnCreditFlag Then
                    'moving backward since it has moved to next trans or
                    'it might be eof in case of calculating consolidated quantity and amount
                    mrstCash.MovePrevious
                End If
            End If
        End If
        
        'displaying the details in the grid
        With msflexgrdDayBook
            'displaying the description for the account
            'by finding in the account table
            strCondition = "Acc_Code = '" & mrstCash!Acc_Code & "'"
            mrstAccCode.FindFirst strCondition
            If (mrstAccCode.NoMatch) Then
                MsgBox "Details for the Account Code " & mrstCash!Acc_Code & "is missing", vbOKOnly
            Else
                .TextMatrix(intCurRow, 0) = mrstAccCode!Acc_Desc
            End If
            'displaying narration
            If (Not blnCreditFlag) And (Not IsNull(mrstCash!Narration)) Then
                'displaying narration for other trans if present
                'since product Credit Sales trans has no narration
                .TextMatrix(intCurRow, 1) = mrstCash!Narration
            End If
            If (blnCreditFlag) Then
                'in case of product credit Sales trans displaying consolidated quantity
                .TextMatrix(intCurRow, 2) = Format(curQuantity, "####.00")
            Else
                'in case of other trans
                'displaying quantity if present
                If (Not IsNull(mrstCash!Quantity)) Then
                    .TextMatrix(intCurRow, 2) = Format(mrstCash!Quantity, "####.00")
                End If
            End If
            If (Not blnCreditFlag) Then
                If (mrstCash![Debit/Credit] = "D") Then
                    'in case of other trans diaplying trans amount
                    .TextMatrix(intCurRow, 5) = Format(mrstCash!Amount, "######.00")
                    'adding trans amount to total cash debit amount
                    mcurCashDb = mcurCashDb + mrstCash!Amount
                ElseIf (mrstCash![Debit/Credit] = "C") Then
                    'in case of other trans displaying trans amount
                    .TextMatrix(intCurRow, 6) = Format(mrstCash!Amount, "######.00")
                    'adding trans amount to total cash credit amount
                    mcurCashCr = mcurCashCr + mrstCash!Amount
                End If
            Else
                'in case of product credit sales trans displaying consolidated amount
                    .TextMatrix(intCurRow, 6) = Format(curAmount, "######.00")
                    'adding consolidated amount to total Cash Credit amount
                    mcurCashCr = mcurCashCr + curAmount
            End If
        End With
        mrstCash.MoveNext
        If mrstCash.EOF Then
            'in case of end of recordset exit loop
            Exit Do
        End If
    Loop

End Sub
Private Sub RecordPosition(strPos As String)
    
    'setting the navigation buttons according
    'to the position
    If strPos = "F" Then
        mnuFirst.Enabled = False
        mnuPrevious.Enabled = False
        mnuNext.Enabled = True
        mnuLast.Enabled = True
    ElseIf strPos = "L" Then
        mnuFirst.Enabled = True
        mnuPrevious.Enabled = True
        mnuNext.Enabled = False
        mnuLast.Enabled = False
    ElseIf strPos = "M" Then
        mnuFirst.Enabled = True
        mnuPrevious.Enabled = True
        mnuNext.Enabled = True
        mnuLast.Enabled = True
    ElseIf strPos = "" Then
        mnuFirst.Enabled = False
        mnuPrevious.Enabled = False
        mnuNext.Enabled = False
        mnuLast.Enabled = False
    End If
End Sub
Private Sub mnuClose_Click()
    Unload Me
End Sub
Private Sub mnuFirst_Click()
    
    'setting the start date as the current date
    mstrDate = gstrStDate
    'invoking the function to set the menu options according to the current date
    Call RecordPosition("F")
    'invoking function to display the transactions of the current date
    Call PreparingDetails
End Sub
Private Sub mnuLast_Click()
    'setting the end date as the current date
    mstrDate = gstrEndDate
    'invoking the function to set the menu options according to the current date
    Call RecordPosition("L")
    'invoking function to display the transactions of the current date
    Call PreparingDetails
End Sub
Private Sub mnuNext_Click()
    'setting the next date to the current date as the current date
    mstrDate = Format(DateAdd("d", 1, mstrDate), "mm/dd/yyyy")
    'invoking the function to set the menu options according to the current date
    If mstrDate = gstrEndDate Then
        Call RecordPosition("L")
    Else
        Call RecordPosition("M")
    End If
    'invoking function to display the transactions of the current date
    Call PreparingDetails
End Sub
Private Sub mnuPrevious_Click()
    'setting the previous date to the current date as the current date
    mstrDate = Format(DateAdd("d", -1, mstrDate), "mm/dd/yyyy")
    'invoking the function to set the menu options according to the current date
    If mstrDate = gstrStDate Then
        Call RecordPosition("F")
    Else
        Call RecordPosition("M")
    End If
    'invoking function to display the transactions of the current date
    Call PreparingDetails
End Sub

Private Sub mnuPrint_Click()
Dim intCurRow As Integer
    Call FillingDayBook
    With msflexgrdDayBook
        intCurRow = .Rows - 2
        DataEnvironment1.DayBook
            If (gstrCompanyCode = "SNTC") Then
                rptDayBook.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI NARAYANA TRADING COMPANY"
            ElseIf (gstrCompanyCode = "VPPS") Then
                rptDayBook.Sections("PageHeader").Controls.Item("lblCompany").Caption = "V.P.PACHAIYAPPA MUDALIAR & SONS"
            Else
                rptDayBook.Sections("PageHeader").Controls.Item("lblCompany").Caption = "SRI PACHAIYAPPA MARKETING AGENCIES"
            End If
            rptDayBook.Sections("PageHeader").Controls.Item("lblAccYear").Caption = "Account Year " & gstrAccYear
            rptDayBook.Sections("PageHeader").Controls.Item("lblAccDate").Caption = "Day Book for: " & Format(mstrDate, "dd/mm/yyyy")
            If (Not IsNull(.TextMatrix(intCurRow, 2))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblTotalQuantity").Caption = .TextMatrix(intCurRow, 2)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 3))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblTotalDebit").Caption = .TextMatrix(intCurRow, 3)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 4))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblTotalCredit").Caption = .TextMatrix(intCurRow, 4)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 5))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblTotalCashPayment").Caption = .TextMatrix(intCurRow, 5)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 6))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblTotalCashReceipt").Caption = .TextMatrix(intCurRow, 6)
            End If
            intCurRow = intCurRow + 1
            If (Not IsNull(.TextMatrix(intCurRow, 2))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblBalanceQuantity").Caption = .TextMatrix(intCurRow, 2)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 3))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblBalanceDebit").Caption = .TextMatrix(intCurRow, 3)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 4))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblBalanceCredit").Caption = .TextMatrix(intCurRow, 4)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 5))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblBalanceCashPayment").Caption = .TextMatrix(intCurRow, 5)
            End If
            If (Not IsNull(.TextMatrix(intCurRow, 6))) Then
                rptDayBook.Sections("ReportFooter").Controls.Item("lblBalanceCashReceipt").Caption = .TextMatrix(intCurRow, 6)
            End If
        rptDayBook.Show vbModal
        DataEnvironment1.rsDayBook.Close
    End With
End Sub
Private Sub FillingDayBook()
Dim dbsDayBook As Database
Dim rstDayBook As Recordset
Dim intCurRow As Integer
Dim intTotRow As Integer

    Set dbsDayBook = OpenDatabase(gstrDatabase)
    dbsDayBook.Execute "Delete * from DayBook"
    Set rstDayBook = dbsDayBook.OpenRecordset("Select * from DayBook", dbOpenDynaset)
        
    With msflexgrdDayBook
        intTotRow = .Rows - 2
        intCurRow = 1
        While intCurRow < intTotRow
            rstDayBook.AddNew
            If (.TextMatrix(intCurRow, 0) <> "") Then
                rstDayBook("AccDesc") = .TextMatrix(intCurRow, 0)
            End If
            If (.TextMatrix(intCurRow, 1) <> "") Then
                rstDayBook("Narration") = .TextMatrix(intCurRow, 1)
            End If
            If (.TextMatrix(intCurRow, 2) <> "") Then
                rstDayBook!Quantity = CSng(.TextMatrix(intCurRow, 2))
            End If
            If (.TextMatrix(intCurRow, 3) <> "") Then
                rstDayBook("Debit") = .TextMatrix(intCurRow, 3)
            End If
            If (.TextMatrix(intCurRow, 4) <> "") Then
                rstDayBook("Credit") = .TextMatrix(intCurRow, 4)
            End If
            If (.TextMatrix(intCurRow, 5) <> "") Then
                rstDayBook("CashPayment") = .TextMatrix(intCurRow, 5)
            End If
            If (.TextMatrix(intCurRow, 6) <> "") Then
                rstDayBook("CashReceipt") = .TextMatrix(intCurRow, 6)
            End If
            rstDayBook.Update
            intCurRow = intCurRow + 1
        Wend
    End With
    rstDayBook.Close
    dbsDayBook.Close
End Sub
