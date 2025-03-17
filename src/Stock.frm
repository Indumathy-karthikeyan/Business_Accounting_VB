VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9975
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   9975
   ScaleWidth      =   12705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRegister 
      Caption         =   "Stock Register"
      Height          =   3735
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   6120
      Width           =   8895
      Begin MSFlexGridLib.MSFlexGrid msflxgrdStock 
         Height          =   3375
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5953
         _Version        =   393216
         Rows            =   13
      End
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   3480
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   10680
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txtDate 
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Frame fraRegister 
      Caption         =   "Sales Register"
      Height          =   5295
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   12255
      Begin TabDlg.SSTab sstabRegister 
         Height          =   4575
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   8070
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Product Details"
         TabPicture(0)   =   "Stock.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "msflxgrdProduct"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Personal Details"
         TabPicture(1)   =   "Stock.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "msflxgrdPersonal"
         Tab(1).ControlCount=   1
         Begin MSFlexGridLib.MSFlexGrid msflxgrdPersonal 
            Height          =   3615
            Left            =   -74040
            TabIndex        =   10
            Top             =   600
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   6376
            _Version        =   393216
         End
         Begin MSFlexGridLib.MSFlexGrid msflxgrdProduct 
            Height          =   3855
            Left            =   240
            TabIndex        =   2
            Top             =   480
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   6800
            _Version        =   393216
            Rows            =   15
         End
      End
   End
   Begin VB.Label lblDate 
      Caption         =   "Date"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
         Begin VB.Menu mnuSalesReg 
            Caption         =   "Sales Register"
         End
         Begin VB.Menu mnuStock 
            Caption         =   "Stock"
         End
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuNavigation 
      Caption         =   "Navigation"
      Begin VB.Menu mnuFirst 
         Caption         =   "First"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "Prev"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mdbsAccounts As Database
Dim mstrCurDate As String
Dim mblnTrans As Boolean
Dim mrstAccCode As Recordset
Dim mblnClose As Boolean

Private Sub cmdCancel_Click()
    'Cancelation of the Find operation
    'To check if current date is startdate or enddate or middle date
    'and to set the Navigation buttons according to it
    If (DateDiff("d", mstrCurDate, gstrStDate) = 0) Then
        Call RecordPosition("F")
    ElseIf (DateDiff("d", mstrCurDate, gstrStDate) = 0) Then
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
    sstabRegister.Enabled = True
    msflxgrdStock.Enabled = True

End Sub
Private Sub RecordPosition(strPos As String)
    
    'setting the navigation buttons according
    'to the position
    If strPos = "F" Then
        mnuFirst.Enabled = False
        mnuPrev.Enabled = False
        mnuNext.Enabled = True
        mnuLast.Enabled = True
    ElseIf strPos = "L" Then
        mnuFirst.Enabled = True
        mnuPrev.Enabled = True
        mnuNext.Enabled = False
        mnuLast.Enabled = False
    ElseIf strPos = "M" Then
        mnuFirst.Enabled = True
        mnuPrev.Enabled = True
        mnuNext.Enabled = True
        mnuLast.Enabled = True
    ElseIf strPos = "" Then
        mnuFirst.Enabled = False
        mnuPrev.Enabled = False
        mnuNext.Enabled = False
        mnuLast.Enabled = False
    End If
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
        sstabRegister.Enabled = False
        msflxgrdStock.Enabled = False
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
            mstrCurDate = cboDate.Text
            'invoking function to display the transaction of the selected date
            Call DisplayingAll
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
            sstabRegister.Enabled = True
            msflxgrdStock.Enabled = True
        End If
    End If
End Sub

Private Sub Form_Activate()
    If mblnClose Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim strCond As String

    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    strCond = "Select Acc_Code,Acc_Desc,YTop_Qty from AccCode " & _
            " Where ((Acc_Code like 'I*') or (Acc_Code like 'P*')" & _
            " or (Acc_Code like 'A*')) Order by Acc_Code"
    Set mrstAccCode = mdbsAccounts.OpenRecordset(strCond, dbOpenDynaset)
    If (Not mrstAccCode.BOF And Not mrstAccCode.EOF) Then
        If (gstrStDate = gstrEndDate) Then
            mnuNavigation.Visible = False
            cmdFind.Enabled = False
            cmdCancel.Enabled = False
        Else
            mnuNavigation.Visible = True
            Call RecordPosition("F")
        End If
        mstrCurDate = gstrStDate
        mblnTrans = False
        If (gstrRegister = "Cash") Then
            sstabRegister.TabEnabled(1) = False
        End If
        Call DisplayingAll
        mblnClose = False
        cboDate.Visible = False
        cmdFind.Tag = "Find"
    Else
        mblnClose = True
        MsgBox "No Accounts Details exists", vbOKOnly
    End If
End Sub
Private Sub DisplayingAll()
            
    Call DisplayingProduct
    If (gstrRegister = "Credit") Then
        Call DisplayingPersonal
    End If
    Call DisplayingStock
End Sub
Private Sub DisplayingProduct()
Dim strCond As String
    
    Call ProductHeading
    txtDate.Text = Format(mstrCurDate, "mm/dd/yyyy")
    If (gstrRegister = "Cash") Then
        strCond = "Select Acc_Code, BillNo, Quantity, Amount from Cash " & _
                "Where ((Date = CDate('" & mstrCurDate & "')) and (Acc_Code like 'I*'))" & _
                " Order by Acc_Code,BillNo"
    ElseIf (gstrRegister = "Credit") Then
        strCond = "Select Acc_Code, BillNo, Quantity, Amount from Journal " & _
                "Where ((Date = CDate('" & mstrCurDate & "')) and (Acc_Code like 'I*'))" & _
                " Order by Acc_Code,BillNo"
    End If
    Call DisplayingDetails(strCond)
    If (mblnTrans) Then
    Else
        MsgBox "Transaction for Specified date does not exist", vbOKOnly
    End If
End Sub
Private Sub DisplayingDetails(strCond As String)
Dim rstTrans As Recordset
Dim strCurAccCode As String
Dim intRow As Integer
Dim intConRow As Integer
Dim blnMore As Boolean
Dim intBillNo As Integer
Dim curQuantity As Currency
Dim curAmount As Currency
Dim strFindCond As String
Dim blnEof As Boolean
Dim curTotal As Currency
Dim strAccDesc As String
Dim intPos As Integer

    Set rstTrans = mdbsAccounts.OpenRecordset(strCond, dbOpenSnapshot)
    If (Not rstTrans.BOF And Not rstTrans.EOF) Then
        mblnTrans = True
        rstTrans.MoveLast
        rstTrans.MoveFirst
        intRow = 1
        intConRow = 1
        blnEof = False
        curTotal = 0
        Do While (Not rstTrans.EOF)
            strCurAccCode = rstTrans!Acc_Code
            strFindCond = "Acc_Code = '" & strCurAccCode & "'"
            mrstAccCode.FindFirst strFindCond
            If Not mrstAccCode.NoMatch Then
                intPos = InStr(1, mrstAccCode!Acc_Desc, "(")
                strAccDesc = Mid(mrstAccCode!Acc_Desc, intPos + 1)
                intPos = InStr(1, strAccDesc, ")")
                strAccDesc = Mid(strAccDesc, 1, intPos - 1)
            Else
                strAccDesc = ""
            End If
            
            intBillNo = 0
            curQuantity = 0
            curAmount = 0
            blnMore = False
            Do While (strCurAccCode = rstTrans!Acc_Code)
                If (Not blnMore) Then
                    rstTrans.MoveNext
                    If (Not rstTrans.EOF) Then
                        If (rstTrans!Acc_Code = strCurAccCode) Then
                            blnMore = True
                        End If
                        rstTrans.MovePrevious
                    Else
                        rstTrans.MovePrevious
                    End If
                End If
                If (blnMore) Then
                    curQuantity = curQuantity + rstTrans!Quantity
                    intBillNo = 0
                    curAmount = curAmount + rstTrans!Amount
                    With msflxgrdProduct
                        If (intRow >= intConRow) Then
                            .Rows = .Rows + 1
                        End If
                        .TextMatrix(intRow, 5) = strAccDesc
                        .TextMatrix(intRow, 6) = rstTrans!BillNo
                        .TextMatrix(intRow, 7) = Format(rstTrans!Quantity, "####.00")
                        .TextMatrix(intRow, 8) = Format(rstTrans!Amount, "######.00")
                        intRow = intRow + 1
                    End With
                Else
                    curQuantity = rstTrans!Quantity
                    intBillNo = rstTrans!BillNo
                    curAmount = rstTrans!Amount
                End If
                rstTrans.MoveNext
                If rstTrans.EOF Then
                    blnEof = True
                    rstTrans.MovePrevious
                    Exit Do
                End If
            Loop
            With msflxgrdProduct
                If (blnMore) Then
                    .TextMatrix(intRow, 5) = "Total"
                    .TextMatrix(intRow, 7) = Format(curQuantity, "####.00")
                    .TextMatrix(intRow, 8) = Format(curAmount, "######.00")
                End If
                If (intConRow >= intRow) Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(intConRow, 0) = strAccDesc
                If Not blnMore Then
                    .TextMatrix(intConRow, 1) = intBillNo
                End If
                .TextMatrix(intConRow, 2) = Format(curQuantity, "####.00")
                .TextMatrix(intConRow, 3) = Format(curAmount, "######.00")
                curTotal = curTotal + curAmount
                intConRow = intConRow + 1
            End With
            If blnEof Then
                Exit Do
            End If
        Loop
        msflxgrdProduct.TextMatrix(intConRow, 0) = "Total"
        msflxgrdProduct.TextMatrix(intConRow, 3) = Format(curTotal, "######.00")
    End If
End Sub
Private Sub ProductHeading()

    With msflxgrdProduct
        .Clear
        .Rows = 2
        .Cols = 9
        .FixedRows = 1
        .FixedCols = 0
        .ColWidth(0) = 2500
        .TextMatrix(0, 0) = "Description"
        .ColWidth(1) = 700
        .TextMatrix(0, 1) = "Bill No."
        .ColWidth(2) = 700
        .TextMatrix(0, 2) = "Quantity"
        .ColWidth(3) = 1250
        .TextMatrix(0, 3) = "Amount"
        .ColWidth(4) = 500
        .ColWidth(5) = 2500
        .TextMatrix(0, 5) = "Descriptioin"
        .ColWidth(6) = 700
        .TextMatrix(0, 6) = "Bill No."
        .ColWidth(7) = 700
        .TextMatrix(0, 7) = "Quantity"
        .ColWidth(8) = 1250
        .TextMatrix(0, 8) = "Amount"
    End With
End Sub
Private Sub DisplayingPersonal()
Dim strCond As String
Dim rstJournal As Recordset
Dim intRow As Integer
Dim strFindCond As String
Dim curTotal As Currency
        
    Call PersonalHeading
    strCond = "Select Acc_Code,BillNo,Amount from Journal " & _
            "Where ((Date = Cdate('" & mstrCurDate & "')) " & _
            " and (Acc_Code like 'A*')) Order By BillNo"
    Set rstJournal = mdbsAccounts.OpenRecordset(strCond, dbOpenSnapshot)
    If (Not rstJournal.BOF And Not rstJournal.EOF) Then
        rstJournal.MoveLast
        rstJournal.MoveFirst
        intRow = 1
        curTotal = 0
        Do While (Not rstJournal.EOF)
            With msflxgrdPersonal
                .Rows = intRow + 1
                strFindCond = "Acc_Code = '" & rstJournal!Acc_Code & "'"
                mrstAccCode.FindFirst strFindCond
                If (Not mrstAccCode.NoMatch) Then
                    .TextMatrix(intRow, 0) = mrstAccCode!Acc_Desc
                End If
                .TextMatrix(intRow, 1) = rstJournal!BillNo
                .TextMatrix(intRow, 2) = Format(rstJournal!Amount, "######.00")
                curTotal = curTotal + rstJournal!Amount
            End With
            rstJournal.MoveNext
            intRow = intRow + 1
        Loop
        With msflxgrdPersonal
            .Rows = intRow + 1
            .TextMatrix(intRow, 0) = "Total"
            .TextMatrix(intRow, 2) = Format(curTotal, "######.00")
        End With
    End If
End Sub
Private Sub DisplayingStock()
Dim strCond As String
Dim rstAccCode As Recordset
Dim rstCash As Recordset
Dim rstJournal As Recordset
    
Dim strSalesAccCode As String
Dim strPurAccCode As String
Dim strAccDesc As String
Dim intPos As Integer

Dim curOpenBalQty As Currency
Dim curSalesQty As Currency
Dim curPurQty As Currency
Dim curBal As Currency

Dim vntBookmark As Variant
Dim strFindCond As String
Dim intRow As Integer

    'displaying heading of the Stock Grid
    Call StockHeading
    'getting all product sales and purchase accounts
    'getting Code,Description and Year top quantity for all product
    strCond = "Select Acc_Code, Acc_Desc, YTop_Qty from AccCode " & _
                " Where ((Acc_Code like 'I*') or (Acc_Code like 'P*')) " & _
                "Order by Acc_Code"
    Set rstAccCode = mdbsAccounts.OpenRecordset(strCond, dbOpenSnapshot)
    If (Not rstAccCode.BOF And Not rstAccCode.EOF) Then
        intRow = 1
        'if product sales and purchase exists
        Do While (Left(rstAccCode!Acc_Code, 1) = "I")
            'loop for all Sales AccCode of all products (I)
            'Sales account starts with I
            'setting the sales and puchase code
            strSalesAccCode = rstAccCode!Acc_Code
            'setting the open balance as year top qty
            If (Not IsNull(rstAccCode!YTop_Qty)) Then
                curOpenBalQty = rstAccCode!YTop_Qty
            Else
                curOpenBalQty = 0
            End If
            'getting the letters except the first letter and adding letter p
            'purchase account starts with p
            strPurAccCode = "P" & Mid(strSalesAccCode, 2)
            'setting the bookmark to the current Sales account
            vntBookmark = rstAccCode.Bookmark
            'getting the purchase account for the product
            strFindCond = "Acc_Code = '" & strPurAccCode & "'"
            rstAccCode.FindFirst strFindCond
            If (Not rstAccCode.NoMatch) Then
                'if purchase account of the product exist
                'setting all quantities to 0
                curSalesQty = 0
                curPurQty = 0
                curBal = 0
                'getting the description of the product "Sales account(Chillies)"
                'getting only the name Chillies
                intPos = InStr(1, rstAccCode!Acc_Desc, "(")
                strAccDesc = Mid(rstAccCode!Acc_Desc, intPos + 1)
                intPos = InStr(1, strAccDesc, ")")
                strAccDesc = Mid(strAccDesc, 1, intPos - 1)
                
                
                If (gstrRegister = "Cash") Then
                    'getting the sales and purchase transactions for the product
                    'upto specified date from Cash
                    strCond = "Select Acc_Code, Quantity, Date from Cash " & _
                                " where (((Acc_Code = '" & strSalesAccCode & "') or " & _
                                " (Acc_Code = '" & strPurAccCode & "')) and " & _
                             " (Date <= Cdate('" & mstrCurDate & "'))) Order by Acc_Code"
                    Set rstCash = mdbsAccounts.OpenRecordset(strCond, dbOpenSnapshot)
                    If (Not rstCash.BOF And Not rstCash.EOF) Then
                        rstCash.MoveLast
                        rstCash.MoveFirst
                        Do While (Not rstCash.EOF)
                            If (rstCash!Acc_Code = strSalesAccCode) Then
                                If (CDate(rstCash!Date) = CDate(mstrCurDate)) Then
                                    curSalesQty = curSalesQty + rstCash!Quantity
                                Else
                                    curOpenBalQty = curOpenBalQty - rstCash!Quantity
                                End If
                            ElseIf (rstCash!Acc_Code = strPurAccCode) Then
                                If (CDate(rstCash!Date) = CDate(mstrCurDate)) Then
                                    curPurQty = curPurQty + rstCash!Quantity
                                Else
                                    curOpenBalQty = curOpenBalQty + rstCash!Quantity
                                End If
                            End If
                            rstCash.MoveNext
                        Loop
                    End If
                ElseIf (gstrRegister = "Credit") Then
                    'getting the sales and purchase transactions for the product
                    'upto specified date from Journal
                    strCond = "Select Acc_Code, Quantity, Date from Journal " & _
                                " where (((Acc_Code = '" & strSalesAccCode & "') or " & _
                                " (Acc_Code = '" & strPurAccCode & "')) and " & _
                                " (Date <= Cdate('" & mstrCurDate & "'))) Order by Acc_Code"
                    Set rstJournal = mdbsAccounts.OpenRecordset(strCond, dbOpenSnapshot)
                    If (Not rstJournal.BOF And Not rstJournal.EOF) Then
                        rstJournal.MoveLast
                        rstJournal.MoveFirst
                        Do While (Not rstJournal.EOF)
                            If (rstJournal!Acc_Code = strSalesAccCode) Then
                                If (CDate(rstJournal!Date) = CDate(mstrCurDate)) Then
                                    curSalesQty = curSalesQty + rstJournal!Quantity
                                Else
                                    curOpenBalQty = curOpenBalQty - rstJournal!Quantity
                                End If
                            ElseIf (rstJournal!Acc_Code = strPurAccCode) Then
                                If (CDate(rstJournal!Date) = CDate(mstrCurDate)) Then
                                    curPurQty = curPurQty + rstJournal!Quantity
                                Else
                                    curOpenBalQty = curOpenBalQty + rstJournal!Quantity
                                End If
                            End If
                            rstJournal.MoveNext
                        Loop
                    End If
                End If
                With msflxgrdStock
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = strAccDesc
                    .TextMatrix(intRow, 1) = Format(curOpenBalQty, "####.00")
                    .TextMatrix(intRow, 2) = Format(curPurQty, "####.00")
                    .TextMatrix(intRow, 3) = Format(curSalesQty, "####.00")
                    curBal = curOpenBalQty
                    curBal = curBal + curPurQty
                    curBal = curBal - curSalesQty
                    .TextMatrix(intRow, 4) = Format(curBal, "####.00")
                End With
                intRow = intRow + 1
            End If
            rstAccCode.Bookmark = vntBookmark
            rstAccCode.MoveNext
            If (rstAccCode.EOF) Then
                Exit Do
            End If
        Loop
    End If
        
End Sub
Private Sub StockHeading()

    With msflxgrdStock
        .Clear
        .Rows = 2
        .Cols = 5
        .FixedRows = 1
        .FixedCols = 0
        .ColWidth(0) = 2500
        .TextMatrix(0, 0) = "Description"
        .ColWidth(1) = 1200
        .TextMatrix(0, 1) = "Open Balance"
        .ColWidth(2) = 1200
        .TextMatrix(0, 2) = "Purchase"
        .ColWidth(3) = 1200
        .TextMatrix(0, 3) = "Sales"
        .ColWidth(4) = 1200
        .TextMatrix(0, 4) = "Net Balance"
    End With
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub mnuFirst_Click()
    
    'setting the start date as the current date
    mstrCurDate = gstrStDate
    'invoking the function to set the menu options according to the current date
    Call RecordPosition("F")
    'invoking function to display the transactions of the current date
    Call DisplayingAll
End Sub
Private Sub mnuLast_Click()
    'setting the end date as the current date
    mstrCurDate = gstrEndDate
    'invoking the function to set the menu options according to the current date
    Call RecordPosition("L")
    'invoking function to display the transactions of the current date
    Call DisplayingAll
End Sub
Private Sub mnuNext_Click()
    'setting the next date to the current date as the current date
    mstrCurDate = Format(DateAdd("d", 1, mstrCurDate), "mm/dd/yyyy")
    'invoking the function to set the menu options according to the current date
    If mstrCurDate = gstrEndDate Then
        Call RecordPosition("L")
    Else
        Call RecordPosition("M")
    End If
    'invoking function to display the transactions of the current date
    Call DisplayingAll
End Sub
Private Sub mnuPrev_Click()
    'setting the previous date to the current date as the current date
    mstrCurDate = Format(DateAdd("d", -1, mstrCurDate), "mm/dd/yyyy")
    'invoking the function to set the menu options according to the current date
    If mstrCurDate = gstrStDate Then
        Call RecordPosition("F")
    Else
        Call RecordPosition("M")
    End If
    'invoking function to display the transactions of the current date
    Call DisplayingAll
End Sub
Private Sub PersonalHeading()

    With msflxgrdPersonal
        .Clear
        .Rows = 2
        .Cols = 3
        .FixedRows = 1
        .FixedCols = 0
        .ColWidth(0) = 4500
        .TextMatrix(0, 0) = "Description"
        .ColWidth(1) = 1000
        .TextMatrix(0, 1) = "Bill No."
        .ColWidth(2) = 2000
        .TextMatrix(0, 2) = "Amount"
    End With
End Sub

