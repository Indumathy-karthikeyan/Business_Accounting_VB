VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2940
   ClientLeft      =   2160
   ClientTop       =   2520
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5445
   Begin VB.Menu mnuEntry 
      Caption         =   "&Entry"
      Begin VB.Menu mnuCoding 
         Caption         =   "C&oding"
         Begin VB.Menu mnuAccount 
            Caption         =   "&Account"
         End
         Begin VB.Menu mnuNarration 
            Caption         =   "&Narration"
         End
      End
      Begin VB.Menu mnuCash 
         Caption         =   "&Cash"
      End
      Begin VB.Menu mnuJournal 
         Caption         =   "&Journal"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
         Begin VB.Menu mnuCashSales 
            Caption         =   "Cash"
         End
         Begin VB.Menu mnuCreditSales 
            Caption         =   "Credit"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuDayBook 
         Caption         =   "&DayBook"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "&Ledger"
      End
      Begin VB.Menu mnuTrialBalance 
         Caption         =   "&Trial Balance"
         Begin VB.Menu mnuAssLiab 
            Caption         =   "&Assets and Liabilities"
         End
         Begin VB.Menu mnuNomAcc 
            Caption         =   "&Nominal Accounts"
         End
      End
      Begin VB.Menu mnuSalesReg 
         Caption         =   "Sales Register"
         Begin VB.Menu mnuCashRegister 
            Caption         =   "Cash"
         End
         Begin VB.Menu mnuCreditRegister 
            Caption         =   "Credit"
         End
      End
   End
   Begin VB.Menu mnuQuit 
      Caption         =   "&Quit"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnStart As Boolean    'To store the starting status
Private Sub Form_Activate()
Dim strYear As String   'To store the starting year
Dim strDb As String
    If mblnStart Then
        'Invoking company form to get the company information
        frmCompany.Show vbModal
        If (gstrCompanyCode <> "") Then
            'To get the starting year of the Accouting year
            strYear = Mid(gstrAccYear, 1, 4)
            'To get the Location of the Database
            gstrDatabase = Left(App.Path, (Len(App.Path) - 8)) & "\Yr" & strYear
            'selects database according to the company choosen
            'and sets the forms caption accordingly
            If gstrCompanyCode = "SNTC" Then
                gstrDatabase = gstrDatabase & "\sntc.mdb"
                Me.Caption = "Sri Narayana Trading Company"
            ElseIf gstrCompanyCode = "VPPS" Then
                gstrDatabase = gstrDatabase & "\Vpps.mdb"
                Me.Caption = "V.P. Pachaiyappa Mudliar And Sons"
            ElseIf gstrCompanyCode = "SPMA" Then
                gstrDatabase = gstrDatabase & "\Spma.mdb"
                Me.Caption = "Sri Pachaiyappa Marketing Agencies"
            ElseIf gstrCompanyCode = "" Then
                Unload Me
            End If
            
            
            
            DataEnvironment1.Accounts.Open gstrDatabase, "Admin", ""
            'resets the starting flag
            mblnStart = False
        Else
            Unload Me
        End If
    End If
End Sub
Private Sub Form_Load()
    'sets the starting flag
    mblnStart = True
End Sub
Private Sub mnuAccount_Click()
    'invokes the Account Details form
    frmAccDet.Show
End Sub
Private Sub mnuAssLiab_Click()
    mnuAssLiab.Tag = "Prepare"
    frmTrialBalance.Show vbModal
    mnuAssLiab.Tag = ""
End Sub

Private Sub mnuCash_Click()
    frmCash.Show
End Sub

Private Sub mnuCashRegister_Click()
    gstrRegister = "Cash"
    gstrGetInfo = "Sales Register"
    frmInformation.Show vbModal
    gstrGetInfo = ""
    gstrStCode = ""
    gstrEndCode = ""
    gstrStDate = ""
    gstrEndDate = ""
End Sub

Private Sub mnuCashSales_Click()
    gstrRegister = "Cash Sales"
    frmSales.Show vbModal
End Sub

Private Sub mnuCredit_Click()

End Sub

Private Sub mnuCreditRegister_Click()
    gstrRegister = "Credit"
    gstrGetInfo = "Sales Register"
    frmInformation.Show vbModal
    gstrGetInfo = ""
    gstrStCode = ""
    gstrEndCode = ""
    gstrStDate = ""
    gstrEndDate = ""
End Sub

Private Sub mnuCreditSales_Click()
    gstrRegister = "Credit Sales"
    frmSales.Show vbModal
End Sub

Private Sub mnuDayBook_Click()
    
    gstrGetInfo = "Day Book"
    frmInformation.Show vbModal
    gstrGetInfo = ""
    gstrStCode = ""
    gstrEndCode = ""
    gstrStDate = ""
    gstrEndDate = ""
End Sub

Private Sub mnuJournal_Click()
    frmJournal.Show
End Sub


Private Sub mnuLedger_Click()

    gstrGetInfo = "Ledger"
    frmInformation.Show vbModal
    gstrGetInfo = ""
    gstrStCode = ""
    gstrEndCode = ""
    gstrStDate = ""
    gstrEndDate = ""
End Sub

Private Sub mnuNarration_Click()
    frmNarration.Show
End Sub


Private Sub mnuNomAcc_Click()
    mnuNomAcc.Tag = "Prepare"
    frmTrialBalance.Show vbModal
    mnuNomAcc.Tag = ""
End Sub

Private Sub mnuQuit_Click()
    Unload Me
    End
End Sub



Private Sub mnuStock_Click()

End Sub

