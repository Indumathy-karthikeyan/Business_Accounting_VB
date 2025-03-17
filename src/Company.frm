VERSION 5.00
Begin VB.Form frmCompany 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Company Details"
   ClientHeight    =   2280
   ClientLeft      =   1905
   ClientTop       =   4560
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   1695
      TabIndex        =   5
      Top             =   360
      Width           =   1515
   End
   Begin VB.Frame fraCompany 
      Caption         =   "Accounts for the Company"
      Height          =   1230
      Left            =   360
      TabIndex        =   2
      Top             =   900
      Width           =   2865
      Begin VB.OptionButton optSpma 
         Caption         =   "SP&MA"
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   870
      End
      Begin VB.OptionButton optSntc 
         Caption         =   "&SNTC"
         Height          =   285
         Left            =   330
         TabIndex        =   4
         Top             =   375
         Width           =   870
      End
      Begin VB.OptionButton optVpps 
         Caption         =   "&VPPS"
         Height          =   285
         Left            =   1665
         TabIndex        =   3
         Top             =   375
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3705
      TabIndex        =   1
      Top             =   675
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   3705
      TabIndex        =   0
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblYear 
      Caption         =   "Accounting Year"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   390
      Width           =   1290
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    'resetting the global variable since
    'the action was cancelled
    gstrAccYear = ""
    gstrCompanyCode = ""
    'unloading the form
    Unload Me
End Sub
Private Sub cmdOk_Click()
    'Validating the fields
    If txtYear.Text = "" Then
        MsgBox "Specify the Accounting year"
        txtYear.SetFocus
    ElseIf Not optSntc.Value And Not optVpps.Value And Not optSpma.Value Then
        MsgBox "Specify the Company Name"
        optSntc.SetFocus
    Else
        'setting the global variables to the selected datas
        gstrAccYear = txtYear.Text
        If optSntc.Value Then
            gstrCompanyCode = "SNTC"
        ElseIf optVpps.Value = True Then
            gstrCompanyCode = "VPPS"
        Else
            gstrCompanyCode = "SPMA"
        End If
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    'intializing the fields
    txtYear.Text = Year(Date) & "-" & Year(Date) + 1
    optSntc.Value = False
    optVpps.Value = False
    optSpma.Value = False
End Sub
