VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDetails 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3435
   ClientLeft      =   2160
   ClientTop       =   1545
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data datAccCode 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "AccCode"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2145
      TabIndex        =   3
      Top             =   2880
      Width           =   885
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   3495
      TabIndex        =   2
      Top             =   2880
      Width           =   885
   End
   Begin VB.Data datNarration 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4755
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "NarrCode"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1155
   End
   Begin MSDBGrid.DBGrid dbgrdAccCode 
      Bindings        =   "Details.frx":0000
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "Details.frx":0019
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin MSDBGrid.DBGrid dbgrdNarration 
      Bindings        =   "Details.frx":0A11
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "Details.frx":0A2C
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mdbsAccounts As Database
Dim mrstDetails As Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()

    If gstrDetName = "Account" Then
        If gstrFormName = "Journal" Then
            frmJournal.txtAccCode.Text = mrstDetails!Acc_Code
            frmJournal.txtAccDesc.Text = mrstDetails!Acc_Desc
            frmJournal.txtAccCode.Tag = ""
        ElseIf gstrFormName = "Cash" Then
            frmCash.txtAccCode.Text = mrstDetails!Acc_Code
            frmCash.txtAccDesc.Text = mrstDetails!Acc_Desc
        ElseIf gstrFormName = "Sales AccCode" Then
            frmSales.txtAccCode.Text = mrstDetails!Acc_Code
            frmSales.txtAccDesc.Text = mrstDetails!Acc_Desc
        ElseIf gstrFormName = "Sales ProdAccCode" Then
            frmSales.txtProdAccCode.Text = mrstDetails!Acc_Code
            frmSales.txtProdAccDesc.Text = mrstDetails!Acc_Desc
        End If
    ElseIf gstrDetName = "Narration" Then
        If gstrFormName = "Journal" Then
            'frmJournal.txtNarrCode.Visible = False
            'frmJournal.txtNarration.Visible = True
            frmJournal.txtNarration.Text = mrstDetails!Narr_Desc
        ElseIf gstrFormName = "Cash" Then
            frmCash.txtNarration.Text = mrstDetails!Narr_Desc
        ElseIf gstrFormName = "Sales Narration" Then
            frmSales.txtNarration.Text = mrstDetails!Narr_Desc
        ElseIf gstrFormName = "Sales ProdNarration" Then
            frmSales.txtProdNarration.Text = mrstDetails!Narr_Desc
        End If
    End If
    Unload Me
End Sub

Private Sub Form_Load()

    Set mdbsAccounts = OpenDatabase(gstrDatabase)
    If gstrDetName = "Account" Then
        Me.Caption = "Account Details"
        If (gstrFormName = "Sales AccCode") Then
            Set mrstDetails = mdbsAccounts.OpenRecordset("Select Acc_Code, Acc_Desc from AccCode where Acc_Code like 'A*'", dbOpenSnapshot)
        ElseIf (gstrFormName = "Sales ProdAccCode") Then
            Set mrstDetails = mdbsAccounts.OpenRecordset("Select Acc_Code, Acc_Desc from AccCode where Acc_Code like 'I*' or Acc_Code like 'E*'", dbOpenSnapshot)
        Else
            Set mrstDetails = mdbsAccounts.OpenRecordset("Select Acc_Code, Acc_Desc, Qty_toRec from AccCode", dbOpenSnapshot)
        End If
        Set datAccCode.Recordset = mrstDetails
        
        datNarration.Visible = False
        dbgrdNarration.Visible = False
    ElseIf gstrDetName = "Narration" Then
        Me.Caption = "Narration Details"
        Set mrstDetails = mdbsAccounts.OpenRecordset("Select Narr_Code,Narr_Desc from NarrCode", dbOpenSnapshot)
        Set datNarration.Recordset = mrstDetails
        
        datAccCode.Visible = False
        dbgrdAccCode.Visible = False
    End If
End Sub



