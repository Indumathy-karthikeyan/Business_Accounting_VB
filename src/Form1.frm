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
      DatabaseName    =   "C:\ACCOUNTS\YR2001\SNTC.MDB"
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
      DatabaseName    =   "C:\ACCOUNTS\YR2001\SNTC.MDB"
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
      Bindings        =   "Form1.frx":0000
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "Form1.frx":0019
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin MSDBGrid.DBGrid dbgrdNarration 
      Bindings        =   "Form1.frx":0A11
      Height          =   2415
      Left            =   240
      OleObjectBlob   =   "Form1.frx":0A2C
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
Private Sub Form_Load()

End Sub
