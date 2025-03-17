VERSION 5.00
Begin VB.Form frmInformation 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3855
   ClientLeft      =   2700
   ClientTop       =   2790
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6150
   Begin VB.Frame fraInformation 
      Height          =   3375
      Left            =   240
      TabIndex        =   10
      Top             =   240
      Width           =   5655
      Begin VB.TextBox txtStartDate 
         Height          =   300
         Left            =   1440
         TabIndex        =   5
         Top             =   1860
         Width           =   1290
      End
      Begin VB.TextBox txtEndDate 
         Height          =   300
         Left            =   4050
         TabIndex        =   7
         Top             =   1860
         Width           =   1290
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
         Height          =   285
         Left            =   1620
         TabIndex        =   8
         Top             =   2730
         Width           =   930
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   285
         Left            =   3090
         TabIndex        =   9
         Top             =   2730
         Width           =   930
      End
      Begin VB.TextBox txtEndCode 
         Height          =   300
         Left            =   4050
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1200
         Width           =   1290
      End
      Begin VB.TextBox txtStartCode 
         Height          =   300
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   1
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label lblInfo 
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   11
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         Caption         =   "Start Date"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1890
         Width           =   1020
      End
      Begin VB.Label lblInfo 
         Caption         =   "End Date"
         Height          =   255
         Index           =   4
         Left            =   3045
         TabIndex        =   6
         Top             =   1890
         Width           =   960
      End
      Begin VB.Label lblInfo 
         Caption         =   "Ending Code"
         Height          =   255
         Index           =   2
         Left            =   3045
         TabIndex        =   2
         Top             =   1230
         Width           =   960
      End
      Begin VB.Label lblInfo 
         Caption         =   "Starting Code"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   0
         Top             =   1230
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    gstrStCode = ""
    gstrEndCode = ""
    gstrStDate = ""
    gstrEndDate = ""
    gstrGetInfo = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()

    Dim dbsAccounts As Database
    Dim rstCode As Recordset
    
    If FieldCheck Then
        If gstrGetInfo = "Ledger" Then
            Set dbsAccounts = OpenDatabase(gstrDatabase)
            If Len(txtStartCode.Text) = 4 Then
                Set rstCode = dbsAccounts.OpenRecordset("Select Acc_Code " _
                        & "from AccCode where Acc_Code = '" _
                        & txtStartCode.Text & "'", dbOpenSnapshot)
                If rstCode.EOF Then
                    MsgBox "Invalid Starting Account Code"
                    txtStartCode.SetFocus
                    Exit Sub
                End If
                rstCode.Close
            End If
            gstrStCode = txtStartCode.Text
            If txtEndCode.Text <> "" Then
                If Len(txtEndCode.Text) = 4 Then
                    Set rstCode = dbsAccounts.OpenRecordset("Select " _
                            & "Acc_Code from AccCode where Acc_Code = '" _
                            & txtEndCode.Text & "'", dbOpenSnapshot)
                    If rstCode.EOF Then
                        MsgBox "Invalid Ending Account Code"
                        txtEndCode.SetFocus
                        Exit Sub
                    End If
                    rstCode.Close
                End If
                gstrEndCode = txtEndCode.Text
            Else
                gstrEndCode = gstrStCode
            End If
        ElseIf gstrGetInfo = "Day Book" Then
            gstrStCode = ""
            gstrEndCode = ""
        End If
        gstrStDate = Format(txtStartDate.Text, "mm/dd/yyyy")
        If txtEndDate.Text = "" Then
            gstrEndDate = gstrStDate
        Else
            gstrEndDate = Format(txtEndDate.Text, "mm/dd/yyyy")
        End If
        'gintPageNo = txtPageNo.Text
        Unload Me
        If gstrGetInfo = "Ledger" Then
            frmLedger.Show vbModal
        ElseIf gstrGetInfo = "Day Book" Then
            frmDayBook.Show vbModal
        ElseIf gstrGetInfo = "Sales Register" Then
            Form1.Show vbModal
        End If
    End If

End Sub



Private Sub Form_Load()
    If gstrGetInfo = "Day Book" Or gstrGetInfo = "Sales Register" Then
        Me.Caption = "Day Book Information"
        lblInfo(0).Caption = "Day Book Entries"
        
        lblInfo(1).Visible = False
        txtStartCode.Visible = False
        lblInfo(2).Visible = False
        txtEndCode.Visible = False
        
        lblInfo(3).Top = lblInfo(3).Top - 500
        txtStartDate.Top = txtStartDate.Top - 500
        lblInfo(4).Top = lblInfo(4).Top - 500
        txtEndDate.Top = txtEndDate.Top - 500
        cmdOk.Top = cmdOk.Top - 500
        cmdCancel.Top = cmdCancel.Top - 500
        fraInformation.Height = fraInformation.Height - 500
        Me.Height = Me.Height - 500
    ElseIf gstrGetInfo = "Ledger" Then
        Me.Caption = "Ledger Information"
        lblInfo(0).Caption = "Ledger Entries"
    End If
End Sub







Private Sub txtEndCode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            Beep
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Len(txtEndCode.Text) = 0 Then
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


Private Sub txtEndDate_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) And (KeyAscii <> Asc("/")) And _
        (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        KeyAscii = 0
        Beep
    End If
End Sub


Private Sub txtStartCode_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 8 Then
        If (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
            (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
            (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
            Beep
            KeyAscii = 0
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If Len(txtStartCode.Text) = 0 Then
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


Private Sub txtStartDate_KeyPress(KeyAscii As Integer)
    If (KeyAscii <> 8) And (KeyAscii <> Asc("/")) And _
        (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) Then
        KeyAscii = 0
        Beep
    End If
End Sub



Private Function FieldCheck() As Boolean
    FieldCheck = False
    If gstrGetInfo = "Ledger" Then
        If txtStartCode.Text = "" Then
            MsgBox "Starting Code should not be blank"
            txtStartCode.SetFocus
            Exit Function
        End If
    End If
    If txtStartDate.Text = "" Then
        MsgBox "Starting date should not be blank"
        txtStartDate.SetFocus
    Else
        FieldCheck = True
    End If
        
End Function


