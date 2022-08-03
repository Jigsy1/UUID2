VERSION 5.00
Begin VB.Form frmUUID2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UUID2"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAutomatic 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5160
      Top             =   1080
   End
   Begin VB.Timer tmrNoteClear 
      Interval        =   2000
      Left            =   4680
      Top             =   1080
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame fmeSettings 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox chkAutomatic 
         Caption         =   "Automatic (10s)"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox chkBrackets 
         Caption         =   "{Brackets}"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkUpper 
         Caption         =   "Uppercase letters"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.Frame fmeStyle 
      Height          =   735
      Left            =   2040
      TabIndex        =   7
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtUUID2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmUUID2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAutomatic_Click()
  If chkAutomatic.Value = 1 Then
    tmrAutomatic.Enabled = True
  Else
    tmrAutomatic.Enabled = False
  End If
End Sub

Private Sub cmdClose_Click()
  End
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText txtUUID2.Text
  Me.Caption = Me.Caption & " - Copied to clipboard"
  tmrNoteClear.Enabled = True
End Sub

Private Sub cmdGenerate_Click()
  Call makeUUID2
End Sub

Private Function makeUUID2() As String
  Randomize
  Dim baseString As String
  If chkUpper.Value = 1 Then
    baseString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Else
    baseString = "abcdefghijklmnopqrstuvwxyz"
  End If
  baseString = baseString & "0123456789"
  Dim loopNumber As Integer, outString As String
  For loopNumber = 0 To 32 - 1
    outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
    If loopNumber = 7 Or loopNumber = 11 Or loopNumber = 15 Or loopNumber = 19 Then
      outString = outString & "-"
    End If
  Next
  If chkBrackets.Value = 1 Then
    outString = "{" & outString & "}"
  End If
  txtUUID2.Text = outString
  outString = ""
End Function

Private Sub Form_Load()
  Me.Tag = Me.Caption
  Call makeUUID2
End Sub

Private Sub Form_Terminate()
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
  ' `-> I doubt this has any use; but just incase...
End Sub

Private Sub tmrAutomatic_Timer()
  Call makeUUID2
End Sub

Private Sub tmrNoteClear_Timer()
  Me.Caption = Me.Tag
  tmrNoteClear.Enabled = False
End Sub

' EOF
