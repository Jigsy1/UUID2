VERSION 5.00
Begin VB.Form frmUUID2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UUID2"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAutomatic 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6000
      Top             =   1080
   End
   Begin VB.Timer tmrNoteClear 
      Interval        =   2000
      Left            =   5520
      Top             =   1080
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6840
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame fmeSettings 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox chkRandomness 
         Caption         =   "Increase randomness"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.CheckBox chkAutomatic 
         Caption         =   "Automatically generate (10s)"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.CheckBox chkBrackets 
         Caption         =   "Use {Braces}"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkUpper 
         Caption         =   "Use Uppercase"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   2295
      End
   End
   Begin VB.Frame fmeStyle 
      Height          =   735
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtUUID2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
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
  Dim baseString As String
  If chkUpper.Value = 1 Then
    baseString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Else
    baseString = "abcdefghijklmnopqrstuvwxyz"
  End If
  baseString = baseString & "0123456789"
  Randomize
  Dim loopNumber As Integer, outString As String, randNumber As Integer
  For loopNumber = 0 To 32 - 1
    If chkRandomness.Value = 1 Then
      randNumber = Int(Val(1 + Val(Rnd * 2)))
      If randNumber = 1 Then
        outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
      Else
        outString = outString & Mid(StrReverse(baseString), Int(Val(1 + Val(Rnd * Len(StrReverse(baseString))))), 1)
      End If
    Else
      outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
    End If
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
