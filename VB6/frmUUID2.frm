VERSION 5.00
Begin VB.Form frmUUID2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UUID2"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAutomatic 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6480
      Top             =   1080
   End
   Begin VB.Timer tmrNoteClear 
      Interval        =   2000
      Left            =   6000
      Top             =   1080
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Frame fmeSettings 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox chkRandomness 
         Caption         =   "Increased randomness"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.CheckBox chkAutomatic 
         Caption         =   "Automatically generate (10s)"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.CheckBox chkBrackets 
         Caption         =   "Use {Braces}"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkUpper 
         Caption         =   "Use Uppercase characters (A-Z)"
         Height          =   195
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   1  'Checked
         Width           =   2655
      End
   End
   Begin VB.Frame fmeStyle 
      Height          =   735
      Left            =   3120
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
  On Error GoTo endCopy
  Clipboard.Clear
  Clipboard.SetText txtUUID2.Text
  Me.Caption = Me.Caption & " - Copied to clipboard"
  tmrNoteClear.Enabled = True
  Exit Sub

endCopy:
  MsgBox "Failed to copy to clipboard.", vbExclamation, "Error"
End Sub

Private Sub cmdGenerate_Click()
  Call makeUUID2
End Sub

Private Function makeUUID2() As String
  Dim baseString As String
  baseString = "abcdefghijklmnopqrstuvwxyz"
  If chkUpper.Value = 1 Then baseString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  baseString = baseString & "0123456789"
  Randomize
  Dim loopNumber As Integer, outString As String, randNumber As Integer
  For loopNumber = 0 To 32 - 1
    If chkRandomness.Value = 1 Then
      randNumber = Int(Val(1 + Val(Rnd * 6)))
      Select Case randNumber
        Case 1
          outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
        Case 2
            outString = outString & Mid(StrReverse(baseString), Int(Val(1 + Val(Rnd * Len(StrReverse(baseString))))), 1)
        Case 3
          ' ,-> Former half
          outString = outString & Mid(Mid(baseString, 1, Val(Len(baseString) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(baseString, 1, Val(Len(baseString) / 2)))))), 1)
        Case 4
          ' ,-> Latter half
          outString = outString & Mid(Mid(baseString, Val(Len(baseString) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(baseString, Val(Len(baseString) / 2)))))), 1)
        Case 5
          ' ,-> Former half (Reverse)
          outString = outString & Mid(Mid(StrReverse(baseString), 1, Val(Len(StrReverse(baseString)) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(StrReverse(baseString), 1, Val(Len(StrReverse(baseString)) / 2)))))), 1)
        Case 6
          ' ,-> Latter half (Reverse)
          outString = outString & Mid(Mid(StrReverse(baseString), Val(Len(StrReverse(baseString)) / 2)), Int(Val(1 + Val(Rnd * Len(Mid(StrReverse(baseString), Val(Len(StrReverse(baseString)) / 2)))))), 1)
      End Select
    Else
      outString = outString & Mid(baseString, Int(Val(1 + Val(Rnd * Len(baseString)))), 1)
    End If
    If loopNumber = 7 Or loopNumber = 11 Or loopNumber = 15 Or loopNumber = 19 Then outString = outString & "-"
  Next
  If chkBrackets.Value = 1 Then outString = "{" & outString & "}"
  txtUUID2.Text = outString
  outString = ""
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then Call makeUUID2
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then End
End Sub

Private Sub Form_Load()
  Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
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
