VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00800080&
   Caption         =   "Server"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton meetme 
      Caption         =   "About Me"
      Height          =   855
      Left            =   6720
      TabIndex        =   16
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00800080&
      Caption         =   "Execute"
      Height          =   855
      Left            =   5040
      TabIndex        =   15
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   14
      Top             =   5400
      Width           =   4215
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   5880
      TabIndex        =   12
      Top             =   2400
      Width           =   1815
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   5880
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5880
      TabIndex        =   10
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton cmdMsgBox 
      BackColor       =   &H00800080&
      Caption         =   "Popup Message"
      Height          =   255
      Left            =   2400
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdCaption 
      BackColor       =   &H00800080&
      Caption         =   "Set client Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtReceived 
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   4455
   End
   Begin VB.TextBox txtSendMessage 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtErrors 
      Height          =   1335
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock wsArray 
      Index           =   0
      Left            =   3120
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2500
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   2280
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   2400
   End
   Begin VB.ListBox userlist 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800080&
      Caption         =   "Frame1"
      ForeColor       =   &H0000FFFF&
      Height          =   4575
      Left            =   5400
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Received"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Send Message"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Error Log"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Users"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************ÉLAN Softwares and Technologies**************

'This code is winsock control which can servers about 100 users.
'This code can be used to chat with , and to execute the file at the server side.
' Needs feedback for this please do so.
'As everything Over here is self explanatory , nothing needs explanation.



Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim Client(0 To 100) As String
Dim sss As String

Private Sub cmdCaption_Click()
Dim User As Integer
    User = RetrieveUser(userlist.Text)
    If User = -1 Then
        MsgBox "Invalid User!", vbCritical, "Error"
        Exit Sub
    End If
    wsArray(User).SendData "c" & Chr(1) & InputBox("What do you want to have their caption set to?", "Alter Caption", "Hi!")
End Sub

Private Sub cmdMsgBox_Click()
Dim User As Integer
    User = RetrieveUser(userlist.Text)
    If User = -1 Then
        MsgBox "Invalid User!", vbCritical, "Error"
        Exit Sub
    End If
    wsArray(RetrieveUser(userlist.Text)).SendData "m" & Chr(1) & InputBox("What do you want to have displayed on their machine?", "Popup MsgBox", "Hi!")
End Sub

Private Sub Command1_Click()
Call ShellExecute(0&, vbNullString, Text1.Text, vbNullString, _
vbNullString, vbNormalFocus)
End Sub

Private Sub Form_Load()
    Dir1.Path = "C:\"
    Dir1.Refresh
    wsListen.Listen  ' make it listen
End Sub

Private Sub meetme_Click()
Load Form3
Form3.Show 1
End Sub

Private Sub txtSendMessage_KeyDown(KeyCode As Integer, Shift As Integer)
Dim User As Integer
    If userlist.ListCount = 0 And KeyCode = 13 Then
        MsgBox "Nobody to send to!", vbExclamation, "Cannot send"
        txtSendMessage.Text = ""
        Exit Sub
    End If
 If KeyCode = 13 And Shift = 0 Then
        User = RetrieveUser(userlist.Text)
        If User = -1 Then
            Exit Sub
        End If
        wsArray(User).SendData "t" & Chr(1) & txtSendMessage.Text
        txtSendMessage.Text = ""
    ElseIf KeyCode = 13 And Shift = 1 Then
        For X = 0 To 100
            If Client(X) <> "" Then
                wsArray(X).SendData "t" & Chr(1) & txtSendMessage.Text
                DoEvents
            End If
        Next X
        txtSendMessage.Text = ""
    End If

End Sub

Private Function RetrieveUser(UserName As String) As Integer
    Dim X As Integer
    If UserName = "" Then
        If userlist.ListCount = 0 Then
            RetrieveUser = -1
            Exit Function
        End If
        UserName = userlist.List(0)
    End If
    For X = 0 To 100
        If Client(X) = UserName Then
            RetrieveUser = X
            Exit Function
        End If
    Next X
    RetrieveUser = -1
End Function

Private Sub txtSendMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub wsArray_Close(Index As Integer)
    For X = 0 To userlist.ListCount - 1
        If userlist.List(X) = Client(Index) Then
            Client(Index) = ""
            userlist.RemoveItem X
            Exit For
        End If
    Next X
End Sub

Private Sub wsArray_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Data As String, CtrlChar As String
    wsArray(Index).GetData Data
    If InStr(1, Data, Chr(1)) <> 2 Then
        MsgBox "Unknown Data Format: " & vbCrLf & _
                Data, vbCritical, "Error receiving"
        Exit Sub
    End If
    'Retrieve First Character
    CtrlChar = Left(Data, 1)
     'Make sure to trim it, and chr(1), off
    Data = Mid(Data, 3)
    Select Case LCase(CtrlChar)
        Case "m"
            MsgBox Data, vbInformation, "Msg from client"
        Case "c"
            Me.Caption = "Server - " & Data
        Case "u"
            userlist.AddItem Data
            Client(Index) = Data
        Case "f"
               msg = ""
               For i = 0 To Drive1.ListCount - 1
               msg = msg + Chr(1) + Drive1.List(i)
               Next
               wsArray(RetrieveUser(userlist.Text)).SendData "f" & Chr(1) & Drive1.ListCount & msg & Chr(1)
        Case "d"
                 On Error Resume Next
                 msg = ""
                File1.Refresh
                File1.Path = Data
                Dir1.Path = Data
                For i = 0 To Dir1.ListCount - 1
                    msg = msg + Chr(1) + Dir1.List(i)
                Next
                msg = msg + Chr(2)
                For i = 0 To File1.ListCount - 1
                    msg = msg + Chr(1) + File1.List(i)
                Next
                wsArray(RetrieveUser(userlist.Text)).SendData "d" & Chr(1) & File1.ListCount & Chr(1) & Dir1.ListCount & Chr(1) & msg & Chr(1)

               '  On Error Resume Next
               ' msg = ""
               ' Dir1.Refresh
               ' Dir1.Path = Data
               ' For i = 0 To Dir1.ListCount - 1
                '    msg = msg + Chr(1) + Dir1.List(i)
               ' Next
               ' wsArray(RetrieveUser(userlist.Text)).SendData "d" & Chr(1) & Dir1.ListCount & msg & Chr(1)
        Case "v"
                msg = ""
                File1.Refresh
                File1.Path = Data
                Dir1.Path = Data
                For i = 0 To Dir1.ListCount - 1
                    msg = msg + Chr(1) + Dir1.List(i)
                Next
                msg = msg + Chr(2)
                For i = 0 To File1.ListCount - 1
                    msg = msg + Chr(1) + File1.List(i)
                Next
                wsArray(RetrieveUser(userlist.Text)).SendData "v" & Chr(1) & File1.ListCount & Chr(1) & Dir1.ListCount & Chr(1) & msg & Chr(1)
        Case "s"
               
                Open Data For Binary As #2
                    Get #2, , sss
                Close #2
                wsArray(RetrieveUser(userlist.Text)).SendData "s" & Chr(1) & sss
        Case "e"
                Text1.Text = Data
                Command1_Click
        Case Else
            txtReceived.SelStart = Len(txtReceived.Text)
            txtReceived.SelText = Data & vbCrLf
    End Select
End Sub

Private Sub wsArray_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txtErrors.SelStart = Len(txtErrors.Text)
    txtErrors.SelText = "wsArray(" & Index & ") - " & Number & " - " & Description & vbCrLf
    wsArray(Index).Close
End Sub

Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
    Index = FindOpenWinsock
    ' Accept the request using the created winsock
    wsArray(Index).Accept requestID
End Sub

Private Sub wsListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txtErrors.SelStart = Len(txtErrors.Text)
    txtErrors.SelText = "wsListen - " & Number & " - " & Description & vbCrLf
End Sub

Private Function FindOpenWinsock()
Static LocalPorts As Integer
    For X = 0 To wsArray.UBound
        If wsArray(X).State = 0 Then
            FindOpenWinsock = X
            Exit Function
        End If
    Next X
    Load wsArray(wsArray.UBound + 1)
    LocalPorts = LocalPorts + 1
    wsArray(wsArray.UBound).LocalPort = wsArray(wsArray.UBound).LocalPort + LocalPorts
    FindOpenWinsock = wsArray.UBound
End Function
