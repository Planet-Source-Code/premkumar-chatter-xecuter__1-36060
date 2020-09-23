VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   Caption         =   "Client"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5505
   ScaleWidth      =   8775
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton meetme 
      Caption         =   "ABOUT ME"
      Height          =   855
      Left            =   1200
      TabIndex        =   16
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Top             =   4560
      Width           =   3615
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   2535
      Left            =   9000
      TabIndex        =   14
      Top             =   480
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4471
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Execute"
      Height          =   855
      Left            =   7440
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dia 
      Left            =   480
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Inet1 
      Height          =   120
      Left            =   120
      ScaleHeight     =   60
      ScaleWidth      =   180
      TabIndex        =   0
      Top             =   4680
      Width           =   240
   End
   Begin MSWinsockLib.Winsock wsMain 
      Left            =   480
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.2"
      RemotePort      =   2400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800080&
      Caption         =   "REMOTE FILE EXECUTION"
      ForeColor       =   &H0000FFFF&
      Height          =   4095
      Left            =   5040
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.CommandButton Command2 
         Caption         =   "File Details"
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ListBox List2 
         Height          =   3570
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800080&
      Caption         =   "CHATTING"
      ForeColor       =   &H0000FFFF&
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtReceived 
         Height          =   1575
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtMessage 
         Enabled         =   0   'False
         Height          =   1125
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtUserName 
         Height          =   405
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Connect"
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Received From Server"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Send Message"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************ÉLAN Softwares and Technologies**************
'Every thing is self explanatory
'Feedback is needed


Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim node1 As Node
Dim node2 As Node
Dim filename As String
Dim node3 As Node
Dim nodeno As Integer

Private Sub Combo1_DblClick()
wsMain.SendData "d" & Chr(1) & Left(Combo1.Text, 2)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
wsMain.SendData "d" & Chr(1) & Left(Combo1.Text, 2)
End Sub

Private Sub Command1_Click()
    If txtUserName.Text = "" Then
        MsgBox "You need to type your username!", vbCritical, "Unable to complete"
        Exit Sub
    End If
    wsMain.Connect
    Do Until wsMain.State = 7
        If wsMain.State = 0 Or wsMain.State = 9 Then
            MsgBox "Error in connecting!", vbCritical, "Winsock Error"
            Exit Sub
        End If
        DoEvents
    Loop
    wsMain.SendData "U" & Chr(1) & txtUserName.Text
    txtUserName.Enabled = False
    txtMessage.Enabled = True
End Sub
Private Sub Command1_KeyPress(KeyAscii As Integer)
wsMain.SendData "d" & Chr(1) & Left(Combo1.Text, 2)
End Sub

Private Sub Command2_Click()
wsMain.SendData "f" & Chr(1) & "drives"
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
MsgBox State
End Sub
Private Sub Command4_Click()
' no need here
'Call ShellExecute(0&, vbNullString, filename, vbNullString, _
'vbNullString, vbNormalFocus)
If TreeView1.SelectedItem <> "" Then
    filename = TreeView1.SelectedItem & "\" & List2.Text
Else
    filename = Left(Combo1.Text, 2) & "\" & List2.Text
End If
wsMain.SendData "e" & Chr(1) & filename
End Sub
Private Sub List2_Click()
filename = List2.Text
If TreeView1.SelectedItem <> "" Then
    wsMain.SendData "s" & Chr(1) & TreeView1.SelectedItem & "\" & List2.Text
    filename = TreeView1.SelectedItem & "\" & List2.Text
Else
    wsMain.SendData "s" & Chr(1) & Left(Combo1.Text, 2) & "\" & List2.Text
    filename = Left(Combo1.Text, 2) & "\" & List2.Text
End If
End Sub

Private Sub meetme_Click()
Load Form3
Form3.Show 1
End Sub

'not needed here
Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
'wsMain.SendData "v" & Chr(1) & TreeView1.SelectedItem
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
wsMain.SendData "v" & Chr(1) & TreeView1.SelectedItem
nodeno = Node.Index
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        wsMain.SendData "t" & Chr(1) & txtMessage.Text
        txtMessage.Text = ""
        KeyAscii = 0
    End If
End Sub

Private Sub wsMain_DataArrival(ByVal bytesTotal As Long)
Dim Data As String, CtrlChar As String
Dim no, j As Integer
    wsMain.GetData Data
    CtrlChar = Left(Data, 1)
    Data = Mid(Data, 3)
    Select Case LCase(CtrlChar)
        Case "m"
            MsgBox Data, vbInformation, "Msg from server"
        Case "c"
            Me.Caption = "Client - " & Data
        Case "f"
                Combo1.Clear
                no = Left(Data, 1)
                Data = Right(Data, Len(Data) - 2)
                For i = 0 To no - 1
                    j = InStr(1, Data, Chr(1))
                    Combo1.AddItem Left(Data, j - 1)
                    Data = Right(Data, Len(Data) - j)
                Next
        Case "d"
                
                TreeView1.Nodes.Clear
                'TreeView1.SelectedItem = ""
                List2.Clear
                j = InStr(1, Data, Chr(1))
                no = Left(Data, j - 1)
                Data = Right(Data, Len(Data) - j)
                j = InStr(1, Data, Chr(1))
                nod = Left(Data, j - 1)
                Data = Right(Data, Len(Data) - j - 1)
                
                For i = 0 To nod - 1
                    j = InStr(1, Data, Chr(1))
                   ' List1.AddItem Left(Data, j - 1)
                    Set node1 = TreeView1.Nodes.Add(, , , Left(Data, j - 1))
                    Data = Right(Data, Len(Data) - j)
                Next
                For i = 0 To no - 1
                    j = InStr(1, Data, Chr(1))
                        List2.AddItem Left(Data, j - 1)
                    Data = Right(Data, Len(Data) - j)
                Next
                Data = ""
         Case "v"
                'TreeView1.SelectedItem = ""
                List2.Clear
                'List1.Clear
                j = InStr(1, Data, Chr(1))
                no = Left(Data, j - 1)
                Data = Right(Data, Len(Data) - j)
                j = InStr(1, Data, Chr(1))
                nod = Left(Data, j - 1)
                Data = Right(Data, Len(Data) - j - 1)
                
                For i = 0 To nod - 1
                    j = InStr(1, Data, Chr(1))
                    'List1.AddItem Left(Data, j - 1)
                    Set node2 = TreeView1.Nodes.Add(nodeno, tvwChild, , Left(Data, j - 1))
                    Data = Right(Data, Len(Data) - j)
                Next
                For i = 0 To no - 1
                    j = InStr(1, Data, Chr(1))
                        List2.AddItem Left(Data, j - 1)
                    Data = Right(Data, Len(Data) - j)
                Next
                Data = ""
        Case "s"
                'This code enables you to download the files
                ' some bug is there it seems
                If Data <> "" Then
                Dim fil As String
                On Error Resume Next
                dia.DialogTitle = "Save As"
                dia.filename = ""
                dia.CancelError = True
                fil = "*." & Right(List2.Text, Len(List2.Text) - InStr(1, List2.Text, "."))
                dia.DefaultExt = fil
                dia.Filter = "ALL FILES(*.*)|*.*"
                dia.Action = 2
                If dia.filename <> "" Then
                 Open "d:\pp.txt" For Binary As #3
                 Put #3, 1, "skfhlskjdflijgoisd"
                Close #3
                End If
                filename = dia.filename
                Text1.Text = filename
             End If
            Case Else
            txtReceived.SelStart = Len(txtReceived.Text)
            txtReceived.SelText = Data & vbCrLf
    End Select
End Sub

Private Sub wsMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock Error: " & Number & vbCrLf & Description, vbCritical, "Winsock Error"
End Sub
