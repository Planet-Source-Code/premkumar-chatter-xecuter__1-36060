VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00400040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information about me..."
   ClientHeight    =   3405
   ClientLeft      =   4695
   ClientTop       =   2865
   ClientWidth     =   4575
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "EXIT"
      Height          =   1095
      Left            =   3840
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00400040&
      Caption         =   "premon24@rediffmail.com"
      ForeColor       =   &H00FFFF00&
      Height          =   195
      Left            =   1200
      TabIndex        =   5
      Top             =   3120
      Width           =   1830
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   1200
      Picture         =   "info.frx":0000
      Top             =   840
      Width           =   2130
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400040&
      Caption         =   "Attention : This Product is a Duly Licensed Software of Mr.PremKumar of ÉLAN SOFTWARES & TECHNOLOGIES. I can be reached at"
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3600
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ver. :1.0.0 //Jan '02"
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1650
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   135
      X2              =   4560
      Y1              =   2175
      Y2              =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400040&
      Caption         =   "Its a 32 Bit Application Software , copyrighted to"
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "xecuter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************ÉLAN Softwares and Technologies**************

Private Sub Command1_Click()
Form3.Hide
Unload Form3
End Sub
