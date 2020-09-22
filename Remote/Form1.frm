VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delorean Remote Controls"
   ClientHeight    =   3645
   ClientLeft      =   3345
   ClientTop       =   3645
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4935
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   2400
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   1395
      ScaleWidth      =   2115
      TabIndex        =   14
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Low"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   -120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "169.169.169.70"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "High"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   1815
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
      Begin VB.CommandButton Command8 
         Caption         =   "<"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "v"
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   ">"
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command5 
         Caption         =   "^"
         Height          =   375
         Left            =   720
         TabIndex        =   10
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "N"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Remote"
      Height          =   1095
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "When this baby hits 88mph, you are going to see some serious shit."
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "STATUS: NOT CONNECTED"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClear_Click()
Dim str1data As String
str1data = "N"
Winsock1.SendData str1data
End Sub

Private Sub Command1_Click()
Winsock1.Close
Winsock1.Connect Text1.Text, 2002
End Sub

Private Sub Command2_Click()
Dim str3data As String
str3data = "CLOSEME"
Winsock1.SendData str3data
Label4.Caption = "STATUS: Disconnected"
End Sub

Private Sub Command3_Click()
Dim str1data As String
str1data = "high"
Winsock1.SendData str1data
End Sub

Private Sub Command4_Click()
Dim str1data As String
str1data = "low"
Winsock1.SendData str1data
End Sub


Private Sub Command5_Click()
Dim str1data As String
str1data = "up"
Winsock1.SendData str1data
End Sub

Private Sub Command6_Click()
Dim str1data As String
str1data = "right"
Winsock1.SendData str1data
End Sub

Private Sub Command7_Click()
Dim str1data As String
str1data = "down"
Winsock1.SendData str1data
End Sub

Private Sub Command8_Click()
Dim str1data As String
str1data = "left"
Winsock1.SendData str1data
End Sub

Private Sub Winsock1_Connect()
'Me.Caption = "I think we're connected"
Label4.Caption = "STATUS:  CONNECTED!"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strdata1, strdata2, strdata3, strdata4 As String

Winsock1.GetData strdata1, vbString
Label4.Caption = "STATUS:  READY!"

End Sub
