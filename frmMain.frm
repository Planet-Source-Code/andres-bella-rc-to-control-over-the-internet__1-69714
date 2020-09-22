VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delorean Controls"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   360
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   8
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.PictureBox picOne 
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      Picture         =   "frmMain.frx":179E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picZero 
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      Picture         =   "frmMain.frx":1A23
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picBit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   3
      Left            =   3240
      Picture         =   "frmMain.frx":1C5C
      ScaleHeight     =   390
      ScaleWidth      =   390
      TabIndex        =   3
      Top             =   960
      Width           =   420
   End
   Begin VB.PictureBox picBit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   2
      Left            =   4200
      Picture         =   "frmMain.frx":1E95
      ScaleHeight     =   390
      ScaleWidth      =   390
      TabIndex        =   2
      Top             =   960
      Width           =   420
   End
   Begin VB.PictureBox picBit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   1
      Left            =   3720
      Picture         =   "frmMain.frx":20CE
      ScaleHeight     =   390
      ScaleWidth      =   390
      TabIndex        =   1
      Top             =   1440
      Width           =   420
   End
   Begin VB.PictureBox picBit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Index           =   0
      Left            =   3720
      Picture         =   "frmMain.frx":2307
      ScaleHeight     =   390
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   480
      Width           =   420
   End
   Begin VB.Frame Frame1 
      Caption         =   "Controls"
      Height          =   1935
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   360
      Picture         =   "frmMain.frx":2540
      ScaleHeight     =   1755
      ScaleWidth      =   2355
      TabIndex        =   10
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long

Private Sub Form_Load()

Winsock1.LocalPort = 2002 'this will show you the port where you
Winsock1.Listen           'are connected to

End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close

Winsock1.Accept requestID
'This will show you if is connected in the remote
End Sub

Private Sub cmdClear_Click()
    Call vbOut(888, 0) 'this will call the parallel port
    Call ClearBits
End Sub

Private Sub ClearBits()
        For i% = 0 To 3 'this will clear the parallel port
            picBit(i).Picture = picZero.Picture
        Next i
End Sub


Private Sub cmdExit_Click()
    End 'to exit the program
End Sub


Private Sub cmdSend_Click()
    Dim dec As Integer 'this are the controllers
          
        Call vbOut(888, dec)
    
        Call ClearBits

    
        If (dec >= 8) Then
            picBit(3).Picture = picOne.Picture
            dec = dec - 8
        End If
    
        If (dec >= 4) Then
            picBit(2).Picture = picOne.Picture
            dec = dec - 4
        End If
    
        If (dec >= 2) Then
            picBit(1).Picture = picOne.Picture
            dec = dec - 2
        End If
    
        If (dec >= 1) Then
            picBit(0).Picture = picOne.Picture
            dec = dec - 1
        End If
   
End Sub



Private Sub picBit_Click(Index As Integer)
    Dim SendBit As Integer 'this change the pictute in the controllers
        
    If (picBit(Index).Picture = picOne.Picture) Then
        picBit(Index).Picture = picZero.Picture
    Else
        picBit(Index).Picture = picOne.Picture
    End If
    SendBit = My_BinToDec()
  
    
    Call vbOut(888, SendBit)
End Sub



Private Function My_BinToDec() As Integer
    Dim DoubleByte As Integer 'i dunno what this does
    
    For i% = 0 To 3
        If (picBit(i).Picture = picOne.Picture) Then
            DoubleByte = DoubleByte + (2 ^ i)
        End If
    Next i

    My_BinToDec = DoubleByte
End Function

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)


DoEvents               'from here you will see the events you
Dim strdata As String
Dim dec As Integer   'want to control
Call Winsock1.GetData(strdata$, vbString)


DoEvents
If strdata = "low" Then
Picture1.Visible = True
Picture2.Visible = False
ElseIf strdata = "high" Then
Picture1.Visible = False
Picture2.Visible = True

ElseIf strdata = "N" Then
picBit(0).Picture = picOne.Picture
Call vbOut(888, 0)
Call ClearBits

ElseIf strdata = "up" Then
picBit(0).Picture = picOne.Picture
Call vbOut(888, 1)

ElseIf strdata = "down" Then
picBit(1).Picture = picOne.Picture
Call vbOut(888, 2)


ElseIf strdata = "left" Then
picBit(2).Picture = picOne.Picture
Call vbOut(888, 4)

ElseIf strdata = "right" Then
picBit(3).Picture = picOne.Picture
Call vbOut(888, 8)

End If
End Sub
