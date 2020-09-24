VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Tutorial 13 - Client"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      Height          =   255
      Left            =   4920
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Data"
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Text            =   "socialvb.com is cool"
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Text            =   "6669"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.SendData Text3.Text 'sends data from the client to the server
End Sub

Private Sub Command2_Click()
With Winsock1
.Close 'closes the winsock incase it was open
.RemoteHost = Text1.Text 'declares text1 as the remote host (server)
.RemotePort = Text2.Text 'declares text2 as the port to connect to, default for this example is 6669
.Connect 'connects the client to the server using the defined data above
End With
End Sub
