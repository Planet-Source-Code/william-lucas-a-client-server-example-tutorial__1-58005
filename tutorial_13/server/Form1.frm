VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Tutorial 13 - Server"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Incoming Data:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Host:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
With Winsock1
.Close 'incase winsock1 is open for any reason
.LocalPort = 6669 'defines the port to monitor for incoming connections
.Listen 'starts the listening process
End With
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close 'closes winsock1 if it was open
Winsock1.Accept (requestID) 'accepts the incoming connection from the client
Text1.Text = Winsock1.RemoteHostIP 'sets text1 to the ip of the connected client
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim thedata As String 'dims the data as a string
Winsock1.GetData thedata 'gets the incoming data from the client
Text2.Text = thedata 'prints the data in text2
End Sub

