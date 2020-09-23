VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "IP Catcher"
   ClientHeight    =   2175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2175
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2280
      Top             =   360
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1680
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "http://www.yahoo.com/"
      Top             =   720
      Width           =   2535
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Page Not Found"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "80"
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Activate IP Catcher on port "
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Send Users To:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then Winsock1.LocalPort = Val(Text1): Winsock1.Listen
If Check1.Value = 0 Then Winsock1.Close
End Sub

Private Sub Text2_GotFocus()
Option2.Value = True
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
If Option2.Value = True Then
  headers = "HTTP/1.0 200 OK" & vbCrLf & "Date: Sat, 23 Feb 2002 02:36:17 GMT" & vbCrLf & "Connection: Close" & vbCrLf & "Content-Type: text/html" & vbCrLf & vbCrLf
  body = "<SCRIPT LANGUAGE=JavaScript>location.href=" & """" & Text2 & """" & ";</SCRIPT>"
  tosend = headers & body
  Winsock1.SendData tosend
  DoEvents
End If
List1.AddItem (Winsock1.RemoteHostIP)
Winsock1.Close
Winsock1.LocalPort = Val(Text1)
Winsock1.Listen
End Sub
