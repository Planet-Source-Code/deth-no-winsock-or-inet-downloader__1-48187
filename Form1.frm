VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Error"
      Height          =   330
      Left            =   2340
      TabIndex        =   6
      Top             =   2250
      Width           =   1005
   End
   Begin Project1.Download Download1 
      Left            =   3600
      Top             =   2205
      _ExtentX        =   847
      _ExtentY        =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File/String"
      Height          =   330
      Left            =   1125
      TabIndex        =   5
      Top             =   2250
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Picture"
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   2250
      Width           =   960
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":0000
      Top             =   1800
      Width           =   4245
   End
   Begin VB.PictureBox picWeb 
      Height          =   1680
      Left            =   45
      ScaleHeight     =   1620
      ScaleWidth      =   4185
      TabIndex        =   2
      Top             =   45
      Width           =   4245
   End
   Begin VB.PictureBox picProgress 
      BackColor       =   &H00404040&
      Height          =   285
      Left            =   90
      ScaleHeight     =   225
      ScaleMode       =   0  'User
      ScaleWidth      =   131.429
      TabIndex        =   0
      Top             =   2700
      Width           =   4200
      Begin VB.Label lblProgress 
         BackColor       =   &H000000FF&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   195
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

    Download1.BeginDownload "http://www.google.com/images/logo.gif", "google pic", vbAsyncTypePicture

End Sub

Private Sub Command2_Click()

    Download1.BeginDownload "http://www.babbleon.net/testfile.txt", "testfile.txt"

End Sub

Private Sub Command3_Click()

    Download1.BeginDownload "http://www.somesitethatdontexist.com/file.txt", "(error test)"

End Sub

Private Sub Download1_DownloadComplete(ByVal Key As String, ByVal Value As Variant, ByVal FileType As AsyncTypeConstants)

    If Key = "google pic" Then
        
        'neato
        Set picWeb = Value

    ElseIf Key = "testfile.txt" Then
    
        'the default download type is a byte array,
        'which must be converted to string
        Text1 = StrConv(Value, vbUnicode)
    
    End If

    lblProgress.Width = 0

End Sub

Private Sub Download1_DownloadError(ByVal Key As String, ByVal Code As Long, ByVal Description As String)

    Download1.Cancel Key

    MsgBox Key & " Download Failed With Error: " & Description

End Sub

Private Sub Download1_DownloadProgress(ByVal Key As String, ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusNumber As AsyncStatusCodeConstants, ByVal StatusText As String)

    On Error Resume Next
        
        picProgress.ScaleWidth = ProgressMax
        lblProgress.Width = Progress

        Caption = StatusText

End Sub

Private Sub Form_Load()
    
    lblProgress.Width = 0

End Sub
