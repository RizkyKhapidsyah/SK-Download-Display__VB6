VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmDownloadDisplay 
   Caption         =   "Download and Display HTML, RTF, or Text"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   Icon            =   "frmDownloadDisplay.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5085
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox URL 
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Text            =   "http://www.geocities.com/SiliconValley/Way/6445/main.html"
      Top             =   3360
      Width           =   4695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5760
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Display"
      Default         =   -1  'True
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5530
      _Version        =   327680
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmDownloadDisplay.frx":0442
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Written for the VB Center Code Library"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4140
      TabIndex        =   6
      Top             =   4560
      Width           =   2355
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "http://www.geocities.com/SiliconValley/Way/6445"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2160
      TabIndex        =   5
      Top             =   4800
      Width           =   4365
   End
   Begin VB.Label Label2 
      Caption         =   $"frmDownloadDisplay.frx":050B
      Height          =   400
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   6375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "URL:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3390
      Width           =   375
   End
End
Attribute VB_Name = "frmDownloadDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Written exclusively for VB Center by Marco Cordero.

Private Sub Command1_Click()
    
    Dim txt As String
    Dim b() As Byte
    
    On Error GoTo ErrorHandler
    
    
    Command1.Enabled = False
    
    ' This opens the file specified in the URL text box
    b() = Inet1.OpenURL(URL.Text, 1)
    
    txt = ""
    
    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    ' This loads the opened file into the RichTextBox control
    RichTextBox1.Text = txt
    
    Command1.Enabled = True
    
    Exit Sub
    
ErrorHandler:

    MsgBox "The document you requested could not be found.", vbCritical

    Exit Sub

End Sub
Private Sub Form_Load()

End Sub
