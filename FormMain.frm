VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "侠盗飞车3辅助"
   ClientHeight    =   6165
   ClientLeft      =   7065
   ClientTop       =   4335
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6540
   Begin VB.Timer Invoker 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   5640
   End
   Begin VB.CommandButton End 
      Caption         =   "退出(&E)"
      Height          =   615
      Left            =   360
      TabIndex        =   25
      Top             =   4920
      Width           =   5655
   End
   Begin VB.CommandButton Start 
      Caption         =   "开始(&S)"
      Height          =   615
      Left            =   360
      TabIndex        =   24
      Top             =   4080
      Width           =   5655
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   11
      Left            =   4200
      TabIndex        =   23
      Top             =   3360
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   10
      Left            =   4200
      TabIndex        =   21
      Top             =   2712
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   9
      Left            =   4200
      TabIndex        =   19
      Top             =   2064
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   8
      Left            =   4200
      TabIndex        =   17
      Top             =   1416
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   7
      Left            =   4200
      TabIndex        =   15
      Top             =   768
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   6
      Left            =   4200
      TabIndex        =   13
      Top             =   120
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   5
      Left            =   1080
      TabIndex        =   11
      Top             =   3360
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   1080
      TabIndex        =   9
      Text            =   "LEAVEMEALONE"
      Top             =   2712
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   1080
      TabIndex        =   7
      Text            =   "ASPIRINE"
      Top             =   2064
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   1080
      TabIndex        =   5
      Text            =   "NUTTERTOOLS"
      Top             =   1416
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Text            =   "PROFESSIONALTOOLS"
      Top             =   768
      Width           =   1800
   End
   Begin VB.TextBox TextCmd 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Text            =   "THUGSTOOLS"
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label LabelF 
      Caption         =   "F12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   3360
      TabIndex        =   22
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F11"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   10
      Left            =   3360
      TabIndex        =   20
      Top             =   2715
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   9
      Left            =   3360
      TabIndex        =   18
      Top             =   2070
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   3360
      TabIndex        =   16
      Top             =   1410
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   14
      Top             =   765
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   3360
      TabIndex        =   12
      Top             =   120
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   8
      Top             =   2715
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   2070
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1410
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   765
      Width           =   495
   End
   Begin VB.Label LabelF 
      Caption         =   "F1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_F1 = &H70
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F2 = &H71
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const KEYEVENTF_KEYUP = &H2


Private Sub End_Click()
End
End Sub

Private Sub Invoker_Timer()

For k = VK_F1 To VK_F12 Step 1
Dim b() As Byte
If GetAsyncKeyState(k) < 0 Then
    s = UCase(TextCmd(k - VK_F1).Text)
    b = StrConv(s, vbFromUnicode)
    l = Len(s)
    If "" <> s Then
        For i = 1 To l Step 1
        keybd_event b(i - 1), 0, 0, 0
        keybd_event b(i - 1), 0, KEYEVENTF_KEYUP, 0
        Next i
    End If
    
End If

Next k

End Sub

Private Sub Start_Click()
Invoker.Enabled = True
End Sub
