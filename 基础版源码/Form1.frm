VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "简易文本编辑器"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8190
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame4 
      Caption         =   "字体位置"
      Height          =   2775
      Left            =   6240
      TabIndex        =   4
      Top             =   3240
      Width           =   1815
      Begin VB.OptionButton Option10 
         Caption         =   "居中"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   1455
      End
      Begin VB.OptionButton Option9 
         Caption         =   "右对齐"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option8 
         Caption         =   "左对齐"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "字体"
      Height          =   2775
      Left            =   4200
      TabIndex        =   3
      Top             =   3240
      Width           =   1815
      Begin VB.VScrollBar VScroll1 
         Height          =   495
         Left            =   1200
         Max             =   1
         Min             =   1000
         TabIndex        =   17
         Top             =   1920
         Value           =   1
         Width           =   495
      End
      Begin VB.OptionButton Option7 
         Caption         =   "宋体"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "黑体"
         BeginProperty Font 
            Name            =   "黑体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "字号"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "字体颜色"
      Height          =   2775
      Left            =   2160
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
      Begin VB.OptionButton Option5 
         Caption         =   "黑色(&B)"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "绿色(&G)"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "蓝色(&B)"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "黄色(&Y)"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "红色(&R)"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "特殊字体"
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   1815
      Begin VB.CheckBox Check4 
         Caption         =   "删除线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "下划线"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "加粗"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "斜体"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private FormOldWidth As Long
'Private FormOldHeight As Long
'Public Sub ResizeInit(FormName As Form)
'    Dim Obj As Control
'    FormOldWidth = FormName.ScaleWidth
'    FormOldHeight = FormName.ScaleHeight
'    On Error Resume Next
'    For Each Obj In FormName
'    Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
'    Next Obj
'    On Error GoTo 0
'End Sub
'Public Sub ResizeForm(FormName As Form)
'    Dim Pos(4) As Double
'    Dim i As Long, TempPos As Long, StartPos As Long
'    Dim Obj As Control
'    Dim ScaleX As Double, ScaleY As Double
'    ScaleX = FormName.ScaleWidth / FormOldWidth
'    ScaleY = FormName.ScaleHeight / FormOldHeight
'    On Error Resume Next
'    For Each Obj In FormName
'    StartPos = 1
'    For i = 0 To 4
'        TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
'        If TempPos > 0 Then
'        Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
'        StartPos = TempPos + 1
'        Else
'        Pos(i) = 0
'        End If
'        Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
'    Next i
'    Next Obj
'    On Error GoTo 0
'End Sub
'Private Sub Form_Resize()
'Call ResizeForm(Me)
'End Sub
Private Sub Check1_Click()
Text1.FontItalic = Not Text1.FontItalic
End Sub
Private Sub Check2_Click()
Text1.FontBold = Not Text1.FontBold
End Sub
Private Sub Check3_Click()
Text1.FontUnderline = Not Text1.FontUnderline
End Sub
Private Sub Check4_Click()
Text1.FontStrikethru = Not Text1.FontStrikethru
End Sub
Private Sub Form_Load()
'Call ResizeInit(Me)
VScroll1.Value = Int(Text1.FontSize)
End Sub
Private Sub Option1_Click()
Text1.ForeColor = vbRed
End Sub
Private Sub Option2_Click()
Text1.ForeColor = vbYellow
End Sub
Private Sub Option3_Click()
Text1.ForeColor = vbBlue
End Sub
Private Sub Option4_Click()
Text1.ForeColor = vbGreen
End Sub
Private Sub Option5_Click()
Text1.ForeColor = vbBlack
End Sub
Private Sub Option6_Click()
Text1.FontName = Option6.Caption
End Sub
Private Sub Option7_Click()
Text1.FontName = Option7.Caption
End Sub
Private Sub Option8_Click()
Text1.Alignment = 0
End Sub
Private Sub Option9_Click()
Text1.Alignment = 1
End Sub
Private Sub Option10_Click()
Text1.Alignment = 2
End Sub
Private Sub VScroll1_Change()
Text1.FontSize = VScroll1.Value
Label2.Caption = VScroll1.Value
End Sub
