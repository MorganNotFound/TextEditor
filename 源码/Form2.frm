VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "字体"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "颜色"
      Height          =   4455
      Left            =   4200
      TabIndex        =   21
      Top             =   120
      Width           =   2655
      Begin VB.HScrollBar HSB 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   34
         Top             =   3840
         Width           =   1455
      End
      Begin VB.HScrollBar HSG 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   33
         Top             =   3360
         Width           =   1455
      End
      Begin VB.HScrollBar HSR 
         Height          =   255
         Left            =   840
         Max             =   255
         TabIndex        =   32
         Top             =   2880
         Width           =   1455
      End
      Begin VB.OptionButton Option10 
         Caption         =   "自定义颜色..."
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   2280
         Width           =   1575
      End
      Begin VB.OptionButton Option9 
         Caption         =   "黑色(&B)"
         Height          =   375
         Left            =   1440
         TabIndex        =   29
         Top             =   1680
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option8 
         Caption         =   "白色(&W)"
         Height          =   375
         Left            =   1440
         TabIndex        =   28
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option7 
         Caption         =   "洋红(&M)"
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option6 
         Caption         =   "青色(&C)"
         Height          =   375
         Left            =   1440
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "黄(&Y)"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "绿(&G)"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option3 
         Caption         =   "蓝(&B)"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "红(&R)"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "蓝(B)"
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "绿(G)"
         Height          =   375
         Left            =   360
         TabIndex        =   36
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "红(R)"
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label LabelColor 
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   2280
         Width           =   255
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1215
      Left            =   3720
      Max             =   1
      Min             =   1000
      TabIndex        =   18
      Top             =   3120
      Value           =   1
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   2160
      TabIndex        =   17
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "删除线(&K)"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "下划线(&U)"
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "字形"
      Height          =   1455
      Left            =   2160
      TabIndex        =   12
      Top             =   120
      Width           =   1815
      Begin VB.CheckBox Check2 
         Caption         =   "斜体(&I)"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "粗体(&B)"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "字体"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Option1 
         Caption         =   "Georgia"
         Height          =   375
         Index           =   10
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Terminal"
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   10
         Top             =   3480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Broadway"
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   9
         Top             =   3120
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cambria"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "幼圆"
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "等线"
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "仿宋"
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "隶书"
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "楷体"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "黑体"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "宋体"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "字号"
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "妖王Azis"
      Height          =   1095
      Left            =   2160
      TabIndex        =   19
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2
Private Sub Check1_Click()
If Check1.Value = 0 Then
Form1.Text1.FontBold = False
End If
If Check1.Value = 1 Then
Form1.Text1.FontBold = True
End If
End Sub
Private Sub Check2_Click()
If Check2.Value = 0 Then
Form1.Text1.FontItalic = False
End If
If Check2.Value = 1 Then
Form1.Text1.FontItalic = True
End If
End Sub
Private Sub Check3_Click()
If Check3.Value = 0 Then
Form1.Text1.FontUnderline = False
End If
If Check3.Value = 1 Then
Form1.Text1.FontUnderline = True
End If
End Sub
Private Sub Check4_click()
If Check4.Value = 0 Then
Form1.Text1.FontStrikethru = False
End If
If Check4.Value = 1 Then
Form1.Text1.FontStrikethru = True
End If
End Sub
Private Sub HSR_Change()
LabelColor.BackColor = RGB(HSR.Value, HSG.Value, HSB.Value)
Form1.Text1.ForeColor = LabelColor.BackColor
End Sub
Private Sub HSG_Change()
LabelColor.BackColor = RGB(HSR.Value, HSG.Value, HSB.Value)
Form1.Text1.ForeColor = LabelColor.BackColor
End Sub
Private Sub HSB_Change()
LabelColor.BackColor = RGB(HSR.Value, HSG.Value, HSB.Value)
Form1.Text1.ForeColor = LabelColor.BackColor
End Sub
Private Sub Option1_Click(Index As Integer)
Form1.Text1.FontName = Option1(Index).Caption
Label1.FontName = Option1(Index).Caption
End Sub
Private Sub Option2_Click()
Form1.Text1.ForeColor = vbRed
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option3_Click()
Form1.Text1.ForeColor = vbBlue
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option4_Click()
Form1.Text1.ForeColor = vbGreen
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option5_Click()
Form1.Text1.ForeColor = vbYellow
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option6_Click()
Form1.Text1.ForeColor = vbCyan
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option7_Click()
Form1.Text1.ForeColor = vbMagenta
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option8_Click()
Form1.Text1.ForeColor = vbWhite
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option9_Click()
Form1.Text1.ForeColor = vbBlack
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End Sub
Private Sub Option10_Click()
HSR.Enabled = True
HSG.Enabled = True
HSB.Enabled = True
LabelColor.BackColor = RGB(HSR.Value, HSG.Value, HSB.Value)
Form1.Text1.ForeColor = Label1.BackColor
End Sub
Private Sub Text2_Change()
If Text2.Text = "" Then
Text2.Text = 1
End If
If Text2.Text < 1 Then
Text2.Text = 1
End If
If Text2.Text > 1000 Then
Text2.Text = 1000
End If
VScroll1.Value = Text2.Text
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub
Private Sub VScroll1_Change()
Label1.FontSize = VScroll1.Value
Text2.Text = VScroll1.Value
Form1.Text1.FontSize = VScroll1.Value
End Sub
Private Sub form_load()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
VScroll1.Value = Int(Form1.Text1.FontSize)
Option1(0).FontName = Option1(0).Caption
Option1(1).FontName = Option1(1).Caption
Option1(2).FontName = Option1(2).Caption
Option1(3).FontName = Option1(3).Caption
Option1(4).FontName = Option1(4).Caption
Option1(5).FontName = Option1(5).Caption
Option1(6).FontName = Option1(6).Caption
Option1(7).FontName = Option1(7).Caption
Option1(8).FontName = Option1(8).Caption
Option1(9).FontName = Option1(9).Caption
Option1(10).FontName = Option1(10).Caption
If Option1(1).Caption = Form1.Text1.FontName Then
Option1(1).Value = True
End If
If Option1(2).Caption = Form1.Text1.FontName Then
Option1(2).Value = True
End If
If Option1(3).Caption = Form1.Text1.FontName Then
Option1(3).Value = True
End If
If Option1(4).Caption = Form1.Text1.FontName Then
Option1(4).Value = True
End If
If Option1(5).Caption = Form1.Text1.FontName Then
Option1(5).Value = True
End If
If Option1(6).Caption = Form1.Text1.FontName Then
Option1(6).Value = True
End If
If Option1(7).Caption = Form1.Text1.FontName Then
Option1(7).Value = True
End If
If Option1(8).Caption = Form1.Text1.FontName Then
Option1(8).Value = True
End If
If Option1(9).Caption = Form1.Text1.FontName Then
Option1(9).Value = True
End If
If Option1(10).Caption = Form1.Text1.FontName Then
Option1(10).Value = True
End If
If Form1.Text1.FontBold = True Then
Let Check1.Value = 1
End If
If Form1.Text1.FontItalic = True Then
Let Check2.Value = 1
End If
If Form1.Text1.FontUnderline = True Then
Let Check3.Value = 1
End If
If Form1.Text1.FontStrikethru = True Then
Let Check4.Value = 1
End If
If Form1.Text1.ForeColor = vbRed Then
Option2.Value = True
ElseIf Form1.Text1.ForeColor = vbBlue Then
Option3.Value = True
ElseIf Form1.Text1.ForeColor = vbGreen Then
Option4.Value = True
ElseIf Form1.Text1.ForeColor = vbYellow Then
Option5.Value = True
ElseIf Form1.Text1.ForeColor = vbCyan Then
Option6.Value = True
ElseIf Form1.Text1.ForeColor = vbMagenta Then
Option7.Value = True
ElseIf Form1.Text1.ForeColor = vbWhite Then
Option8.Value = True
ElseIf Form1.Text1.ForeColor = vbBlack Then
Option9.Value = True
Else
Option10.Value = True
End If
If Option10.Value = False Then
HSR.Enabled = False
HSG.Enabled = False
HSB.Enabled = False
End If
End Sub
