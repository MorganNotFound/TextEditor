VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "�ı��༭�� Copyright (c) 2022"
   ClientHeight    =   7680
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14070
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   14070
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13560
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   14055
   End
   Begin VB.Menu Menu1 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu File1 
         Caption         =   "��(&O)"
      End
      Begin VB.Menu File2 
         Caption         =   "����(&S)"
      End
      Begin VB.Menu File3 
         Caption         =   "���Ϊ(&A)"
      End
      Begin VB.Menu File4 
         Caption         =   "�رմ򿪵��ĵ�(&C)"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "�༭(&E)"
      Begin VB.Menu Edit1 
         Caption         =   "���(&C)"
      End
   End
   Begin VB.Menu Menu3 
      Caption         =   "��ʽ(&O)"
      Begin VB.Menu Format1 
         Caption         =   "�Զ�����(&W)"
         Checked         =   -1  'True
         Enabled         =   0   'False
      End
      Begin VB.Menu Format2 
         Caption         =   "����(&F)"
      End
      Begin VB.Menu Format3 
         Caption         =   "���뷽ʽ...(&A)"
         Begin VB.Menu Alleft 
            Caption         =   "�����(&L)"
         End
         Begin VB.Menu Alright 
            Caption         =   "�Ҷ���(&R)"
         End
         Begin VB.Menu Alcenter 
            Caption         =   "����(&C)"
         End
      End
   End
   Begin VB.Menu Menu4 
      Caption         =   "����(&H)"
      Begin VB.Menu Help1 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu Help2 
         Caption         =   "����(&W)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FormOldWidth As Long
Private FormOldHeight As Long
Public Sub ResizeInit(FormName As Form)
    Dim Obj As Control
    FormOldWidth = FormName.ScaleWidth
    FormOldHeight = FormName.ScaleHeight
    On Error Resume Next
    For Each Obj In FormName
    Obj.Tag = Obj.Left & " " & Obj.Top & " " & Obj.Width & " " & Obj.Height & " "
    Next Obj
    On Error GoTo 0
End Sub
Public Sub ResizeForm(FormName As Form)
    Dim Pos(4) As Double
    Dim i As Long, TempPos As Long, StartPos As Long
    Dim Obj As Control
    Dim ScaleX As Double, ScaleY As Double
    ScaleX = FormName.ScaleWidth / FormOldWidth
    ScaleY = FormName.ScaleHeight / FormOldHeight
    On Error Resume Next
    For Each Obj In FormName
    StartPos = 1
    For i = 0 To 4
        TempPos = InStr(StartPos, Obj.Tag, " ", vbTextCompare)
        If TempPos > 0 Then
        Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
        StartPos = TempPos + 1
        Else
        Pos(i) = 0
        End If
        Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
    Next i
    Next Obj
    On Error GoTo 0
End Sub
Private Sub Alcenter_Click()
Text1.Alignment = 2
End Sub
Private Sub Alleft_Click()
Text1.Alignment = 0
End Sub
Private Sub Alright_Click()
Text1.Alignment = 1
End Sub
Private Sub Edit1_Click()
Text1.Text = ""
End Sub
Private Sub File1_Click()
On Error GoTo MyErr
Dim temp As String
Dim all As String
Dim file As String
    With CommonDialog1
        .Filter = "TxtFiles(*.txt)|*.txt|All Files(*.*)|*.*|"
        .ShowOpen
    End With
    file = CommonDialog1.FileName
Open CommonDialog1.FileName For Input As #1
Do While Not EOF(1)
Input #1, temp
all = all & temp & Chr(13) & Chr(10)
Loop
Close #1
Text1.Text = all
MyErr:
    On Error GoTo 0
End Sub
Private Sub File2_Click()
On Error GoTo MyErr
Dim temp As String
Dim all As String
Dim file As String
Open CommonDialog1.FileName For Output As #1
Do While Not EOF(1)
Input #1, temp
all = all & temp & Chr(13) & Chr(10)
Loop
Print Text1.Text
Close #1
Open CommonDialog1.FileName For Output As #1
Print #1, Text1.Text
Close #1
Exit Sub
MyErr:
    MsgBox "δ���κ��ĵ����뵥�����Ϊ��ť", vbOKOnly, "����"
End Sub
Private Sub File3_Click()
On Error GoTo MyErr
Dim temp As String
Dim all As String
Dim file As String
    With CommonDialog1
        .Filter = "Txt�ı��ĵ�(*.txt)|*.txt|Markdown Files(*.md)|*.md|PythonFiles(*.py;*.pyw)|*.py|Cmd(*.cmd)|*.cmd|Bat�������ļ�(*.bat)|*.bat|VBScripts(*.VBS)|*.VBS|All Files(*.*)|*.*|"
        .ShowSave
    End With
Open CommonDialog1.FileName For Output As #1
Do While Not EOF(1)
Input #1, temp
all = all & temp & Chr(13) & Chr(10)
Loop
Print Text1.Text
Close #1
Open CommonDialog1.FileName For Output As #1
Print #1, Text1.Text
Close #1
MyErr:
    On Error GoTo 0
End Sub
Private Sub File4_Click()
CommonDialog1.FileName = ""
Text1.Text = ""
End Sub
Private Sub form_load()
Call ResizeInit(Me)
Text1.ForeColor = vbBlack
File4.Enabled = False
End Sub
Private Sub Form_Resize()
Call ResizeForm(Me)
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Format2_Click()
Form2.Show
End Sub
Private Sub Help1_Click()
frmAbout.Show
End Sub
Private Sub Help2_Click()
Shell "explorer https://github.com/MorganNotFound/"
End Sub
Private Sub Text1_Change()
If CommonDialog1.FileName = "" Then
File4.Enabled = False
Else
File4.Enabled = True
End If
End Sub
