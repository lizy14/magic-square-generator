VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "幻方计算工具 (C)Li_Zaodie"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "导出到Excel表格"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "计算另一个幻方"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "完成"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   8415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(C)Li_Zaodie,2009.12"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   6360
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim HuanFang() As Long
Dim lieshu1 As Long

Sub yyj(lieshu As Long)
ReDim HuanFang(lieshu, lieshu)
Dim i As Long
Dim LastHang As Long
Dim LastLie As Long
Dim LastHang1 As Long
Dim LastLie1 As Long


LastLie = (lieshu + 1) / 2
LastHang = 1
Form2.Show
For i = 1 To lieshu * lieshu
Form2.Label3.Caption = Format(i / (lieshu * lieshu), "0.0000%") & " (" & i & " of " & lieshu * lieshu & ")"
DoEvents


HuanFang(LastHang, LastLie) = i

LastHang1 = LastHang
LastLie1 = LastLie


LastHang = LastHang - 1
If LastHang < 1 Then LastHang = lieshu
LastLie = LastLie + 1
If LastLie > lieshu Then LastLie = 1

If HuanFang(LastHang, LastLie) <> 0 Then
LastHang = LastHang1 + 1
LastLie = LastLie1
End If
Next i
Unload Form2

Dim Changdu
Changdu = lieshu * lieshu
Changdu = Str(Changdu)
Changdu = Len(Changdu)

Exit Sub
Form2.Show

Dim i1, i2, temp As String
Text1.Text = ""

For i2 = 1 To lieshu
For i1 = 1 To lieshu
temp = temp & Space(Changdu - Len(CStr(HuanFang(i2, i1)))) & HuanFang(i2, i1)

Form2.Label3.Caption = Format((((i2 - 1) * lieshu + i1) / (lieshu * lieshu)), "0.0000%") & " (" & (i2 - 1) * lieshu + i1 & " of " & lieshu * lieshu & ")"
Form2.Label1.Caption = "正在输出幻方,请稍候."
DoEvents

Next
temp = temp & nline
Next
Beep
Unload Form2
Text1.Text = temp


End Sub

Function nline()
nline = Chr(13) + Chr(10)
End Function

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Unload Me
Me.Show
End Sub
Function fixPath(yy As String) As String
If Right(yy, 1) <> "\" Then yy = yy & "\"
fixPath = yy
End Function
Private Sub Command3_Click()
On Error GoTo haNdEr
Form2.Label1.Caption = "正在导出幻方,请稍候."
Form2.Show
Dim opa As String
opa = fixPath(App.Path) & "Huanfang.csv"

Dim i1, i2, temp As String
Open opa For Output As #15
Print #15,
Print #15, ",," & lieshu1 & "*" & lieshu1 & "幻方"
Print #15,
For i2 = 1 To lieshu1
For i1 = 1 To lieshu1
Print #15, "," & HuanFang(i2, i1),

Form2.Label3.Caption = Format((((i2 - 1) * lieshu1 + i1) / (lieshu1 * lieshu1)), "0.0000%") & " (" & (i2 - 1) * lieshu1 + i1 & " of " & lieshu1 * lieshu1 & ")"

DoEvents

Next
Print #15,
Next
Beep
Unload Form2
Close #15
Exit Sub
haNdEr:
MsgBox "保存过程中出现错误,未能成功导出。" & nline & nline & "出错信息: " & Err & " " & Error(Err), vbCritical
Unload Form2
Form1.Show
End Sub

Private Sub Form_Load()
Dim bei
start1:
bei = (InputBox("请输入一个奇数。", Form1.Caption))
If bei = "" Then End


If Val(bei) Mod 2 <> 1 Then
MsgBox "“" & bei & "”也能算奇数？！" & nline & nline & "请重新输入。", vbExclamation
GoTo start1
End If

yyj CLng(bei)
'Form1.Caption = "幻方计算工具 (C)Li_Zaodie" & " - " & bei & "*" & bei
lieshu1 = CLng(bei)
Command3_Click

ShellExecute Me.hwnd, vbNullString, fixPath(App.Path) & "Huanfang.csv", vbNullString, vbNullString, SW_SHOWNORMAL

End
End Sub

