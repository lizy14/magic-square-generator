VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "幻方计算工具 (C)Li_Zaodie"
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "进度:"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "正在计算幻方,请稍候。"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1890
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

