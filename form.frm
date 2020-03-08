VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "211王成之万历表"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   5880
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "输出结果"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "该天是星期："
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "格式：年.月.日（如2000.01.01）"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入日期："
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, Year, Month, Day, yeardiff, monthdiff, daydiff, d, d1, d2, d3, diff As Single
Dim result As String

Private Sub Command1_Click()
Year = Val(Mid(Text1.Text, 1, 4))
Month = Val(Mid(Text1.Text, 6, 2))
Day = Val(Mid(Text1.Text, 9, 2))

yeardiff = Abs(2000 - Year)
monthdiff = Month - 1
daydiff = Day - 1

If Year > 2000 Then d1 = yeardiff * 365 + (yeardiff - 1) \ 4 + 1 Else: d1 = yeardiff * 365 + yeardiff \ 4

If Month > 8 Then
    i = (Month - 1) \ 2
  ElseIf Month > 7 Then
    i = 3
  ElseIf Month > 3 Then
    i = Month \ 3
  Else
    i = 0
End If

If Month > 2 And yeardiff Mod 4 = 0 Then
  d2 = 60 + (monthdiff - 2) * 30 + i
 ElseIf Month > 2 And yeardiff Mod 4 <> 0 Then
  d2 = 59 + (monthdiff - 2) * 30 + i
 ElseIf Month > 1 Then
  d2 = 31
 Else: d2 = 0
End If

d3 = daydiff

If Year >= 2000 Then d = d1 + d2 + d3: diff = d Mod 7 + 1: result = Mid("六天一二三四五", diff, 1) Else: d = Abs(d1 - d2 - d3): diff = d Mod 7 + 1: result = Mid("六五四三二一天", diff, 1)
Text2.Text = result

End Sub

