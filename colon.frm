VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6396
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   10152
   LinkTopic       =   "Form1"
   ScaleHeight     =   6396
   ScaleWidth      =   10152
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   5268
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   6852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   612
      Left            =   7800
      TabIndex        =   0
      Top             =   720
      Width           =   2172
   End
   Begin VB.Label Label1 
      Caption         =   "Gene seq    t-value"
      Height          =   180
      Left            =   720
      TabIndex        =   2
      Top             =   360
      Width           =   3732
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Open App.Path & "\colon.txt" For Input As #1

Dim fileArray(2002, 62) As Variant
s = 0
Dim oneclass(62) As Integer
Dim twoclass(62) As Integer

'讀檔並存入二維陣列
Do While Not EOF(1)
    Line Input #1, tmpline
    b = Split(tmpline, ",")
    For i = 0 To 62
        fileArray(s, i) = b(i)
    Next i
    s = s + 1 '用s紀錄2000筆，用i記錄每筆的62個資料
Loop

Dim totalOne As Double
Dim XOne As Double
Dim totalTwo As Double
Dim XTwo As Double
Dim squOne As Double
Dim squTwo As Double
Dim SOne As Double
Dim STwo As Double
Dim tvalue As Double
Dim tArray(2002) As Double
Dim tArrayIndex(2002) As Integer
'計算t值
For k = 2 To 2001
    For i = 1 To 62
        If fileArray(1, i) = "1" Then
            totalOne = totalOne + CDbl(fileArray(k, i))
            XOne = totalOne / 22
        Else
            totalTwo = totalTwo + CDbl(fileArray(k, i))
            XTwo = totalTwo / 40
        End If
    Next i
    
    For i = 1 To 62
        If fileArray(1, i) = "1" Then
            squOne = squOne + CDbl((fileArray(k, i) - XOne) ^ 2)
        Else
            squTwo = squTwo + CDbl((fileArray(k, i) - XOne) ^ 2)
        End If
    Next i
    SOne = squOne / (22 - 1)
    STwo = squTwo / (40 - 1)
    tvalue = (XOne - XTwo) / Sqr((SOne ^ 2) / 22 + (STwo ^ 2) / 40)
    
    tArray(k) = tvalue
    tArrayIndex(k) = k - 1
    '把計算值全部歸零
    totalOne = 0
    XOne = 0
    totalTwo = 0
    XTwo = 0
    squOne = 0
    squTwo = 0
    SOne = 0
    STwo = 0
    tvalue = 0
Next k


For i = 2 To 2001
    For j = 2 To 2001
        If tArray(j) > tArray(j + 1) Then '由小到大排序
            x = tArray(j)
            y = tArrayIndex(j)
            tArray(j) = tArray(j + 1)
            tArrayIndex(j) = tArrayIndex(j + 1)
            tArray(j + 1) = x
            tArrayIndex(j + 1) = y
        End If
    Next j
Next i

For m = 2 To 2002
    List1.AddItem (tArrayIndex(m) & vbTab & tArray(m))
Next m

Close #1
End Sub

