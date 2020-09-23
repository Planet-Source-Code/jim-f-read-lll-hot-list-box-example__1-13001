VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "mamalukes list example"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "kill dupes"
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "load list"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "save list"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "copy selected"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "list 2 clipboard"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "remove selected"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "clear list"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ListBox lstOne 
      Height          =   1065
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "add"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "example by jim reed if u need help email baseballover9@aol.com ok later"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()

lstOne.AddItem Text1
Text1 = ""
End Sub



Private Sub Command1_Click()
On Error Resume Next

lstOne.Clear
End Sub



Private Sub Command2_Click()
On Error Resume Next
Clipboard.Clear

Clipboard.SetText lstOne.ListIndex
End Sub

Private Sub Command3_Click()
On Error Resume Next

lstOne.RemoveItem ListIndex
End Sub



Private Sub Command4_Click()

 
 Dim savelist As Long
    On Error Resume Next
    Open "C:\windows\savelist.dat" For Output As #1
    For savelist& = 0 To lstOne.ListCount - 1
        Print #1, lstOne.LisT(savelist&)
    Next savelist&
    Close #1
    Call MsgBox("list saved", vbOKOnly, LisT)

End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim SN As Long, TheList As String
For SN = 0 To lstOne.ListCount - 1
If SN = 0 Then
    TheList = lstOne.LisT(SN)
Else
    TheList = TheList & "," & lstOne.LisT(SN)
End If
Next
Clipboard.Clear

Clipboard.SetText TheList
End Sub

Private Sub Command6_Click()
Dim MyString As String
    On Error Resume Next

    Open "C:\windows\savelist.dat" For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
       If Len(MyString$) = 0 Or MyString$ = " " Then
       
       GoTo 10
       End If
       
        lstOne.AddItem MyString$
        lstOne.Refresh
10
    Wend
20
    Close #1
End Sub

Private Sub Command7_Click()
Dim Y
Dim i

For i = 0 To lstOne.ListCount - 1
Current = lstOne.LisT(i)
For Y = 0 To lstOne.ListCount - 1
Nower = lstOne.LisT(Y)

If UCase(Nower) = UCase(Current) Then
lstOne.RemoveItem Y

End If

dontkill:
Next Y
Next i
End Sub
