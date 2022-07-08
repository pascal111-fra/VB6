VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Percentage%"
   ClientHeight    =   3510
   ClientLeft      =   8505
   ClientTop       =   4755
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7020
   Begin VB.CommandButton Command2 
      Caption         =   "much"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "%"
      Height          =   615
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtresult 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtof 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "of"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim x_of As Double, total As Double, result As Double

If IsNumeric(txttotal) Then
total = Val(txttotal)
Else
txtresult = "Error"
Exit Sub
End If

If IsNumeric(txtof) Then
x_of = Val(txtof)
Else
txtresult = "Error"
Exit Sub
End If


result = (x_of * 100) / total

txtresult = Str(result)

End Sub

Private Sub Command2_Click()

Dim x_of As Double, total As Double, result As Double

If IsNumeric(txttotal) Then
total = Val(txttotal)
Else
txtof = "Error"
Exit Sub
End If

If IsNumeric(txtresult) Then
result = Val(txtresult)
Else
txtof = "Error"
Exit Sub
End If


x_of = (result / 100) * total

txtof = Str(x_of)

End Sub
