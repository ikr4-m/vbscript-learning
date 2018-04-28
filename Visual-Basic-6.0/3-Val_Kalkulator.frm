VERSION 5.00
Begin VB.Form kalkulator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kalkulator Sederhana"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4950
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4950
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option12 
      Caption         =   "Mod"
      Height          =   255
      Left            =   2880
      TabIndex        =   19
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton Option11 
      Caption         =   "^"
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   3360
      Width           =   855
   End
   Begin VB.OptionButton Option10 
      Caption         =   "/"
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton Option9 
      Caption         =   "*"
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.OptionButton Option8 
      Caption         =   "-"
      Height          =   255
      Left            =   2880
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.OptionButton Option7 
      Caption         =   "+"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   1920
      Width           =   495
   End
   Begin VB.OptionButton Option6 
      Caption         =   "<="
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   855
   End
   Begin VB.OptionButton Option5 
      Caption         =   ">="
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   3360
      Width           =   1575
   End
   Begin VB.OptionButton Option4 
      Caption         =   "="
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   3000
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "<>"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "<"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   ">"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operasi Aritmatika"
      Height          =   2655
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operasi Logika"
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Hasil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Angka Kedua"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Angka Pertama"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "kalkulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Text3.Enabled = False
Text1 = ""
Text2 = ""
Text3 = ""
End Sub

Private Sub Option1_Click()
Text3 = Val(Text1) > Val(Text2)
End Sub

Private Sub Option10_Click()
Text3 = Val(Text1) / Val(Text2)
End Sub

Private Sub Option11_Click()
Text3 = Val(Text1) ^ Val(Text2)
End Sub

Private Sub Option12_Click()
Text3 = Val(Text1) Mod Val(Text2)
End Sub

Private Sub Option2_Click()
Text3 = Val(Text1) < Val(Text2)
End Sub

Private Sub Option3_Click()
Text3 = Val(Text1) <> Val(Text2)
End Sub

Private Sub Option4_Click()
Text3 = Val(Text1) = Val(Text2)
End Sub

Private Sub Option5_Click()
Text3 = Val(Text1) >= Val(Text2)
End Sub

Private Sub Option6_Click()
Text3 = Val(Text1) <= Val(Text2)
End Sub

Private Sub Option7_Click()
Text3 = Val(Text1) + Val(Text2)
End Sub

Private Sub Option8_Click()
Text3 = Val(Text1) - Val(Text2)
End Sub

Private Sub Option9_Click()
Text3 = Val(Text1) * Val(Text2)
End Sub
