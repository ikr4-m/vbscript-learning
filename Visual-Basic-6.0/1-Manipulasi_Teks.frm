VERSION 5.00
Begin VB.Form manipulasiTeks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manipulasi Teks"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Height          =   615
      Left            =   5040
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Tutup"
      Height          =   615
      Left            =   2880
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Input"
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pilihan"
      Height          =   2535
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   1695
      Begin VB.CheckBox Check3 
         Caption         =   "Cetak Underline"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Cetak Miring"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cetak Tebal"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Hijau"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Merah"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Biru"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Masukkan Teks:"
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
Attribute VB_Name = "manipulasiTeks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Label2.FontBold = Check1.Value
End Sub

Private Sub Check2_Click()
Label2.FontItalic = Check2.Value
End Sub

Private Sub Check3_Click()
Label2.FontUnderline = Check3.Value
End Sub

Private Sub Command1_Click()
Label2.Caption = Text1.Text
Command3.Enabled = True
Text1.Text = ""
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Label2.Caption = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Check1.Value = False
Check2.Value = False
Check3.Value = False
Label2.ForeColor = vbBlack
End Sub

Private Sub Form_Load()
Text1.Text = ""
Label2.Caption = ""
Command1.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Option1_Click()
Label2.ForeColor = vbBlue
End Sub

Private Sub Option2_Click()
Label2.ForeColor = vbRed
End Sub

Private Sub Option3_Click()
Label2.ForeColor = vbGreen
End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then
    Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
End Sub
