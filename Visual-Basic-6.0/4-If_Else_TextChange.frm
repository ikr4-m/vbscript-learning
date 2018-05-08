VERSION 5.00
Begin VB.Form databaseSederhana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Sederhana"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Database Siswa"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Jurusan"
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
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Tahun"
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
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Siswa"
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
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "NIS"
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
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "databaseSederhana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Mengosongkan semua textbox
Text1 = ""
Text2 = ""
Text3 = ""
Text4 = ""

'Biar app tidak galat/rusak, textbox yang terisi otomatis dimatikan
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

Private Sub Text1_Change()
'Untuk ngubah Text3 secara otomatis
If Left(Text1, 2) = "17" Then
    Text3 = "2017"
Else
    Text3 = ""
    End If
    
'Untuk ngubah Text4 secara otomatis
'ElseIf digunakan untuk mengurangi penutupan End If..End If yang terlalu banyak
If Mid(Text1, 3, 1) = "1" Then
    Text4 = "Teknik Komputer Jaringan"
ElseIf Mid(Text1, 3, 1) = "2" Then
    Text4 = "Rekayasa Perangkat Lunak"
Else
    Text4 = ""
    End If
    
'Untuk ngubah Text2 secara otomatis
'ElseIf digunakan untuk mengurangi penutupan End If..End If yang terlalu banyak
'Pakai "xxx-xxx" biar nanti kalau belakangnya kembar bakal gabisa tertabrak
'Kemudian, gapake Right,Left,Right function karena kepanjangan
If Text1 = "172-001" Then
    Text2 = "Andi Muh. Syahrul"
ElseIf Text1 = "172-002" Then
    Text2 = "Adrian Pratama"
ElseIf Text1 = "161-001" Then
    Text2 = "Andi Muh. Akbar"
ElseIf Text1 = "161-002" Then
    Text2 = "Abdullah Faqih Mustofa"
Else
    Text2 = ""
    End If
End Sub
