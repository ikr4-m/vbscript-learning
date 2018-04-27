VERSION 5.00
Begin VB.Form listCombo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penggunaan Combo & List Box"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4470
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Nama Guru"
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "listCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.AddItem Combo1.Text
Combo1.SetFocus
End Sub

Private Sub Command2_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
List1.Clear
Combo1.SetFocus
End Sub

Private Sub Form_Load()
Combo1.AddItem "Irman Aras"
Combo1.AddItem "Andy Rafi"
Combo1.AddItem "Achmad"
Combo1.AddItem "Marlina"
Combo1.AddItem "Sumarni"
End Sub
