VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Hitung"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "School Project (Im not Happy about it).frx":0000
      Left            =   1080
      List            =   "School Project (Im not Happy about it).frx":0010
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Hasil 
      Caption         =   "Hasil"
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Luas ="
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Nama2 
      Caption         =   "Lebar ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Nama1 
      Caption         =   "Panjang ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "PILIH YANG ANDA INGINKAN!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
X = Combo1.Text
Select Case X
Case "Persegi Panjang":
Nama1.Caption = "Panjang ="
Nama2.Caption = "Lebar ="
Nama3.Caption = "Luas ="
Hasil.Caption = "..."
Case "Segitiga":
Name1.Caption = "Alas ="
Nama2.Caption = "Tinggi ="
Nama3.Caption = "Luas ="
Hasil.Caption = "..."
Case "Tabung":
Nama1.Caption = "Jari-Jari Alas ="
Nama2.Caption = "Tinggi ="
Nama3.Caption = "Volume ="
Hasil.Caption = "..."
End Select
End Sub

Private Sub Command1_Click()
X = Combo1.Text
Select Case X
Case "Persegi Panjang":
LV = Text1.Text * Text2.Text
Hasil.Caption = "..."
Case "Segitiga":
LV = Text1.Text * Text2.Text / 2
Case "Tabung":
LV = 3.14 * Text1.Text ^ 2 * Text2.Text
Case "Kerucut"
LV = 3.14 * Text1.Text ^ 2 * Text2.Text / 3
End Select
Hasil.Caption = LV
End Sub

Private Sub Label3_Click()

End Sub

