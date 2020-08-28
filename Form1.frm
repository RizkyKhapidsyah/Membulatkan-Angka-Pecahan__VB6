VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Membulatkan Angka Pecahan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function Round(nValue As Double, nDigits As Integer) As Double
    Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
End Function

Private Sub Command1_Click()
MsgBox Round(Val(Text1.Text), 1)
End Sub

Private Sub Form_Load()
  'Ganti '19.8455' dengan bilangan yang ingin Anda
  'bulatkan. Ganti '2' dengan jumlah digit setelah koma
  'untuk hasil setelah pembulatan.
  
End Sub


