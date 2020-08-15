VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menekan beberapa Tombol Keyboard  Bersamaan"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coding berikut ini menunjukkan kepada kita tombol yang 'sedang ditekan, Atas, Bawah, Kiri, Kanan, dan Spasi. 'Anda bisa saja mengganti dengan pengecekan pada tombol 'keyboard lainnya.

Dim currentKeys(0 To 250) As Boolean 'Deklarasi 'variabel global

Private Sub Form_KeyDown(KeyCode As Integer, Shift As _
Integer)
'Menekan sebuah tombol dan menahannya, akan menggunakan 'event KeyDown.
'Jadi kita memeriksa jika User menekan tombol atau
'apakah tombol sudah ditekan.
    If currentKeys(KeyCode) = False Then
    'Update array dari tombol yang ditekan
        currentKeys(KeyCode) = True
        If KeyCode = vbKeyLeft Then Label1 = "Kiri"
        If KeyCode = vbKeyRight Then Label2 = "Kanan"
        If KeyCode = vbKeyUp Then Label3 = "Atas"
        If KeyCode = vbKeyDown Then Label4 = "Bawah"
        If KeyCode = vbKeySpace Then Label5 = "Spasi"
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Update array yang tombol keyboardnya tidak ditekan.
    currentKeys(KeyCode) = False
    If KeyCode = vbKeyLeft Then Label1 = ""
    If KeyCode = vbKeyRight Then Label2 = ""
    If KeyCode = vbKeyUp Then Label3 = ""
    If KeyCode = vbKeyDown Then Label4 = ""
    If KeyCode = vbKeySpace Then Label5 = ""
End Sub


