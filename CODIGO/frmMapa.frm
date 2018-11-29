VERSION 5.00
Begin VB.Form FrmMapa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5775
   ClientLeft      =   3930
   ClientTop       =   1155
   ClientWidth     =   7275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMapa.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   7275
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   7335
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   7560
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.Path & "\Graficos\Mapa.jpg")
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub


