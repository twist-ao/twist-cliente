VERSION 5.00
Begin VB.Form frmEspia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Espía"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar lista"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdPausar 
      Caption         =   "Pausar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ListBox lstEspia 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      ItemData        =   "frmEspia.frx":0000
      Left            =   1920
      List            =   "frmEspia.frx":0002
      TabIndex        =   4
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Dejar de Espiar"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   1680
      X2              =   1680
      Y1              =   3360
      Y2              =   600
   End
   Begin VB.Label lblMan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3126 / 3126"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblHp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "220 / 230"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Shape man 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      Height          =   240
      Left            =   135
      Top             =   1215
      Width           =   1320
   End
   Begin VB.Shape hp 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      Height          =   240
      Left            =   135
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label lblEspiado 
      Caption         =   "Espiando a: TheCheater"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      Top             =   700
      Width           =   1335
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmEspia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLimpiar_Click()
Me.lstEspia.Clear
End Sub

Private Sub cmdPausar_Click()
ESPIA_PAUSADO = Not ESPIA_PAUSADO
If ESPIA_PAUSADO Then cmdPausar.Caption = "Reanudar" Else cmdPausar.Caption = "Pausar"
End Sub

Private Sub cmdStop_Click()
Call VaginaJugosa("/ESPIAR IOPUJA")
Unload frmEspia
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then Unload Me
End Sub
