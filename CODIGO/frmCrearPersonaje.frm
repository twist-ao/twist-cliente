VERSION 5.00
Begin VB.Form frmCrearPersonaje1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox headPbx 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   240
      Left            =   1620
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   48
      Top             =   6840
      Width           =   1035
   End
   Begin VB.TextBox txtCorreoCheck 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   8550
      TabIndex        =   6
      Top             =   3150
      Width           =   3135
   End
   Begin VB.TextBox txtPasswdCheck 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   8520
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox txtPasswd 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   8550
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   4650
      Width           =   3135
   End
   Begin VB.TextBox txtCorreo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   8550
      TabIndex        =   5
      Top             =   2370
      Width           =   3135
   End
   Begin VB.TextBox txtPreg 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   8760
      TabIndex        =   9
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtResp 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   8880
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0000
      Left            =   1365
      List            =   "frmCrearPersonaje.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5535
      Width           =   1500
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0063
      Left            =   1365
      List            =   "frmCrearPersonaje.frx":006D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6375
      Width           =   1500
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":0080
      Left            =   1365
      List            =   "frmCrearPersonaje.frx":0093
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5955
      Width           =   1500
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":00C0
      Left            =   1380
      List            =   "frmCrearPersonaje.frx":00CA
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   7215
      Width           =   1500
   End
   Begin VB.TextBox txtNombre 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8550
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1605
      Width           =   3105
   End
   Begin VB.Label lblMaxAgi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   3180
      TabIndex        =   47
      Top             =   1920
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblMaxFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3180
      TabIndex        =   46
      Top             =   1560
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image boton 
      Height          =   1650
      Index           =   2
      Left            =   600
      MouseIcon       =   "frmCrearPersonaje.frx":00DF
      MousePointer    =   99  'Custom
      Top             =   3480
      Width           =   1680
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   45
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   44
      Top             =   2040
      Width           =   330
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   43
      Top             =   3030
      Width           =   330
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   42
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   41
      Top             =   2745
      Width           =   330
   End
   Begin VB.Label lblCarisma2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   40
      Top             =   2745
      Width           =   330
   End
   Begin VB.Label lblInteligencia2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   39
      Top             =   2400
      Width           =   330
   End
   Begin VB.Label lblConstitucion2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   38
      Top             =   3030
      Width           =   330
   End
   Begin VB.Label lblAgilidad2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   37
      Top             =   2040
      Width           =   330
   End
   Begin VB.Label lblFuerza2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   36
      Top             =   1680
      Width           =   330
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6360
      TabIndex        =   35
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   6330
      TabIndex        =   34
      Top             =   6270
      Width           =   390
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   41
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":0231
      MousePointer    =   99  'Custom
      Top             =   6330
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   40
      Left            =   6840
      MouseIcon       =   "frmCrearPersonaje.frx":0383
      MousePointer    =   99  'Custom
      Top             =   6330
      Width           =   135
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   6330
      TabIndex        =   33
      Top             =   6480
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   6330
      TabIndex        =   32
      Top             =   6705
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   23
      Left            =   6330
      TabIndex        =   31
      Top             =   6930
      Width           =   390
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   42
      Left            =   6840
      MouseIcon       =   "frmCrearPersonaje.frx":04D5
      MousePointer    =   99  'Custom
      Top             =   6510
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   43
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":0627
      MousePointer    =   99  'Custom
      Top             =   6510
      Width           =   375
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   44
      Left            =   6840
      MouseIcon       =   "frmCrearPersonaje.frx":0779
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   45
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":08CB
      MousePointer    =   99  'Custom
      Top             =   6705
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   46
      Left            =   6840
      MouseIcon       =   "frmCrearPersonaje.frx":0A1D
      MousePointer    =   99  'Custom
      Top             =   6930
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   47
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":0B6F
      MousePointer    =   99  'Custom
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6330
      TabIndex        =   30
      Top             =   2400
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6330
      TabIndex        =   29
      Top             =   1980
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6330
      TabIndex        =   28
      Top             =   2205
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   6330
      TabIndex        =   27
      Top             =   2595
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   6330
      TabIndex        =   26
      Top             =   2805
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   6330
      TabIndex        =   25
      Top             =   3015
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   6330
      TabIndex        =   24
      Top             =   3225
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   6330
      TabIndex        =   23
      Top             =   3435
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   6330
      TabIndex        =   22
      Top             =   3675
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   6330
      TabIndex        =   21
      Top             =   3900
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   6330
      TabIndex        =   20
      Top             =   4110
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   6330
      TabIndex        =   19
      Top             =   4320
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   6330
      TabIndex        =   18
      Top             =   4545
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   6330
      TabIndex        =   17
      Top             =   4755
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   6330
      TabIndex        =   16
      Top             =   4980
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   6330
      TabIndex        =   15
      Top             =   5205
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   6330
      TabIndex        =   14
      Top             =   5415
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   6330
      TabIndex        =   13
      Top             =   5610
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   6330
      TabIndex        =   12
      Top             =   5820
      Width           =   390
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   6330
      TabIndex        =   11
      Top             =   6060
      Width           =   390
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   39
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":0CC1
      MousePointer    =   99  'Custom
      Top             =   6105
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   38
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":0E13
      MousePointer    =   99  'Custom
      Top             =   6105
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   37
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":0F65
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   35
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":10B7
      MousePointer    =   99  'Custom
      Top             =   5700
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":1209
      MousePointer    =   99  'Custom
      Top             =   5700
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":135B
      MousePointer    =   99  'Custom
      Top             =   5475
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":14AD
      MousePointer    =   99  'Custom
      Top             =   5475
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   31
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":15FF
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":1751
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   29
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":18A3
      MousePointer    =   99  'Custom
      Top             =   5055
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   28
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":19F5
      MousePointer    =   99  'Custom
      Top             =   5055
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   26
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":1B47
      MousePointer    =   99  'Custom
      Top             =   4830
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":1C99
      MousePointer    =   99  'Custom
      Top             =   4605
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":1DEB
      MousePointer    =   99  'Custom
      Top             =   4380
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":1F3D
      MousePointer    =   99  'Custom
      Top             =   4170
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   18
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":208F
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   16
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":21E1
      MousePointer    =   99  'Custom
      Top             =   3750
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   14
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":2333
      MousePointer    =   99  'Custom
      Top             =   3525
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":2485
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   6720
      MouseIcon       =   "frmCrearPersonaje.frx":25D7
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   8
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":2729
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   6
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":287B
      MousePointer    =   99  'Custom
      Top             =   2655
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   2
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":29CD
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   4
      Left            =   6720
      MouseIcon       =   "frmCrearPersonaje.frx":2B1F
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":2C71
      MousePointer    =   99  'Custom
      Top             =   1995
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   1
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":2DC3
      MousePointer    =   99  'Custom
      Top             =   1995
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   27
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":2F15
      MousePointer    =   99  'Custom
      Top             =   4830
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   25
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":3067
      MousePointer    =   99  'Custom
      Top             =   4605
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   23
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":31B9
      MousePointer    =   99  'Custom
      Top             =   4380
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   21
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":330B
      MousePointer    =   99  'Custom
      Top             =   4170
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   19
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":345D
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   17
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":35AF
      MousePointer    =   99  'Custom
      Top             =   3750
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   15
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":3701
      MousePointer    =   99  'Custom
      Top             =   3525
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   13
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":3853
      MousePointer    =   99  'Custom
      Top             =   3300
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   11
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":39A5
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   9
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":3AF7
      MousePointer    =   99  'Custom
      Top             =   2880
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   7
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":3C49
      MousePointer    =   99  'Custom
      Top             =   2655
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   3
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":3D9B
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   255
      Index           =   5
      Left            =   6000
      MouseIcon       =   "frmCrearPersonaje.frx":3EED
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   36
      Left            =   6750
      MouseIcon       =   "frmCrearPersonaje.frx":403F
      MousePointer    =   99  'Custom
      Top             =   5910
      Width           =   180
   End
   Begin VB.Image boton 
      Height          =   375
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmCrearPersonaje.frx":4191
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   1875
   End
   Begin VB.Image boton 
      Height          =   330
      Index           =   0
      Left            =   9000
      MouseIcon       =   "frmCrearPersonaje.frx":42E3
      MousePointer    =   99  'Custom
      Top             =   8280
      Width           =   2040
   End
End
Attribute VB_Name = "frmCrearPersonaje1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2009 Juan Andres Dalmasso (CHOTS)
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public SkillPoints As Byte


Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

CheckData = True


End Function

Private Sub boton_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 0
        
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
         UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
            UserName = RTrim$(UserName)
            MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        If SelectedHead = 0 Then
            MsgBox "Debes seleccionar una cara antes de crear personaje"
            Exit Sub
        End If
        
        
        UserRaza = lstRaza.List(lstRaza.listIndex)
        UserSexo = lstGenero.List(lstGenero.listIndex)
        UserClase = lstProfesion.List(lstProfesion.listIndex)

        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.listIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            If CheckDatos() Then

    UserPassword = txtPasswd.Text

    UserEmail = txtCorreo.Text
    
    If Not CheckMailString(UserEmail) Then
            MsgBox "Direccion de mail invalida."
            Exit Sub
    End If
    
#If UsarWrench = 1 Then
    frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
#End If

    'SendNewChar = True
    EstadoLogin = CrearNuevoPj
    
    Me.MousePointer = 11

    EstadoLogin = CrearNuevoPj

#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State <> sckConnected Then
#End If
        MsgBox "Error: Se ha perdido la conexion con el server."
        Unload Me
        
    Else
        Call login(RandomCode)
    End If
End If
        End If
        
        
        
    Case 1
        frmConnect.Picture = LoadPicture(App.Path & "\Graficos\conectar.jpg")
        Me.Visible = False
        
    Case 2
        Call Audio.PlayWave(SND_DICE)
        Call tirarDados
      
End Select



End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize Timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function

Private Sub Form_Load()

Me.Picture = LoadPicture(App.Path & "\Graficos\CP-Interface.jpg")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.listIndex = 1

SkillPoints = 10
puntos.Caption = SkillPoints
Call tirarDados

SelectedHead = 0

End Sub


Private Sub headPbx_Click()
Call SiguienteCara
End Sub

Private Sub lstRaza_Click()
Call SetDadosFinal
Call RenderizarCaras
End Sub
Private Sub lstGenero_click()
Call RenderizarCaras
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub tirarDados()

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call SendData(ClientPackages.tirarDados)
    End If

End Sub

Private Sub Command1_Click(index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

puntos.Caption = SkillPoints
End Sub

Function CheckDatos() As Boolean

If txtPasswd.Text <> txtPasswdCheck.Text Then
    MsgBox "Los passwords que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If

If txtCorreo.Text <> txtCorreoCheck.Text Then
    MsgBox "Los Mails que tipeo no coinciden, por favor vuelva a ingresarlos."
    Exit Function
End If
'LEAN--> Limpiando
'If txtPreg.Text = txtResp.Text Then
 '   MsgBox "La pregunta y Respuesta secreta son iguales, por favor vuelva a ingresarlos."
 '   Exit Function
'End If

'If Len(txtResp.Text) < 3 Then
 '   MsgBox "La respuesta debe tener al menos 3 letras."
'    Exit Function
'End If

'If Not TextoValido(txtPasswd.Text) Then
'    MsgBox "Tu password no es seguro"
 '   Exit Function
'End If

'If Not TextoValido(txtPreg.Text) Then
'    MsgBox "Tu pregunta no es segura"
  '  Exit Function
'End If

'If Not TextoValido(txtResp.Text) Then
'    MsgBox "Tu respuesta no es segura"
'    Exit Function
'End If

If Not MailValido(txtCorreo.Text) Then
    MsgBox "Tu mail no es seguro"
    Exit Function
End If

If puntos.Caption > 0 Then
    MsgBox "Aún tienes puntos que asignar"
    Exit Function
End If

CheckDatos = True

End Function

Function TextoValido(ByVal Texto As String) As Boolean
   
   If UCase$(Texto) = "ASD" Or _
      UCase$(Texto) = "ASDASD" Or _
      UCase$(Texto) = "ASDASD123" Or _
      UCase$(Texto) = "123456" Or _
      UCase$(Texto) = "ASD123" Or _
      UCase$(Texto) = "QWERTY" Or _
      UCase$(Texto) = "123" Or _
      UCase$(Texto) = "AAAAAA" Or _
      UCase$(Texto) = "AAA" Then
      TextoValido = False
      Exit Function
   End If

   TextoValido = True
End Function

Function MailValido(ByVal mail As String) As Boolean
   
   If UCase$(mail) = "A@A.A" Or _
      UCase$(mail) = "ASD@ASD.ASD" Or _
      UCase$(mail) = "ASD@ASD.COM" Or _
      UCase$(mail) = "A@A.COM" Or _
      UCase$(mail) = "A@ASD.COM" Or _
      UCase$(mail) = "A@ASD.ASD" Or _
      UCase$(mail) = "ASDASD@ASD.COM" Or _
      UCase$(mail) = "ASD123@ASD.ASD" Or _
      UCase$(mail) = "AAA@AAA.AAA" Or _
      UCase$(mail) = "AAA@AAA.COM" Then
      MailValido = False
      Exit Function
   End If

   MailValido = True
End Function

Private Sub botonn_Click()
    frmCrearPersonaje1.Visible = False
End Sub



'Private Sub txtPreg_GotFocus()
'    MsgBox ("ATENCION! ROBO DE PERSONAJES" & vbNewLine & "Lapsus Corp recomienda seleccionar una pregunta y respuesta que sólo usted sepa" & vbNewLine & "Es la única manera que tendrá usted de recuperar su personaje (y que tendrán los usuarios ajenos de robárselo)")
'End Sub

Public Sub SetDadosFinal()
'CHOTS | Label Final de Atributos
With frmCrearPersonaje1
    Select Case UCase$(.lstRaza.List(.lstRaza.listIndex))
      Case "HUMANO"
            .lblFuerza2.Caption = .lbFuerza.Caption + 1
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 1
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 1
            .lblCarisma2.Caption = .lbCarisma.Caption
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 2
         Case "ELFO"
            .lblFuerza2.Caption = .lbFuerza.Caption
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 4
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 2
            .lblCarisma2.Caption = .lbCarisma.Caption + 2
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 1
        Case "ELFO OSCURO"
            .lblFuerza2.Caption = .lbFuerza.Caption + 2
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 2
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 2
            .lblCarisma2.Caption = .lbCarisma.Caption - 3
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 1
        Case "ENANO"
            .lblFuerza2.Caption = .lbFuerza.Caption + 3
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 1
            .lblInteligencia2.Caption = .lbInteligencia.Caption - 5
            .lblCarisma2.Caption = .lbCarisma.Caption - 2
            .lblConstitucion2.Caption = .lbConstitucion.Caption + 3
        Case "GNOMO"
            .lblFuerza2.Caption = .lbFuerza.Caption
            .lblAgilidad2.Caption = .lbAgilidad.Caption + 3
            .lblInteligencia2.Caption = .lbInteligencia.Caption + 3
            .lblCarisma2.Caption = .lbCarisma.Caption + 1
            .lblConstitucion2.Caption = .lbConstitucion.Caption
        Case Else
            .lblFuerza2.Caption = .lbFuerza.Caption
            .lblAgilidad2.Caption = .lbAgilidad.Caption
            .lblInteligencia2.Caption = .lbInteligencia.Caption
            .lblCarisma2.Caption = .lbCarisma.Caption
            .lblConstitucion2.Caption = .lbConstitucion.Caption
    End Select
    .lblMaxFuerza.Caption = 35
    .lblMaxAgi.Caption = 35

End With

'CHOTS | Label Final de Atributos
End Sub

Sub RenderizarCaras()
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 64
SR.Bottom = 16

DR.Left = 0
DR.Top = 0
DR.Right = 64
DR.Bottom = 16

With frmCrearPersonaje1
Select Case UCase$(.lstGenero.List(.lstGenero.listIndex))
    Case "HOMBRE"
    Select Case UCase$(.lstRaza.List(.lstRaza.listIndex))
        Case "HUMANO"
            SelectedHead = 1
            
        Case "ELFO"
            SelectedHead = 101
            
        Case "ELFO OSCURO"
            SelectedHead = 202
            
        Case "ENANO"
            SelectedHead = 301
            
        Case "GNOMO"
            SelectedHead = 401
            
    End Select
    
    Case "MUJER"
        Select Case UCase$(.lstRaza.List(.lstRaza.listIndex))
        Case "HUMANO"
            SelectedHead = 70
    
        Case "ELFO"
            SelectedHead = 170
            
        Case "ELFO OSCURO"
            SelectedHead = 270
            
        Case "ENANO"
            SelectedHead = 470

        Case "GNOMO"
            SelectedHead = 370

            
    End Select
End Select
End With

Call DrawGrhtoHdc(headPbx.hWnd, headPbx.hdc, HeadData(SelectedHead).Head(1).GrhIndex, SR, DR)
headPbx.Refresh
End Sub

Sub SiguienteCara()
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 64
SR.Bottom = 16

DR.Left = 0
DR.Top = 0
DR.Right = 64
DR.Bottom = 16

If SelectedHead = 0 Then
    RenderizarCaras
    Exit Sub
End If
    
With frmCrearPersonaje1
Select Case UCase$(.lstGenero.List(.lstGenero.listIndex))
    Case "HOMBRE"
    Select Case UCase$(.lstRaza.List(.lstRaza.listIndex))
        Case "HUMANO"
            If SelectedHead = 30 Then
                SelectedHead = 1
            Else
                SelectedHead = SelectedHead + 1
            End If
            
        Case "ELFO"
            If SelectedHead = 201 Then
                SelectedHead = 101
            Else
                SelectedHead = SelectedHead + 1
            End If
            If SelectedHead = 113 Then SelectedHead = 201
            
        Case "ELFO OSCURO"
            If SelectedHead = 209 Then
                SelectedHead = 202
            Else
                SelectedHead = SelectedHead + 1
            End If
            
        Case "ENANO"
            If SelectedHead = 305 Then
                SelectedHead = 301
            Else
                SelectedHead = SelectedHead + 1
            End If
            
        Case "GNOMO"
           If SelectedHead = 406 Then
                SelectedHead = 401
            Else
                SelectedHead = SelectedHead + 1
            End If
            
    End Select
    
    Case "MUJER"
        Select Case UCase$(.lstRaza.List(.lstRaza.listIndex))
        Case "HUMANO"
           If SelectedHead = 76 Then
                SelectedHead = 70
            Else
                SelectedHead = SelectedHead + 1
            End If
    
        Case "ELFO"
            If SelectedHead = 176 Then
                SelectedHead = 170
            Else
                SelectedHead = SelectedHead + 1
            End If
            
        Case "ELFO OSCURO"
            If SelectedHead = 280 Then
                SelectedHead = 270
            Else
                SelectedHead = SelectedHead + 1
            End If
            
        Case "ENANO"
            If SelectedHead = 475 Then
                SelectedHead = 470
            Else
                SelectedHead = SelectedHead + 1
            End If

        Case "GNOMO"
            If SelectedHead = 372 Then
                SelectedHead = 370
            Else
                SelectedHead = SelectedHead + 1
            End If

            
    End Select
End Select
End With

Call DrawGrhtoHdc(headPbx.hWnd, headPbx.hdc, HeadData(SelectedHead).Head(1).GrhIndex, SR, DR)
headPbx.Refresh
End Sub
