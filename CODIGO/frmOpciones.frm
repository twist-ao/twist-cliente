VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5400
   ClientLeft      =   4095
   ClientTop       =   1185
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Soporte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3825
      Width           =   2790
   End
   Begin VB.CommandButton Teclas 
      BackColor       =   &H000080FF&
      Caption         =   "Teclas Configurables"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2715
      Width           =   2775
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Audio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   4215
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Musica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1000
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Sonidos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1000
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         Max             =   100
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   11
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Manual del Juego"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":02A4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4365
      Width           =   2790
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Fps"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   650
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   4215
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Limitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   250
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "No Limitar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Mapa del Mundo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   960
      MouseIcon       =   "frmOpciones.frx":03F6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3270
      Width           =   2790
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   960
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit
Private loading As Boolean

Private Sub Command3_Click()
    Shell ("cmd /c start http://twistao.com/soporte.php"), vbHide
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then Unload Me
End Sub

Private Sub Check1_Click(index As Integer)
    If Not loading Then _
        Call Audio.PlayWave(SND_CLICK)
    
    Select Case index
        Case 0
            If Check1(0).Value = vbUnchecked Then
                Audio.MusicActivated = False
                Slider1(0).Enabled = False
            ElseIf Not Audio.MusicActivated Then  'Prevent the music from reloading
                Audio.MusicActivated = True
                Slider1(0).Enabled = True
                Slider1(0).Value = Audio.MusicVolume
            End If
        
        Case 1
            If Check1(1).Value = vbUnchecked Then
                Audio.SoundActivated = False
                RainBufferIndex = 0
                frmMain.IsPlaying = PlayLoop.plNone
                Slider1(1).Enabled = False
            Else
                Audio.SoundActivated = True
                Slider1(1).Enabled = True
                Slider1(1).Value = Audio.SoundVolume
            End If
    End Select
End Sub

Private Sub Command1_Click(index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case index
    Case 2
        FrmMapa.Show
    Case 3
        Shell ("cmd /c start http://www.twistao.com/manual.php"), vbHide
End Select
End Sub

Private Sub Command2_Click()
Unload frmOpciones
End Sub

Private Sub Form_Load()
    loading = True      'Prevent sounds when setting check's values
    
    If Audio.MusicActivated Then
        Check1(0).Value = vbChecked
        Slider1(0).Enabled = True
        Slider1(0).Value = Audio.MusicVolume
    Else
        Check1(0).Value = vbUnchecked
        Slider1(0).Enabled = False
    End If
    
    If Audio.SoundActivated Then
        Check1(1).Value = vbChecked
        Slider1(1).Enabled = True
        Slider1(1).Value = Audio.SoundVolume
    Else
        Check1(1).Value = vbUnchecked
        Slider1(1).Enabled = False
    End If
    
    If FPSFast Then
        Option2.Value = True
        Option3.Value = False
    Else
        Option3.Value = True
        Option2.Value = False
    End If
    
    loading = False
    
End Sub

Private Sub Option2_Click()
 FPSFast = True
 Call WriteVar(App.Path & "\INIT\FPS.dat", "INIT", "Fast", 1)
 Call LoadGrhData
 frmOpciones.Visible = False
 
End Sub

Private Sub Option3_Click()
 FPSFast = False
 Call WriteVar(App.Path & "\INIT\FPS.dat", "INIT", "Fast", 0)
 Call LoadGrhData
 frmOpciones.Visible = False
End Sub

Private Sub Option5_Click()

End Sub

Private Sub Slider1_Change(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub

Private Sub Slider1_Scroll(index As Integer)
    Select Case index
        Case 0
            Audio.MusicVolume = Slider1(0).Value
        Case 1
            Audio.SoundVolume = Slider1(1).Value
    End Select
End Sub

Private Sub Teclas_Click()
Call frmCustomKeys.Show(vbModeless, frmMain)
End Sub
