VERSION 5.00
Begin VB.Form frmDonativos 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Donativos"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4050
   Icon            =   "frmDonativos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pdonar 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   600
      MouseIcon       =   "frmDonativos.frx":0CCA
      Picture         =   "frmDonativos.frx":0FD4
      ScaleHeight     =   1155
      ScaleWidth      =   2925
      TabIndex        =   3
      Top             =   970
      Width           =   2955
   End
   Begin VB.PictureBox ptargeta 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   960
      Picture         =   "frmDonativos.frx":C0F2
      ScaleHeight     =   225
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   2520
      Width           =   2175
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcolaborar 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Colaborar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDonativos.frx":C8ED
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAceptar 
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&Aceptar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   14737632
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDonativos.frx":C909
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Amo mucho a EE:UU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   7125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "para cumplir mi sue�o de ir a EE:UU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   600
      TabIndex        =   6
      Top             =   120
      Width           =   7125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "con cuenta propia..."
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   2745
   End
   Begin VB.Label lblcard 
      BackStyle       =   0  'Transparent
      Caption         =   "Con tarjetas de cr�ditos"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Width           =   2985
   End
End
Attribute VB_Name = "frmDonativos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Para realizar donac�ones para el proyecto Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Declare Function ShellExecute Lib _
 "shell32.dll" Alias "ShellExecuteA" _
 (ByVal hwnd As Long, ByVal lpOperation As String, _
 ByVal lpFile As String, ByVal lpParameters As String, _
 ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdAceptar_Click()
 Unload Me
End Sub

Private Sub cmdcolaborar_Click()
 ptargeta_Click
End Sub

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 Call cargarIdioma
End Sub

Private Sub Label1_Click()
 ptargeta_Click
End Sub

Private Sub lblcard_Click()
 ptargeta_Click
End Sub

Private Sub pdonar_Click()
 ptargeta_Click
End Sub

Private Sub ptargeta_Click()
 Dim x As String
 x = ShellExecute(Me.hwnd, "Open" _
 , "http://martinsoft0.blogspot.com/p/donar.html", _
 &O0, &O0, 0)
 Unload Me
End Sub
Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(310)
Label2.Caption = lenguaje_Menu(311)
Label3.Caption = lenguaje_Menu(312)
Label1.Caption = lenguaje_Menu(313)
lblcard.Caption = lenguaje_Menu(314)
cmdcolaborar.Caption = lenguaje_Menu(315)
cmdAceptar.Caption = lenguaje_Menu(316)
End Sub

