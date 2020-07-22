VERSION 5.00
Begin VB.Form frmcomentario 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir comentarios"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8865
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstComentario 
      BackColor       =   &H00000000&
      ForeColor       =   &H0080FF80&
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   8655
   End
   Begin VB.TextBox txtComentario 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAniadir 
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Añadir"
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
      MICON           =   "frmcomentario.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdeliminarselecionado 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Eliminar Seleciónado"
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
      MICON           =   "frmcomentario.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdEliminarTodo 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Eliminar Todo"
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
      MICON           =   "frmcomentario.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdAplicar 
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Guardar"
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
      MICON           =   "frmcomentario.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdCancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmcomentario.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdCargarComentarios 
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cargar "
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
      MICON           =   "frmcomentario.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblComentario 
      BackStyle       =   0  'Transparent
      Caption         =   "Comentario:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   200
      Width           =   1215
   End
End
Attribute VB_Name = "frmcomentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* comentarios en  Virtual Martin temporize v1.7
'*
'*
'***************************************************************************

Private Sub cmdAniadir_Click()
If Not (txtComentario.Text = "") Then
lstComentario.AddItem txtComentario.Text
End If
txtComentario.Text = ""
End Sub

Private Sub cmdAplicar_Click()
Dim r As Integer
Open "comentarios.txt" For Output As 1
 For r = 0 To lstComentario.ListCount - 1
 Print #1, lstComentario.List(r)
 Next r
Close #1
Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdCargarComentarios_Click()
lstComentario.Clear
Dim cargar As String

Open "comentarios.txt" For Input As 1
 Do While Not EOF(1)
       Line Input #1, cargar
       lstComentario.AddItem cargar
       Loop
       Close #1
End Sub

Private Sub cmdeliminarselecionado_Click()
If Not (lstComentario.ListIndex = -1) Then
 Select Case MsgBox(lenguaje_Menu(273) _
 , vbYesNo + vbInformation)
  Case (vbYes)
   lstComentario.RemoveItem (lstComentario.ListIndex)
 End Select
End If
End Sub

Private Sub cmdEliminarTodo_Click()
If Not (lstComentario.ListIndex <= -1) Then
 Select Case MsgBox(lenguaje_Menu(274) _
 , vbYesNo + vbInformation)
  Case (vbYes)
   lstComentario.Clear
 End Select
End If
End Sub

Private Sub Form_Load()
Me.Icon = frmprograma.Icon
cmdCargarComentarios_Click
cargarIdioma
End Sub
Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(265)
lblComentario.Caption = lenguaje_Menu(266)
cmdAniadir.Caption = lenguaje_Menu(267)
cmdCargarComentarios.Caption = lenguaje_Menu(268)
cmdCancelar.Caption = lenguaje_Menu(269)
cmdeliminarselecionado.Caption = lenguaje_Menu(270)
cmdEliminarTodo.Caption = lenguaje_Menu(271)
cmdAplicar.Caption = lenguaje_Menu(272)
End Sub

Private Sub lstComentario_Click()
txtComentario.Text = lstComentario.List(lstComentario.ListIndex)
End Sub

Private Sub lstComentario_Scroll()
lstComentario_Click
End Sub
