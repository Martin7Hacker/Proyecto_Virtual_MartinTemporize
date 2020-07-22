VERSION 5.00
Begin VB.Form frmidioma 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instancia de Idioma"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Virtual_Martin_temporize.ChameleonBtn cmdCargarIdioma 
      Height          =   255
      Left            =   7560
      TabIndex        =   12
      Top             =   105
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "..."
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
      MICON           =   "frmIdiomas.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdrenombrar 
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   465
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "&Renombrar"
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
      MICON           =   "frmIdiomas.frx":001C
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
      Left            =   6120
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Aplicar Idioma"
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
      MICON           =   "frmIdiomas.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdguardar 
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Guardar Archivo"
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
      MICON           =   "frmIdiomas.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcargar 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "&Cargar"
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
      MICON           =   "frmIdiomas.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Virtual_Martin_temporize.ChameleonBtn cmdcancelar 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
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
      MICON           =   "frmIdiomas.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2790
      Left            =   7680
      TabIndex        =   6
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton cmdidioma 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   7815
   End
   Begin VB.TextBox txtnombre 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   90
      Width           =   6255
   End
   Begin VB.TextBox txtvalor 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   450
      Width           =   5175
   End
   Begin VB.ListBox lstidioma 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFC0&
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo:"
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
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   510
   End
End
Attribute VB_Name = "frmidioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Configurar idioma en  Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Dim r As Integer
Private Sub cmdAplicar_Click()
Dim int_c As Integer
Dim cargar As String
lstidioma.Clear
Open txtnombre.Text For Input As 1
 Do While Not EOF(1)
  
       Line Input #1, cargar
       Lenguage.lenguaje_Menu(int_c) = cargar
       lstidioma.AddItem cargar
       int_c = int_c + 1
       Loop
       Close #1
       int_c = 0
       Lenguage.definir_lenguage_opciones
       frmprograma.cargarIdioma
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdcargar_Click()
Dim cargar As String
lstidioma.Clear
Open txtnombre.Text For Input As 1
 Do While Not EOF(1)
  
       Line Input #1, cargar
       lstidioma.AddItem cargar
       Loop
       Close #1
End Sub

Private Sub cmdCargarIdioma_Click()
frmCargarIdioma.Show 1
End Sub

Private Sub cmdguardar_Click()
Open txtnombre.Text For Output As 1
 For r = 0 To 385
 Lenguage.lenguaje_Menu(r) = lstidioma.List(r)
 Print #1, Lenguage.lenguaje_Menu(r)
 Next r
Close #1
End Sub

Private Sub cmdrenombrar_Click()
lstidioma.List(lstidioma.ListIndex) = txtvalor.Text
End Sub

Private Sub Form_Load()
Me.Icon = frmprograma.Icon
For r = 0 To 385
lstidioma.AddItem Lenguage.lenguaje_Menu(r)
 Next r
 Call cargarIdioma
 VScroll1_Scroll
   txtnombre.Text = Lenguage.sel
End Sub

Private Sub cargarPrograma()
Me.Icon = frmprograma.Icon
End Sub

Private Sub lstidioma_Click()
lstidioma_Scroll
VScroll1.Value = lstidioma.ListIndex
End Sub

Private Sub lstidioma_Scroll()
txtvalor.Text = lstidioma.List(lstidioma.ListIndex)
VScroll1.Value = lstidioma.ListIndex

End Sub

Private Sub VScroll1_Change()
 On Error GoTo nose
 With VScroll1
 .Max = lstidioma.ListCount - 1
 .Min = 0
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 lstidioma.ListIndex = .Value
 End With
nose:
End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub

Private Sub cargarIdioma()
Me.Caption = lenguaje_Menu(232)
Label2.Caption = lenguaje_Menu(233)
Label1.Caption = lenguaje_Menu(234)
cmdrenombrar.Caption = lenguaje_Menu(235)
cmdidioma.Caption = lenguaje_Menu(236)
cmdCancelar.Caption = lenguaje_Menu(237)
cmdcargar.Caption = lenguaje_Menu(238)
cmdguardar.Caption = lenguaje_Menu(239)
cmdAplicar.Caption = lenguaje_Menu(240)
End Sub
