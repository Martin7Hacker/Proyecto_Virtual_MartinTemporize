VERSION 5.00
Begin VB.Form frmCargarIdioma 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   1710
      Left            =   0
      Pattern         =   "*.txt*"
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmCargarIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* cargar Idioma en  Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Private Sub cmdCargarIdioma()
Lenguage.sel = File1.List(File1.ListIndex)
frmidioma.txtnombre.Text = "idiomas\" & Lenguage.sel
guardar_Click
Unload Me
End Sub

Private Sub File1_Click()
Call cmdCargarIdioma
End Sub

Private Sub Form_Load()
Me.Icon = frmprograma.Icon
File1.Path = "idiomas\"
End Sub

Private Sub guardar_Click()
Dim r As Byte
Open "idiomas\inicio.inf" For Output As 1
 Print #1, "idiomas\" & Lenguage.sel
Close #1
End Sub
