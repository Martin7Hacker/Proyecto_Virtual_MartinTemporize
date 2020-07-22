VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmVisorEventos 
   BackColor       =   &H00000000&
   Caption         =   "Visor de Eventos Programados Actualmente"
   ClientHeight    =   4680
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11100
   Icon            =   "frmvisor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   11100
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView ListView1 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12515
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   255
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Menu menu 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu imprimir 
         Caption         =   "&Imprimir"
      End
      Begin VB.Menu esp 
         Caption         =   "-"
      End
      Begin VB.Menu imprimirMas 
         Caption         =   "&Imprimir M�s"
      End
   End
End
Attribute VB_Name = "frmVisorEventos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*
'*
'* Visor de Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Dim d As Long

Private Sub Form_Load()
 Me.Icon = frmprograma.Icon
 cargar_datos
 Call cargarlenguaje
End Sub

Private Sub Form_Resize()
 On Error GoTo no_se
 ListView1.Width = Me.Width - 400
 ListView1.Height = Me.Height - 1400
no_se:
End Sub

Private Sub Form_Unload(Cancel As Integer)
 frmprograma.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
 frmprograma.Enabled = True
 Unload Me
End Sub

Private Sub cmdcancelar_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdguardarysalir_Click()
 frmprograma.guardard_Click
 End
End Sub

Private Sub cmdguardarysalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cmdsalir_Click()
 End
End Sub

Private Sub cmdsalir_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub cargar_datos()
Const espacio As String = "                               "
On Error GoTo no_se
 With frmprograma
 Dim ah As Integer
 Dim v As String
 Dim et As ListItem
 With ListView1.ColumnHeaders
 .Add , , lenguaje_Menu(257)
 .Add , , lenguaje_Menu(258)
 .Add , , lenguaje_Menu(259)
 .Add , , lenguaje_Menu(260)
 .Add , , lenguaje_Menu(261)
 .Add , , lenguaje_Menu(383)
 End With
 With ListView1
 .View = lvwReport
 .LabelEdit = lvwManual
 .MultiSelect = False
 .HideSelection = False
 End With
 ListView1.View = lvwReport
 For ah = 0 To .listado(0).ListCount - 1
 If .listado(1).List(ah) = lenguaje_Menu(18) Then
 v = "   "
 Else
 v = ""
 End If
 d = Int(ah) + 1
 With ListView1.ListItems.Add(, , lenguaje_Menu(264) & "_____ " & d)
 .SubItems(1) = frmprograma.listado(0).List(ah)
 .SubItems(2) = frmprograma.listado(1).List(ah)
 .SubItems(3) = lenguaje_Menu(382) & frmprograma.listado(2).List(ah)
 .SubItems(4) = frmprograma.listado(3).List(ah)
 .SubItems(5) = " " & frmprograma.domingo.List(ah) & " " & _
 frmprograma.lunes(0).List(ah) & " " & _
 frmprograma.martes.List(ah) & " " & _
 frmprograma.miercoles.List(ah) & " " & _
 frmprograma.jueves.List(ah) & " " & _
 frmprograma.viernes.List(ah) & " " & _
 frmprograma.sabado.List(ah)
 End With
 Next ah
 End With
no_se:
End Sub

Private Sub salir_op(ByVal dig As Byte)
 fc.comp_clave_fSalir False, dig, Hex(dig), 27, "1B", frmVisorEventos
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub imprimir_Click()
 ModImprimir.Imprimir_ListView
End Sub

Private Sub imprimirMas_Click()
 frmimpresor.Show 1
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
 salir_op KeyAscii
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 Select Case Button
 Case (2)
 PopupMenu menu
 End Select
End Sub

Private Sub cargarlenguaje()
Me.Caption = lenguaje_Menu(256)
imprimir.Caption = lenguaje_Menu(262)
imprimirMas.Caption = lenguaje_Menu(263)
End Sub
