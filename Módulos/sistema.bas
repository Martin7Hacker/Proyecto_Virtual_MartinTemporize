Attribute VB_Name = "sistema"
'***************************************************************************
'*
'*
'* sistema de Virtual Martin temporize v1.7
'*
'*
'***************************************************************************
Public comando As String
Public tiempo As String
Public comentario As String
Public textoX As String
Public TextoY As String
Public ven As Byte

Public Sub tomarDatos()
 'le pasa el comando de disparo
 comando = frmfunciones.devolver_comando
 'le pasa el tiempo t en seg.
 tiempo = frmfunciones.DTPicker1.Minute
 'comentario del apagado
 comentario = frmfunciones.txtd.Text
End Sub

Public Sub ingresarDatos()
 'ingresar datos
 With frmprograma
 .liscomando.AddItem (comando)
 .listiempo.AddItem (tiempo)
 .lisdialogo.AddItem (comentario)
 End With
End Sub

Public Sub modificarDatos() 'funciones para modificar datos del timbre
 With frmprograma
 .liscomando.List(.liscomando.ListIndex) = sistema.comando
 .lisdialogo.List(.lisdialogo.ListIndex) = sistema.comentario
 .listiempo.List(.listiempo.ListIndex) = sistema.tiempo
 End With
End Sub

Public Sub eleminarDatos() 'Elimniar datos en memoria
 With frmprograma
 .liscomando.Clear
 .lisdialogo.Clear
 .listiempo.Clear
 End With
End Sub
