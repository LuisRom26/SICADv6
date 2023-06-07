<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>

 <%
			  		set cnn = server.CreateObject("ADODB.CONNECTION")
					set rst = server.CreateObject("ADODB.RECORDSET")
					Archivo = request.ServerVariables("APPL_PHYSICAL_PATH") & "/config.txt"
					set ConFile = createobject ("scripting.filesystemobject")
					set Fichero = ConFile.OpenTextFile(Archivo)
					TextoFichero = Fichero.ReadAll()
									
					Fichero.Close()
					
									
					strConexion = TextoFichero
					cnn.open strConexion
				
					rst.CursorLocation = 2
					rst.CursorType = 0
					rst.LockType = 3

					Nuevo = request("cmdUnidades")
					
					rst.open "Select * from Empleados where idEmpleados = '" & Nuevo & "'", cnn
					'response.Write("Select * from usrSicad where nombre = '" & Nuevo & "'")
						AreaNueva = rst.fields("Area")
						NivelNuevo = rst.fields("NivelCorr")
						Nombre = rst.fields("Nombre") & chr(32) & rst.fields("ApellidoP") & chr(32) & rst.fields("ApellidoM")
					rst.close
					
					
						rst.open "SELECT * FROM Oficios Where Folio = " & Session("FolioFW") , cnn
							rst.fields("Turnado") = Nuevo
							rst.fields("Area") = AreaNueva
							rst.fields("ComReasigna") = request("txtComentario")
							rst.fields("QuienReasigna") = session("Usuario")
							rst.fields("FechaReasigna") = date
							
							
						rst.update
						
						rst.close

		'Para el log de actividades
		rst.open "select * from Monitoreo where idMonitoreo = 1", cnn
			rst.addnew
				rst.fields("usrID") = Session("IDUsuario")
				rst.fields("usrNombre") = Session("Usuario")
				rst.fields("usrArea") = Session("Area")
				rst.fields("Comentario") =  "El usuario reasignó a " & Nombre & " el folio No. " & Session("FolioFW") & " del módulo de correspondencia.<br><br><em><strong>Texto del comentario:</strong></em><br>"  & request("txtComentario")
				rst.fields("Fecha") = date
				rst.fields("Hora") = time
				rst.fields("Modulo") = "REASIGNAR FOLIO"
			rst.update
		rst.close						
						
				
				%>

	<!-- Para los redireccionamientos -->
		<script language="Javascript">
			window.open('BuscaCorr.asp', '_self');
 		</script>


</body>
</html>
