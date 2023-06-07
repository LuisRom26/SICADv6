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
			set rst2 = server.CreateObject("ADODB.RECORDSET")
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

			rst.open "select * from DetalleMinutarios where idDetalleMinutarios=" & request("IDMin"), cnn
			NumOf = rst.fields("Consecutivo") & "/" & rst.fields("Anio")
			rst.fields("Fase") = 1
			rst.update
			rst.close
			
		'Para el log de actividades
		rst.open "select * from Monitoreo where idMonitoreo = 1", cnn
			rst.addnew
				rst.fields("usrID") = Session("IDUsuario")
				rst.fields("usrNombre") = Session("Usuario")
				rst.fields("usrArea") = Session("Area")
				rst.fields("Comentario") =  "El usuario eliminó el acuse de recibido al oficio No. " & NumOf
				rst.fields("Fecha") = date
				rst.fields("Hora") = time
				rst.fields("Modulo") = "ELIMINA ACUSE DE RECIBIDO"
			rst.update
		rst.close

	
%>

<script language="Javascript">
	window.open('BuscaMin.asp', '_top');
</script>
</body>
</html>
