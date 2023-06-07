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
		
		Cuantos = 0
		
		rst.open "select * from DetalleMinutarios where idMinutario = " &  Session("MinutarioID") & " and Anio = " & year(date), cnn
			do while not rst.eof
				Cuantos = Cuantos + 1
				rst.movenext
			loop
		rst.close
		
		
		rst.open "select * from DetalleMinutarios where idDetalleMinutarios=1", cnn
		
		for k = 1 to request("cmbCantOf")
			rst.addnew
				Cuantos = Cuantos + 1
				rst.fields("idMinutario") = Session("MinutarioID")
				rst.fields("Anio") = year(date)
				rst.fields("Consecutivo") = Cuantos
				rst.fields("Fecha") = date
				rst.fields("AreaSolicitante") = Session("Area") 'Aqui es donde está el error 
				rst.fields("Seguimiento") = session("Usuario")
			rst.update
		next 
		rst.close
		
		'Para el log de actividades
		rst.open "select * from Monitoreo where idMonitoreo = 1", cnn
			rst.addnew
				rst.fields("usrID") = Session("IDUsuario")
				rst.fields("usrNombre") = Session("Usuario")
				rst.fields("usrArea") = Session("Area")
				rst.fields("Comentario") =  "El usuario solicitó " & request("cmbCantOf") & " número(s) de oficio(s)"
				rst.fields("Fecha") = date
				rst.fields("Hora") = time
				rst.fields("Modulo") = "SOLICITUD NUMERO OFICIO"
			rst.update
		rst.close

		
		'response.Redirect("Minutario.asp?Seccion=4&Parent=MainCorrespondencia.asp&SeccPar=2")
%>

<script language="Javascript">
	window.open('BuscaMin.asp', '_parent');
</script>

</body>
</html>
