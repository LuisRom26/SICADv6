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
		
		rst.open "select * from Oficios where Folio=" & request("Folio"), cnn
			rst.fields("Resuelto") = 0
			rst.fields("EnSeguimiento") = 0
			rst.update
		rst.close
		
%>

<script language="Javascript">
	window.open('Redirect.asp?Destino=BuscaCorr.asp', '_self');
</script>

</body>
</html>
