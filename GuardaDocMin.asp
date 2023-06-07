﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CharSet = "utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/> 

<title>SICAD v3.0.1b</title>

 
<!--<link type="text/css" href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" rel="stylesheet" /> -->
<link type="text/css" href="css/jquery-ui.css" rel="stylesheet" />
<link href="CSS/Estilos.css" media="screen" type="text/css" rel="stylesheet" />
<link href="CSS/EstilosClick.css" media="screen" type="text/css" rel="stylesheet" />
<link href="CSS/impresora.css" media="print" type="text/css" rel="stylesheet" />



<style>
.alternar:hover{ background-color:#d0d4d4;}
</style>

<%
' Generador de claves aleatorias

Function generadordeclaves(longituddeclave)
' Nota para los principientes : el simpolo "_" es el de continuación de linea 
' Definicion del array
Dim numerodecaracteres 
Dim salida
Dim char_array
char_array = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
"K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
"U", "V", "W", "X", "Y", "Z", "a", "b", "c", "d", "e", "f", "g", "h",  "i", "j", "k", "l", "m", _
"n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")


Randomize()

Do While Len(salida) < longituddeclave
salida = salida & char_array(Int(36 * Rnd()))
Loop

' establecemos el valor del resultado a devolver
generadordeclaves = salida
End Function

%>




<style type="text/css">
body {
	
	background-repeat: repeat;
	
}
</style>
</head>

<body>
<%

if Session("SICAD_Active") <> 1 then
	Session("SICAD_Active") = 0
	response.Redirect("default.asp")
end if
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

			


 	 rst.open "select * from detalleminutarios where idDetalleMinutarios = " & request("IDMin"), cnn
	 NumOf = rst.fields("NumOficio")
	 rst.fields("Fase") = 2
	 Ar = rst.fields("ArchivoTMP")
	 NumOf = "DG-" & right("0000" & rst.fields("Consecutivo"),4) & "/"  & rst.fields("Anio")
	 rst.update
	
	NumOfNew = ""
	for k = 1 to len(NumOf)
		if mid(NumOf, k, 1) = "/" then
			NumOfNew = NumOfNew & "-"
		else
			NumOfNew = NumOfNew & mid(NumOf, k, 1)
		end if
	next
	
	dim fs
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	fs.CopyFile request.ServerVariables("APPL_PHYSICAL_PATH") & "/Minutario/TMP/" & Ar, request.ServerVariables("APPL_PHYSICAL_PATH") & "/Minutario/" & NumOfNew & right(Ar,4)
	set fs=nothing
	rst.close
	'response.Write(NumOfNew)
	
	
	
		'Para el log de actividades
		
	

%>

<script language="Javascript">
	window.open('BuscaMin.asp', '_Self');
</script>

</body>
</html>
