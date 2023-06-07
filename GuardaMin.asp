﻿<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CharSet = "utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/> 
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
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
	
	response.write("<script>")
	response.write("window.open('default.asp','_top')")
	response.write("</script>") 
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

if request("cmbSolicito") = 0 then
	response.Write("<div align='center'><br /><br /><br /><br /><h3>Ha ocurrido el siguiente error:</h3>es necesario seleccionar a un responsable para dar seguimiento al acuse de recibido</div><br /><div align='center'><a href='javascript:history.back(-1);' class='btn btn-outline-info'>Regresar</a></div>")
else
rst.open "select * from detalleminutarios where idDetalleMinutarios = " & request("IDMin"), cnn
NumOf = rst.fields("Consecutivo") & "/" & rst.fields("Anio")

 rst.fields("Fase") = 1
 rst.fields("Destinatario") =  ucase(request("txtDestinatario"))
 rst.fields("DepeDestino") = ucase(request("txtDependencia"))
 'rst.fields("Fecha") = date
 rst.fields("Solicitante") = cint(mid(request("cmbSolicito"),3, len(request("cmbSolicito"))-2))
 rst.fields("Seguimiento") = request("txtSolicitante")
 rst.fields("AreaSolicitante") = cint(mid(request("cmbSolicito"),1,2))
 rst.fields("RespuestaA") = ucase(request("txtRespuesta"))
 rst.fields("NumOficio") = ucase(request("txtOficio"))
 rst.fields("Asunto") = ucase(request("txtAsunto"))
 
 rst.update
 rst.close
 
 		'Para el log de actividades
		rst.open "select * from Monitoreo where idMonitoreo = 1", cnn
			rst.addnew
				rst.fields("usrID") = Session("IDUsuario")
				rst.fields("usrNombre") = Session("Usuario")
				rst.fields("usrArea") = Session("Area")
				rst.fields("Comentario") =  "El usuario asignó el numero de oficio " & NumOf & "<br><br><em><strong>Datos de la asignación</strong></em><br>Destinatario: " & ucase(request("txtDestinatario")) & "<br><br>Dependencia: " & ucase(request("txtDependencia")) & "<br><br>Asunto: " & ucase(request("txtAsunto")) & "<br><br>Respuesta a oficio: " & ucase(request("txtRespuesta"))
				rst.fields("Fecha") = date
				rst.fields("Hora") = time
				rst.fields("Modulo") = "ASIGNACION NUMERO OFICIO"
			rst.update
		rst.close

%>
<script language="Javascript">
	window.open('BuscaMin.asp', '_parent');
</script>
<%end if%>

</body>
</html>
