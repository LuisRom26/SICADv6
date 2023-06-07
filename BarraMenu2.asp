
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CharSet = "utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/> 
<title>SICAD v5.01b</title>

<link type="text/css" href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" rel="stylesheet" /> 
<link href="CSS/Estilos.css" media="screen" type="text/css" rel="stylesheet" />
<link href="CSS/impresora.css" media="print" type="text/css" rel="stylesheet" />
<link href="bootstrap-4.0.0-alpha.6-dist/css/bootstrap.css" media="screen" type="text/css" rel="stylesheet"  />
<link href="bootstrap-4.0.0-alpha.6-dist/js/bootstrap.js" type="text/javascript" />


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

