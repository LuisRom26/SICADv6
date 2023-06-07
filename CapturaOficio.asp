<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>

<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">

<script language="javascript" type="text/javascript">
function validar(frm) { 

document.getElementById("boton").style.display="none";
document.getElementById("Archivo").style.display="none";
document.getElementById("Icono").style.display="none";
document.getElementById("Texto").style.display="inline";

document.images["imagen"].style.display="inline";
setTimeout('document.images["imagen"].src="Imagenes/progress_bar.gif"', 10);
//alert("hola");
}
</script>
<style>
.TextoNormal {
	font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;
	color:#333;
	font-size:12px;
	line-height:16px;
	font-weight:bold;
	
}
</style>


</head>

<body>
<%
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

<%
	response.buffer=true
	Func = Request("Func")
	if isempty(Func) Then
		Func = 1
	End if
	
	
Select Case Func
%>

<% Case 1%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center"></td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
</table>

<FORM ENCTYPE="multipart/form-data" ACTION="CapturaOficio.asp?func=2" METHOD="POST" id="form1" name="form1" onSubmit="return validar(this)"> 
<TABLE align="center" cellpadding="0" cellspacing="0" width="60%">
<tr>
  <td height="30" colspan="2" align="center" class="TextoBlanco" >&nbsp;</td>
  </tr>
<tr>
  <td width="26%" class="TextoNormal" >&nbsp;</td>
  <td width="74%" >&nbsp;</td>
</tr>
<TR>
  <TD rowspan="2" align="center" class="TextoNormal" ><div id="Icono"></div> <font size="2">
  </font></TD>
  <TD ><font size="2">
    <div id="Archivo"><input class="btn btn-outline-info" name="File1" size="45" type="file" /></div>
    </font></TD>
</TR>
<TR> 
  <TD align="left" valign="bottom" ><div id="boton" align="right"><br /><input class="btn btn-outline-info" type="submit" value="Adjuntar" /></div></TD>
</TR>
<TR> 
  <TD colspan="2" align="center" class="TextoNormal" ><div id="Texto" style="display:none"> Adjuntando archivo, espere un momento...<br /><br /></div><img alt="" width="500" height="20" id="imagen" style="display: none;" runat="server"/><BR>
  </TD>
</TR>
</TABLE>
<br /><br /></tr>


<%
Case 2

ForWriting = 2
adLongVarChar = 201
lngNumberUploaded = 0

'Get binary data from form 
noBytes = Request.TotalBytes 
binData = Request.BinaryRead (noBytes)
'convery the binary data to a string
Set RST = CreateObject("ADODB.Recordset")
LenBinary = LenB(binData)

if LenBinary > 0 Then
RST.Fields.Append "myBinary", adLongVarChar, LenBinary
RST.Open
RST.AddNew
RST("myBinary").AppendChunk BinData
RST.Update
strDataWhole = RST("myBinary")
End if
'Creates a raw data file for with all da
' ta sent. Uncomment for debuging. 
'Set fso = CreateObject("Scripting.FileSystemObject")
'Set f = fso.OpenTextFile(server.mappath(".") & "\raw.txt", ForWriting, True)
'f.Write strDataWhole
'set f = nothing
'set fso = nothing
'get the boundry indicator
strBoundry = Request.ServerVariables ("HTTP_CONTENT_TYPE")
lngBoundryPos = instr(1,strBoundry,"boundary=") + 8 
strBoundry = "--" & right(strBoundry,len(strBoundry)-lngBoundryPos)
'Get first file boundry positions.
lngCurrentBegin = instr(1,strDataWhole,strBoundry)
lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
Do While lngCurrentEnd > 0
'Get the data between current boundry an
' d remove it from the whole.
strData = mid(strDataWhole,lngCurrentBegin, lngCurrentEnd - lngCurrentBegin)
strDataWhole = replace(strDataWhole,strData,"")



'Get the full path of the current file.
lngBeginFileName = instr(1,strdata,"filename=") + 10
lngEndFileName = instr(lngBeginFileName,strData,chr(34)) 
'Make sure they selected at least one fi
' le. 
if lngBeginFileName = lngEndFileName and lngNumberUploaded = 0 Then

Response.Write "<H2> Ha ocurrido el siguiente error.</H2>"
Response.Write "Debes seleccionar un archivo para adjuntar"
Response.Write "<BR>Presiona el siguiente boton para realizar la seleccion del archivo."
Response.Write "<BR><BR><INPUT class='btn btn-outline-info' type='button' onclick='history.go(-1)' value='Reintentar' id='button'1 name='button'1>"
Response.End 
End if
'There could be one or more empty file b
' oxes. 
if lngBeginFileName <> lngEndFileName Then
strFilename = mid(strData,lngBeginFileName,lngEndFileName - lngBeginFileName)
'Creates a raw data file with data betwe
' en current boundrys. Uncomment for debug
' ing. 
'Set fso = CreateObject("Scripting.FileSystemObject")
'Set f = fso.OpenTextFile(server.mappath(".") & "\raw_" & lngNumberUploaded & ".txt", ForWriting, True)
'f.Write strData
'set f = nothing
'set fso = nothing

'Loose the path information and keep jus
' t the file name. 
tmpLng = instr(1,strFilename,"\")
Do While tmpLng > 0
PrevPos = tmpLng
tmpLng = instr(PrevPos + 1,strFilename,"\")
Loop

FileName = right(strFilename,len(strFileName) - PrevPos)


'Get the begining position of the file d
' ata sent.
'if the file type is registered with the
' browser then there will be a Content-Typ
' e
lngCT = instr(1,strData,"Content-Type:")

if lngCT > 0 Then
lngBeginPos = instr(lngCT,strData,chr(13) & chr(10)) + 4
Else
lngBeginPos = lngEndFileName
End if
'Get the ending position of the file dat
' a sent.
lngEndPos = len(strData) 

'Calculate the file size. 
lngDataLenth = lngEndPos - lngBeginPos
'Get the file data 
strFileData = mid(strData,lngBeginPos,lngDataLenth)
'Create the file. 

FileExt = ""
for k = 1 to len(FileName)
	if mid(FileName, len(FileName)-k + 1,1) = "." then
		exit for
	else
		FileExt = FileExt & mid(FileName, len(FileName) - k + 1,1)
	end if
next

tmp = ""
for k = 0 to len(FileExt)-1
	tmp = tmp & mid(FileExt, len(FileExt)-k, 1)
next 

FileExt = "." & tmp

'response.Write(FileExt & "<br><br>")
FileName = "CORR-TMP-" & generadordeclaves(8) & "-" & generadordeclaves(8) & "-" & Session.SessionID & FileExt
Session("NomArc") = FileName
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(server.mappath("..") & "\SICAD5\Oficios\TMP\" &_
 FileName, ForWriting, True)
f.Write strFileData

Set f = nothing
Set fso = nothing

lngNumberUploaded = lngNumberUploaded + 1

End if

'Get then next boundry postitions if any
' .
lngCurrentBegin = instr(1,strDataWhole,strBoundry)
lngCurrentEnd = instr(lngCurrentBegin + 1,strDataWhole,strBoundry) - 1
loop

'Para guardar el registro
rst.close



set rst2 = server.CreateObject("ADODB.RECORDSET")
	rst2.CursorLocation = 2
	rst2.CursorType = 0
	rst2.LockType = 3


'if Session("Edit") = 0 then
'end if
'response.Write("Se guardo la info")


'Response.Write ("Archivo subido")
'Response.Write (lngNumberUploaded & " archivo ya está en el servidor.<BR>")
'response.write ("<a href='http://contraloriacolima.ddns.net:8080/Cendi/Expedientes/" & FileName & "'> ver imagen </a><br>")
'Response.Write "<BR><BR><INPUT type='button' onclick='document.location=" & chr(34) & ";saveany.asp;" & chr(34) & "' value='<< Volver' id='button'1 name='button'1>" 

'response.Write(request("txtCampo1"))

'response.Redirect("AdminExpedientes.asp?ID=" & Session("AlumnoActual") & "&FiltroLab=" & Session("CicloEscolar") & "&Def=1")
response.Write("<embed src='../SICAD5/Oficios/TMP/" & Session("NomArc") & "' width='105%' height='800'>")
End Select 
%>


</body>
</html> 
