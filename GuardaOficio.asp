<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SICAD v5.01b</title>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>
<!-- #include file="CSSFileManager.asp" -->

<!-- Para el Responsive -->
<script type="text/javascript">

$(document).ready(function(){
      var height = $(window).height();
      $('#divContenidos').height(height-120);
});
function Resp(){
	$(document).ready(function(){
      var height = $(window).height();
     $('#divContenidos').height(height-120);
});
}

</script>
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
%>

</head>

<body onresize="Resp();">
<!-- #include file="MenuFileManager.asp" -->
<%

if session("NomArc") <> "" then




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
					
	'Para guardar la info en la base de datos
	rst.open "select * from Oficios where Folio=0", cnn
		rst.addnew
			rst.fields("Oficio") = ucase(request("txtNoOficio")) & ""
			rst.fields("Depe") = ucase(request("tags")) & ""
			rst.fields("Departamento") = ucase(request("txtDepartamento")) & ""
			rst.fields("Remitente") = ucase(request("txtRemitente")) & ""
			rst.fields("Fecha1") = date
			rst.fields("Destino") = ucase(request("txtDestino")) & ""
			rst.fields("Asunto") = ucase(request("txtAsunto")) & ""
			rst.fields("AreaInicial") = request("cmbTurnado") 
			rst.fields("Comentario") = request("CheckboxGroup1") & ", " & request("CheckboxGroup2") & ","
			rst.fields("Resuelto") = 0
			rst.fields("AreaInicial") = request("cmbTurnado") 
			rst.fields("ArchivoTMP") = session("NomArc")
			
		rst.update
	rst.close
	
	rst.open "select * from Oficios order by Folio ASC", cnn
	do while not rst.eof
		Ultimo = rst.fields("Folio")
		rst.movenext
	loop
	
	
	rst.close
	
'------------------------------------------------------------------------------------
' by gsus internet art 
' http://www.cedecero.com/gsus
' codigo para su libre utilización
'------------------------------------------------------------------------------------
   
    'Previamente el fichero  Anterior.txt 
    'ha de existir en nuestra carpeta.
	
    'Declaracion de variables
    Dim FSO, Fich , NombreAnterior, NombreNuevo 
    'Inicialización
    NombreAnterior = Session("NomArc")   
    NombreNuevo = "C-" & Ultimo & mid(NombreAnterior, instr(NombreAnterior, "."), 4)
	
    'response.Write(instr(NombreAnterior))
	'response.Write(mid(NombreAnterior, instr(NombreAnterior), 4))   
  ' Instanciamos el objeto
   Set FSO = Server.CreateObject("Scripting.FileSystemObject") 
   ' Asignamos el fichero a renombrar a la variable fich
   Set Fich = FSO.GetFile(Server.MapPath("..\SICAD5\Oficios\TMP\" & NombreAnterior)) 
   ' llamamos a la funcion copiar, 
   'y duplicamos el archivo pero con otro nombre
   Call Fich.Copy(Server.MapPath("..\SICAD5\Oficios\" & NombreNuevo)) 
    ' finalmente borramos el fichero original
   'Call Fich.Delete() 
    
   Set Fich = Nothing 
   Set FSO = Nothing 


	
''	response.Write(request("CheckboxGroup1"))
''	response.Write(request("CheckboxGroup2"))
	
	Session("NomArc") = ""
	'response.Redirect("MainCorrespondencia.asp")
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2" align="center" class="TextoNormal">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center" class="TextoNormal">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td align="center" width="90%">
    <iframe src="ImprimeFolio.asp?ID=<%=Ultimo%>" width="100%" height="500" frameborder="0" ></iframe><%Session("NomArc")=""%>
    </td>
    <td align="center" valign="top"><a class="btn btn-danger" href="BuscaCorr.asp" >cerrar</a></td>
  </tr>
</table>

<% else %>


<table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" class="TextoNormal">&nbsp;</td>
  </tr>
  <tr>
    <td align="center" class="TextoNormal">&nbsp;</td>
  </tr>
  <tr>
    <td align="center" ><br /><h5>Error al generar el folio <br /><br />[0x0014]: No se ha especificado el archivo adjunto</h5></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><INPUT class="btn btn-outline-info" type="button" onclick="history.go(-1)" value="Regresar" id="button" name='button'></td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
</table>
<%end if%>

</body>
</html>
