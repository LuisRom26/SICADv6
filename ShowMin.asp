<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CharSet = "utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />-->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/> 
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>
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

rst.open "select * from detalleminutarios inner Join Minutarios on (idMinutarios=idMinutario) and (idDetalleMinutarios = " & request("IDMin") & ")", cnn

rst2.open "select * from Empleados where idEmpleados = " & rst.fields("Solicitante"), cnn

%>

<form id="form1" name="form1" method="post" action="GuardaMin.asp?IDMin=<%=request("IDMin")%>">
<table width="100%" border="0" cellpadding="0" cellspacing="0" >
  <tr>
    <td height="40" align="center"><h4><%=rst.fields("NumOficio")%></h4></td>
  </tr>
  <tr>
    <td><em><strong>Numero de Oficio</strong></em></td>
  </tr>
  <tr>
    <td>
      <label for="txtOficio"><%=rst.fields("NumOficio")%></label></td>
  </tr>
  <tr>
    <td><em><strong>Fecha de solicitud</strong></td>
  </tr>
  <tr>
    <td><label for="txtFecha"><%=formatdatetime(rst.fields("Fecha"),1)%></label></td>
  </tr>
  <tr>
    <td><em><strong>Destinatario</strong></td>
  </tr>
  <tr>
    <td><label for="txtDestinatario"><%=rst.fields("Destinatario")%></label></td>
  </tr>
  <tr>
    <td><em><strong>Dependencia</strong></td>
  </tr>
  <tr>
    <td><label for="txtDependencia"><%=rst.fields("DepeDestino")%></label></td>
  </tr>
  <tr>
    <td><em><strong>Responsable de área</strong></td>
  </tr>
  <tr>
    <td><label for="cmbSolicito"><%=rst2.fields("Nombre")%>&nbsp;<%=rst2.fields("ApellidoP")%>&nbsp;<%=rst2.fields("ApellidoM")%></label></td>
  </tr>
  <tr>
    <td><em><strong>Solicitante</strong></em></td>
  </tr>
  <tr>
    <td><%=rst.fields("Seguimiento")%></td>
  </tr>
  <tr>
    <td><em><strong>Asunto</strong></td>
  </tr>
  <tr>
    <td><label for="txtRespuesta12"><%=rst.fields("Asunto")%></label></td>
  </tr>
  <tr>
    <td><em><strong>Respuesta a oficio</strong></td>
  </tr>
  <tr>
    <td><label for="txtRespuesta"><%=rst.fields("RespuestaA")%></label></td>
  </tr>
  <%if rst.fields("Fase") = 3 then%>
  <tr>
    <td><em><strong>Fecha y motivo de cancelación</strong></td>
  </tr>
  <tr>
    <td><label for="txtMotivo"><%=rst.fields("FechaCancelacion")%>:&nbsp;<%=rst.fields("MotivoCancelacion")%></label></td>
  </tr>
  <%end if%>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center">
    <%if rst.fields("Fase") < 2 then%>
    	<% if datediff("d", rst.fields("fecha"), date) > 5 then%>
            <a data-toggle="modal" data-target="#EdicionCancelada" style="cursor:pointer" class="btn btn-outline-info">editar</a>
        <%else%>
        	<a href="AsignaMin.asp?IDMin=<%=request("IDMin")%>" class="btn btn-outline-info">editar</a>&nbsp;&nbsp;
        <%end if%>
		
        <a data-toggle="modal" data-target="#CancelaOficio" style="cursor:pointer" class="btn btn-outline-danger">cancelar oficio</a>    
    <%end if%>
    </td>
    
  </tr>
</table>
</form>

<!-- Para cancelar oficios -->
  <div class="modal fade" id="CancelaOficio">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">Pregunta</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        	¿Esta usted seguro de cancelar el oficio <%=rst.fields("NumOficio")%>?<br /><br />Una vez cancelado, este número de oficio no podrá ser utilizado nuevamente. <br /><br />
            <a data-toggle="modal" data-target="#ComentarioCancelacion" style="cursor:pointer" class="btn btn-outline-success">si, cancelar</a>    &nbsp;&nbsp;&nbsp;<a class="btn btn-outline-danger" data-dismiss="modal">no, regresar</a></td>
        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button>
        </div>
        
      </div>
    </div>
  </div>


<!-- Para decir que ya no se puede hacer nada al oficio -->
  <div class="modal fade" id="EdicionCancelada">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">Aviso</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        	Despues de 5 dias naturales, no se pueden realizar cambios en la asignación de oficios.
        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button>
        </div>
        
      </div>
    </div>
  </div>


<!-- Para poner el comentario del motivo de cancelacion de un oficio -->
  <div class="modal fade" id="ComentarioCancelacion">
  <form id="formComentario" name="formComentario" method="post" action="CancelaMin.asp?IDMin=<%=request("IDMin")%>">
    <div class="modal-sm modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">comentario</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>motivo de cancelacion</td>
  </tr>
  <tr>
    <td><label for="txtComentario"></label>
      <textarea class="form-control" name="txtComentario" id="txtComentario" rows="5"></textarea>
  </tr>
</table>


        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <input type="submit" name="cmdEnviar" id="cmdEnviar" value="Cancelar oficio" class="btn btn-outline-success" />
          <button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">regresar</span></button>
        </div>
        
      </div>
    </div>
    </form>
  </div>


</body>
</html>
