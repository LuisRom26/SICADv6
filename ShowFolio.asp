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
      $('#divContenidos').height(height-135);
	  $('#divDatos').height(height-185);
	  $('#DocAdjunto').height(height-225);
});
function Resp(){
	$(document).ready(function(){
      var height = $(window).height();
     $('#divContenidos').height(height-135);
	 $('#divDatos').height(height-185);
	 $('#DocAdjunto').height(height-225);
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

<%if Session("SICAD_Active") = 1 then%>

&nbsp;<div class="alert-info" style="padding:5px; height:50px; margin-top:-25px;">
	<div style="float:left; vertical-align:middle; width:70%; padding-left:30px; padding-top:10px;" ><em>Detalles del folio número <%=request("Folio")%></em></div>

	<!--Para buscar -->
	<div align="center" style="width:5%; padding-left:15px; padding-right:15px; float:left; position:relative; padding-top:1px;">&nbsp;</div>

<!-- Para cerrar y ayuda -->
    <div align="right" style="padding-right:15px; width:25%; float:right; padding-top:5px;">
        <a href="#" class="btn btn-sm btn-info" title="Ayuda">?</a> &nbsp;&nbsp;
      <a href="BuscaCorr.asp" class="btn btn-sm btn-danger"><< regresar al listado</a> &nbsp;&nbsp;
       <!-- <input type="button" onclick="tableToExcel('TestTable', 'CADIDO Global')" value="Exportar CADIDO a Excel" class="btn btn-sm  btn-outline-success"> -->
    </div>
</div>


<div id="divContenidos"  align="left" style="overflow:auto; padding-right: 15px; padding-top: 15px; padding-left: 0px; padding-bottom: 0px; border-right: #fff 1px solid; border-top: #fff 1px solid; border-left: #fff 1px solid; border-bottom: #fff 1px solid; scrollbar-arrow-color : #999999; scrollbar-face-color : #cccccc; scrollbar-track-color :#3333333; position: relative; left: 0px; top: 0px; width: 100%; float:left;;">

    <!-- detalle del folio -->
    <%
		rst.open "SELECT * FROM Oficios  inner Join empleados on (Folio = " & request("Folio") & ") and (Turnado=idEmpleados) ", cnn
			if rst.eof and rst.bof then
				rst.close
				rst.open "select * from Oficios where Folio = " & request("Folio"), cnn
				Session("FolioResponder") = rst.fields("Folio")
			end if
			Session("FolioFW") = request("Folio")
	%>
    <div id="divDatos"  align="left" style="overflow:auto; padding: 15px; border-right: #fff 1px solid; border-top: #fff 1px solid; border-left: #fff 1px solid; border-bottom: #fff 1px solid; scrollbar-arrow-color : #999999; scrollbar-face-color : #cccccc; scrollbar-track-color :#3333333; position: relative; left: 0px; top: 0px; width: 40%; float:left;;">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr>
	    <td width="14%"><img src="Imagenes/Mail.png" width="79" height="79" /></td>
	    <td class="TextoNormal16" width="86%" style="padding-left:15px;">F-<%=right("00000" & rst.fields("Folio"),5)%><br /><%=rst.fields("Oficio")%><br /><span class="TextoNormal text-capitalize"><%=lcase(rst.fields("Remitente"))%><br /></span><span class="TextoNormal"><%=lcase(rst.fields("depe"))%></span></td>
      </tr>
	  <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2"><table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
	      <tr>
	        <td>
            <a href="ShowDocumentEmbed.asp?w=630&h=470&Doc=../oficios/C-<%=rst.fields("Folio") & mid(rst.fields("ArchivoTMP"), instr(rst.fields("ArchivoTMP"), "."), 4) %>" target="DocAdjunto"><img src="Imagenes/Download.png" width="35" height="35" /></a> <br />C-<%=rst.fields("Folio")%>.pdf</td>
	        <td><a href="ImprimeFolio.asp?ID=<%=rst.fields("Folio")%>" target="DocAdjunto"><img src="Imagenes/Download.png" width="35" height="35" /></a><br />Caratula.pdf</td>

            <td valign="bottom" align="center">
            	<%if rst.fields("EnSeguimiento") = 0 and rst.fields("Resuelto") = 0 then  %>
                <img src="Imagenes/Forward.png" width="35" height="35" data-toggle="modal" data-target="#Forward" style="cursor:pointer" /><br />Reasignar<%end if%></td>
          </tr>
        </table></td>
      </tr>
	  <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2"><%=formatdatetime(rst.fields("Fecha1"),1)%></td>
      </tr>
	  <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2"><div style="padding-left:30px"><%=rst.fields("Asunto")%></div></td>
      </tr>
	  <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2"><em><strong>Comentarios adicionales de quien turna oficio</strong></em></td>
      </tr>
	  <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2"><div style="padding-left:30px;"><div style="padding-left:10px; border-left:#ccc 1px solid;"><%=rst.fields("FechaReasigna")%> - <%=rst.fields("QuienReasigna")%>:<br /><br /><%=rst.fields("ComReasigna")%><br /><br />&nbsp;</div></div></td>
      </tr>
      <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2">
        	<%if rst.fields("Resuelto")=0 then%>
            <%else%>
            <%end if%>
        </td>
      </tr>
      <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
	  <tr>
	    <td colspan="2">
			<%if rst.fields("Turnado") <> "" then%>
					<%if rst.fields("Resuelto") = 0 then %>
                    
                    <%end if%>
                    
                    <%if rst.fields("EnSeguimiento") = 1 then%>
                    	 <em><strong>Comentario de seguimiento</strong></em><br /><br />
                         <div style="padding-left:30px"><div style="padding-left:10px; border-left:#ccc 1px solid;"><%=rst.fields("FechaSeguimiento")%> - <%=rst.fields("DioSeguimiento")%><%'=rst.fields("Nombre")%>&nbsp;<%'=rst.fields("ApellidoP")%>&nbsp;<%'=rst.fields("ApellidoM")%>:<br /><%=rst.fields("Seguimiento")%></div></div>
                         <br />&nbsp;
                    <%end if%>
					
					<%if rst.fields("Resuelto") = 1 then%>
                        <em><strong>Atención realizada al folio</strong></em><br /><br />
                        <div style="padding-left:30px"><div style="padding-left:10px; border-left:#ccc 1px solid;">
                            <%=rst.fields("Fecha2")%> - <%=rst.fields("atendio")%><%'=rst.fields("Nombre")%>&nbsp;<%'=rst.fields("ApellidoP")%>&nbsp;<%'=rst.fields("ApellidoM")%>:<br /><br />
                        <%=trim(rst.fields("Concepto"))%></div></div>
                    <%end if%>
                    
                    
			<%end if%>
                 </td>
      </tr>
      <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
	    <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
	    <td align="right" colspan="2">
		<%if rst.fields("Resuelto") = 0 and rst.fields("Turnado") <> "" then%>
                <%if rst.fields("EnSeguimiento") <> 1 then%>
                	<a class="btn btn-outline-info" data-toggle="modal" data-target="#Seguimiento">iniciar seguimiento</a>&nbsp;&nbsp;
                <%end if%>
                <a class="btn btn-outline-success" data-toggle="modal" data-target="#Responder">finalizar</a> 
                <!--<input class="btn btn-outline-info" type="submit" name="cmdResponder" id="cmdResponder" value="Guardar comentarios" /> -->
        <%else%>
        	<%if rst.fields("Turnado") = "" then%>
            	<a href="EditCorr.asp?Folio=<%=request("Folio")%>" class="btn btn-outline-info">editar carátula</a>
            <%end if%>
        <%end if%>

        <!-- Para que el administrador de correspondencia pueda volver a habilitar un folio -->
        <%if rst.fields("Resuelto") = 1 and Session("NivelCorr") = 1 then %>
        	<a data-toggle="modal" data-target="#HabilitaFolio" style="cursor:pointer" class="btn btn-outline-info">habilitar nuevamente</a>
        <%end if%>

        &nbsp;&nbsp;<a class="btn  btn-danger" href="BuscaCorr.asp" role="button" >cerrar</a></td>
      </tr>
    </table>    
	 </div>
     <!-- FIN Detalle del Folio -->
	
     <!-- Dato adjunto -->
     <div style="width:60%; float:right; padding:15px;">
	     
         <iframe src="" width="100%" name="DocAdjunto" frameborder="0" id="DocAdjunto" ></iframe>
      </div>
       	<script language="Javascript">
			window.open('ShowDocumentEmbed.asp?w=630&h=520&Doc=../oficios/C-<%=rst.fields("Folio") & mid(rst.fields("ArchivoTMP"), instr(rst.fields("ArchivoTMP"), "."), 4) %>', 'DocAdjunto');
		</script>
     <!-- FIN Dato Adjunto -->

</div>

<!-- Modal para reasignar folios -->
<div class="modal fade" id="Forward" data-keyboard="false" data-backdrop="static">
<form id="form2" name="form2" method="post" action="Reasignar.asp">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">reasignar a otro usuario</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        
        <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td width="50%">Nuevo responsable</td>
            <td width="50%">&nbsp;<%
   	rst.close
		if session("nivelCorr") = 1 then
			
			 impresion = 1
			rst.open "Select * from Empleados where Activo = 1  and Corr = 1 and NivelCorr <= 2 order by nombre", cnn
		end if
		if session("NivelCorr") = 2 or session("NivelCorr") = 3 then
			
			rst.open "SELECT * FROM Empleados Where Activo = 1 and  Area =  " & session("Area") & " and (nivelCorr = 3 or nivelCorr = 2)"   , cnn
		end if
				%>
            </td>
          </tr>
          
          
		  <tr>
            <td colspan="2">
              <label for="cmdUnidades"></label>
              <select class="form-control" name="cmdUnidades" id="cmdUnidades">
              <% do until rst.eof %>
                <option value="<%=ucase(rst.fields("idEmpleados"))%>"><%=ucase(rst.fields("Nombre"))%>&nbsp;<%=rst.fields("ApellidoP")%>&nbsp;<%=rst.fields("ApellidoM")%></option>
              <%   rst.movenext
			     loop
				 
				rst.close
				 
			  %>
              </select>
            </td>
            
          </tr>
		  
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="2">Agregar comentario al nuevo destinatario</td>

          </tr>
          <tr>
            <td colspan="2"><label for="txtComentario"></label>
            <textarea class="form-control" name="txtComentario" cols="51" rows="5" class="TextoNormal" id="txtComentario"></textarea></td>
          </tr>
        </table>


        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <input type="submit" name="cmdAceptar" id="cmdAceptar" value="Reasignar" class="btn btn-outline-info" />&nbsp;&nbsp;
          <button type="button" class="btn btn-danger" data-dismiss="modal">cerrar</button>
        </div>
        
      </div>
    </div>
    </form>
  </div>
  
<!-- Modal para responder folios -->
<div class="modal fade" id="Responder" data-keyboard="false" data-backdrop="static">
<form id="form1" name="form1" method="post" action="Responder.asp?Folio=<%=request("Folio")%>">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">resolución final</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
  			<tr>
   				<td>A continuación, favor de redactar de la forma mas explicita posible el resolutivo al asunto en cuestión.<br />&nbsp;</td>
  			</tr>
  			<tr>
  				<td height="100"><textarea onkeypress="return pulsar(event)" class="form-control" name="txtRespuesta" cols="51" rows="8" id="txtRespuesta"></textarea></td>
  			</tr>
		</table>
        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <input type="submit" name="cmdAceptar" id="cmdAceptar" value="guardar comentario" class="btn btn-outline-info" />&nbsp;&nbsp;
          <button type="button" class="btn btn-danger" data-dismiss="modal">cerrar</button>
        </div>
        
      </div>
    </div>
    </form>
  </div>

<!-- Modal para iniciar un seguimiento -->
<div class="modal fade" id="Seguimiento" data-keyboard="false" data-backdrop="static">
<form id="formSeguimiento" name="formSeguimiento" method="post" action="Seguimiento.asp?Folio=<%=request("Folio")%>">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">comentario de seguimiento</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
  			<tr>
   				<td>A continuación, favor de redactar de la forma mas explicita posible el comentario para iniciar el seguimiento de este asunto<br />&nbsp;</td>
  			</tr>
  			<tr>
  				<td height="100"><textarea onkeypress="return pulsar(event)" class="form-control" name="txtSeguimiento" cols="51" rows="8" id="txtSeguimiento"></textarea></td>
  			</tr>
		</table>
        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <input type="submit" name="cmdAceptar" id="cmdAceptar" value="guardar comentario" class="btn btn-outline-info" />&nbsp;&nbsp;
          <button type="button" class="btn btn-danger" data-dismiss="modal">cerrar</button>
        </div>
        
      </div>
    </div>
    </form>
  </div>

 
<!-- Modal para habilitar folios -->
<div class="modal fade" id="HabilitaFolio" data-keyboard="false" data-backdrop="static">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">pregunta</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        ¿Está seguro de habilitar nuevamente el folio <strong><%=request("Folio")%></strong>?<br /><br /><em>Cuando realice esta acción, la información de seguimiento capturada previamente será eliminada. Esta acción no se puede revertir.</em>
        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <a href="HabilitaFolio.asp?Folio=<%=request("Folio")%>" class=" btn btn-outline-success">Si, habilitar</a>&nbsp;&nbsp;
          <button type="button" class="btn btn-danger" data-dismiss="modal">No, regresar</button>
        </div>
      </div>
    </div>
  </div>
 


<%end if%>

</body>
</html>
