<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SICAD v5.01b</title>
<link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css">
<link rel="stylesheet" type="text/css" href="SearchFiles/jquery.dataTables.min.css">



<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.3.1/jquery.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js"></script>
<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js"></script>

<style type="text/css" class="init">
	
	</style>
	<!--<script type="text/javascript" language="javascript" src="SearchFiles/jquery-3.3.1.js"></script>-->
	<script type="text/javascript" language="javascript" src="SearchFiles/jquery.dataTables.min.js"></script>
	
	<script type="text/javascript" class="init">
		$(document).ready(function() {
			$('#example').DataTable();
		} );
	</script>
<!-- #include file="CSSFileManager.asp" -->

<!-- Para el PopUP -->
<script language="Javascript">
var calendarWindow = null;
var calendarScreenX = 125; // either 'auto' or numeric
var calendarScreenY = 70; // either 'auto' or numeric 
// }}}
// {{{ getCalendar()
 
function getMin() 
{
	
    if (calendarWindow && !calendarWindow.closed) {
        // alert('Calendar window already open.  Attempting focus...');
        try {
            calendarWindow.focus();
        }
        catch(e) {}
        
        return false;
    }
 
    var cal_width = 1100;
    var cal_height = 550;
 
    // IE needs less space to make this thing
    if ((document.all) && (navigator.userAgent.indexOf("Konqueror") == -1)) {
        cal_width = 410;
    }
 
    
	//calendarTarget = in_dateField;
    calendarWindow = window.open('ShowMin.asp?Param=<%=request("Param")%>', 'dateSelectorPopup','toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=0,dependent=no,width='+cal_width+',height='+cal_height + (calendarScreenX != 'auto' ? ',screenX=' + calendarScreenX : '') + (calendarScreenY != 'auto' ? ',screenY=' + calendarScreenY : ''));
 
    return false;
}
 
// }}}
// {{{ killCalendar()
 
function killCalendar() 
{
    if (calendarWindow && !calendarWindow.closed) {
        calendarWindow.close();
    }
}
 
// }}}
 
</script>

<!-- para los Hiddens-->
<script type="text/javascript"> 
    function ShowMinSinAsignar(ID,Fase) { 
	   var varID = ID;
	   var varFase = Fase;
	   	   
		   $('#ShowMinSinAsignar').modal('show');
		   
		   if (varFase == 0){
			   window.open('AsignaMin.asp?IDMin=' + varID, 'frmSinAsignar');
			   window.open('PaginaVacia.asp', 'frmDocs');
			   }
		   if (varFase == 1){
			   window.open('ShowMin.asp?IDMin=' + varID, 'frmSinAsignar');
			   window.open('AdjuntaMin.asp?IDMin=' + varID, 'frmDocs');
			   }
		   if (varFase == 2){
			   window.open('ShowMin.asp?IDMin=' + varID, 'frmSinAsignar');
			   window.open('waitAdjuntaOK.asp?IDMin=' + varID, 'frmDocs');
			   }
		   if (varFase == 3){
			   window.open('ShowMin.asp?IDMin=' + varID, 'frmSinAsignar');
			   window.open('PaginaVacia.asp', 'frmDocs');
			   }
		   
	  return true; 
    }       
</script>


<!-- Para el Responsive -->
<script type="text/javascript">

$(document).ready(function(){
      var height = $(window).height();
      $('#divContenidos').height(height-120);
	  $('#divEjercicio').height(height-170);

});
function Resp(){
	$(document).ready(function(){
      var height = $(window).height();
     $('#divContenidos').height(height-120);
	 $('#divEjercicio').height(height-170);

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
		
					if request("Anio") = "" then
						Anio = year(date)
					else
						Anio = request("Anio")
					end if
%>

</head>

<body onresize="Resp();">
<!-- #include file="MenuFileManager.asp" -->

<%if Session("SICAD_Active") = 1 then%>

&nbsp;<div class="alert-info" style="padding:5px; height:50px; margin-top:-25px;">
	<div style="float:left; vertical-align:middle; width:30%; padding-left:30px; padding-top:10px;" ><em>Minutario de Oficios</em></div>

	<!--Para buscar -->
	<div align="center" style="width:20%; padding-left:15px; padding-right:15px; float:left; position:relative; padding-top:1px;">&nbsp;</div>

<!-- Para cerrar y ayuda -->
    <div align="right" style="padding-right:15px; width:50%; float:right; padding-top:5px;">
      <a href="#" class="btn btn-sm btn-info" title="Ayuda">?</a> &nbsp;&nbsp;
      <!--<a href="#" class="btn btn-sm btn-info" data-toggle="modal" data-target="#GeneraReportes">reporte</a> &nbsp;&nbsp;-->
      <a href="#" class="btn btn-sm btn-info" data-toggle="modal" data-target="#GeneraOficios">nuevo numero de oficio</a> &nbsp;&nbsp;
      <a href="Main.asp" class="btn btn-sm btn-danger">cerrar</a> &nbsp;&nbsp;
       <!-- <input type="button" onclick="tableToExcel('TestTable', 'CADIDO Global')" value="Exportar CADIDO a Excel" class="btn btn-sm  btn-outline-success"> -->
    </div>
</div>


<div id="divContenidos"  align="left" style="overflow:auto; border-right: #fff 1px solid; border-top: #fff 1px solid; border-left: #fff 1px solid; border-bottom: #fff 1px solid; scrollbar-arrow-color : #999999; scrollbar-face-color : #cccccc; scrollbar-track-color :#3333333; position: relative; left: 0px; top: 0px; width: 100%; float:left;;">

    <!-- Ejercicios -->

    <div align="center" id="divEjercicio"  style="overflow:auto; padding: 15px; border-right: #fff 1px solid; border-top: #fff 1px solid; border-left: #fff 1px solid; border-bottom: #fff 1px solid; scrollbar-arrow-color : #999999; scrollbar-face-color : #cccccc; scrollbar-track-color :#3333333; position: relative; left: 0px; top: 0px; width: 15%; float:left;;">
    
    <div style="background-color:#1d96b2; color:#fff; border-radius: 10px 10px 10px 10px;
-moz-border-radius: 10px 10px 10px 10px;
-webkit-border-radius: 10px 10px 10px 10px;
border: 1px solid #cccccc;" align="center">Ejercicio</div>&nbsp;<br />

<% rst.open "SELECT distinct Anio as C1 from DetalleMinutarios order by anio", cnn 
	
	do while not rst.eof
%> 	
	    <%if cint(Anio) = cint(rst.fields("C1")) then%>
        	<img src="Imagenes/Selected.png" width="20" height="20" />
        <%else%>
        	<img src="Imagenes/UnSelected.png" width="20" height="20" />
        <%end if%>
        <a class="enlaces" href="BuscaMin.asp?Anio=<%=rst.fields("C1")%>" ><%=rst.fields("C1")%></a><br />
    

<%
		rst.movenext
	loop
	rst.close
%>
    
	 </div>
     <!-- FIN Ejercicios -->
	
     <!-- Listado -->
     <div id="ImagenLoading" align="center">
     	<img src="Imagenes/ajax-loader.gif" width="293" height="269" />
     </div>
     <div style="width:85%; float:left; padding:15px; display:none" id="Listado">
     
     <%
	
		' para sacar que minutario le corresponde a este usuario
	
	'rst.open "select * from Minutarios where ResponsableID = " & Session("IDUsuario"), cnn
	rst.open "select * from Minutarios", cnn
		Session("MinutarioID") = rst.fields("idMinutarios")
	rst.close

			
			'rst.open "select * from minutarios inner join detalleminutarios inner join empleados on (idempleados=Solicitante) and (idMinutario=IdMinutarios) and (ResponsableID = " & Session("IDUsuario") & ")", cnn
			if session("NivelMinutario") = 1 then
			rst.open "select * from DetalleMinutarios where idMinutario = " & Session("MinutarioID") & " and anio = " & Anio & " order by idDetalleMinutarios DESC", cnn

			end if
			if Session("NivelMinutario") = 2 then
			rst.open "select * from DetalleMinutarios where idMinutario = " & Session("MinutarioID") & " and AreaSolicitante = " & Session("Area") & " and Anio = " & Anio & " order by idDetalleMinutarios DESC", cnn
			end if
			
			if Session("NivelMinutario") = 3 then
			rst.open "select * from DetalleMinutarios where idMinutario = " & Session("MinutarioID") & " and AreaSolicitante = " & Session("Area") & " and Anio = " & Anio & " and Seguimiento = '" &session("Usuario") & "' order by idDetalleMinutarios DESC", cnn
			end if
			

		%>
				<table id="example" class="table-hover TextoNormal" style="width:100%;">
					<thead>
						<tr>
						  	<th>&nbsp;</th>
                            <th>&nbsp;</th>
							<th>Status</th>
							<th>Número</th>
							<th>Responsable</th>
                            <th>Solicitante</th>
                            <th>Destinatario</th>
						</tr>
					</thead>
					<tbody>
						<tr>
							<%do while not rst.eof
								if datediff("d", rst.fields("Fecha"), date) > 5 and rst.fields("Fase") < 2 then
									Alerta = "ALERTA"
								else
									Alerta = ""
								end if
								Clase = ""
									if rst.fields("Fase") = 0 then
										Clase = "#d1ecf1"
										Texto ="SIN ASIGNAR"
									end if
									if rst.fields("Fase") = 1 then
										Clase = "#ffeeba"
										Texto ="SIN ACUSE"
									end if 
									if rst.fields("Fase") = 2 then
										Clase = "#c3e6cb"
										Texto ="COMPLETO"
									end if 
									if rst.fields("Fase") = 3 then
										Clase = "#f8d7da"
										Texto ="CANCELADO"
									end if                             %>
						    <td bgcolor="<%=Clase%>">&nbsp;</td>
							<td><%=Alerta%></td>
							<td><%=Texto%></a></td>
							
						    <td>
			                            <div class="tooltip1"> <img src="Imagenes/DetailsSearch.png" width="19" height="19" />
                                                <span class="tooltiptext">
                                                    ASUNTO:<br /><%=rst.fields("Asunto")%><BR /><BR />DESTINO:<br /><%=rst.fields("Destinatario")%><br /><%=rst.fields("DepeDestino")%>
                                                </span>
                                            </div>
							<!-- Consecutivo -->
                                <a href="javascript:void(0);" onclick="ShowMinSinAsignar(<%=rst.fields("idDetalleMinutarios")%>,<%=rst.fields("Fase")%>);"><%=right("00000" & rst.fields("Consecutivo"),4)%>/<%=rst.fields("Anio")%></a>
                        <!-- Fin consecutivo -->
                            </td>
							<td>
                            <!-- Responsable -->
                           
							<%
								rst2.open "select * from empleados where idEmpleados = " & rst.fields("Solicitante"), cnn
								if rst2.eof and rst2.bof then
									Persona = ""
								else
									Persona = rst2.fields("Nombre") & chr(32) & rst2.fields("ApellidoP") & chr(32) & rst2.fields("ApellidoM")
								end if
								rst2.close
							%>
							<%=Persona%>
                            <!-- Fin Responsable-->
							</td>
                            <td><%=rst.fields("Seguimiento")%></td>
                            <td>
								<%=mid(rst.fields("Destinatario"),1,50)%>
                            </td>
						</tr>
                        <%
	  						rst.movenext
							loop
					    %>
					</tbody>
				</table>

  </div>
     <!-- FIN Listado -->

</div>

<script>
	setTimeout(function(){ document.getElementById("Listado").style.display="inline"; document.getElementById("ImagenLoading").style.display="none";}, 3000);
</script>

<!-- Para mostrar detalles -->
  <div class="modal fade" id="ShowMinSinAsignar" data-keyboard="false" data-backdrop="static">
    <div class="modal-dialog modal-xl">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">detalles de minutario</h4>
          <!--<button type="button" class="close" data-dismiss="modal">&times;</button>-->
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
          
          <div> <!-- contenedor principal -->
          	
            	<div style="float:left; width:50%; height:400px;"> <!-- Detalles -->
                	<iframe src="" height="390" width="100%" id="frmSinAsignar" name="frmSinAsignar" frameborder="0"  ></iframe>
                </div> <!-- Fin detalles -->
                
                <div style="float:left; width:50%; height:400px;"> <!-- Documento -->
                	<iframe src="" height="390" width="100%" id="frmDocs" name="frmDocs" frameborder="0"  ></iframe>
                </div> <!-- Fin Documento -->
          	
          </div> <!-- Fin conetenedor principal -->

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
        	<button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button> 
        </div>
        
      </div>
    </div>
</div> <!-- Fin Mostrar detalles --> 



  <div class="modal fade" id="GeneraOficios" data-keyboard="false" data-backdrop="static"> <!-- Para generar nuevos folios -->
    <div class="modal-dialog modal-sm">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">Nuevos oficios</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
<form id="form1" name="form1" method="post" action="GeneraFolios.asp">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center">Por favor, indique la cantidad de numeros de oficios a reservar</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center">
      <span class="TextoNormal">
      <select class="form-control" name="cmbCantOf" class="TextoNormal" id="cmbCantOf" >
        <option value="1" selected="selected">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
        <option value="6">6</option>
        <option value="7">7</option>
        <option value="8">8</option>
        <option value="9">9</option>
        <option value="10">10</option>
      </select>
      </span></td>
  </tr>
  <tr>
    <td align="center">&nbsp;</td>
  </tr>
  <tr>
    <td align="center">
      <input class="btn btn-outline-success" type="submit" name="cmdGenerar" id="cmdGenerar" value="Reservar" />
</td>
  </tr>
</table> 
</form>
        
        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button>
        </div>
        
      </div>
    </div>
  </div> <!-- FIN Para generar nuevos folios -->


  <div class="modal fade" id="GeneraReportes" > <!-- Para generar reportes -->
<form id="form2" name="form2" method="post" action="ImprimePendientesMinutario.asp" target="Reporte">
    <div class="modal-dialog modal-xl">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">reportes</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td valign="top" style="padding:15px;" width="30%" align="left">Para generar el reporte de los oficios que presentan una antigüedad mayor a 5 dias y que no cuentan con acuse de recibido, seleccione la persona que lo solicitó
	<%
   rst.close
   
   if session("AdminMinutario") = 1 then
   rst.open "select * from SolicitantesMin inner Join Empleados on (SolicitanteID=idEmpleados)", cnn
   else
   rst.open "select * from SolicitantesMin inner Join Empleados on (SolicitanteID=idEmpleados) and (Area =" & Session("Area") & ")", cnn
   end if
  %><br /><br /><select class="form-control" name="cmbResponsable"  id="cmbResponsable">
        <%do while not rst.eof%>
        <option value="<%=rst.fields("SolicitanteID")%>"><%=rst.fields("Nombre")%>&nbsp;<%=rst.fields("ApellidoP")%>&nbsp;<%=rst.fields("ApellidoM")%></option>
        <%
		  rst.movenext
		  loop
		%>
</select></td>
    <td rowspan="4" style="padding:15px;"><iframe src="" frameborder="0" name="Reporte" id="Reporte" width="100%" height="360" ></iframe></td>
  </tr>
  
  <tr>
    <td valign="top" style="padding:15px;" align="center">
      <span class="TextoNormal">
      <input class="btn btn-outline-success" type="submit" name="cmdGenerar" id="cmdGenerar" value="Generar" />&nbsp;&nbsp;
      </span></td>
  </tr>
  <tr>
    <td style="padding:15px;" align="center">&nbsp;</td>
  </tr>
  <tr>
    <td style="padding:15px;" align="center"><span class="TextoNormal">
      
    </span></td>
  </tr>
</table> 

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          
          <button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button>
        </div>
        
      </div>
    </div>
    </form>
  </div> <!-- FIN Para generar reportes -->



<%end if%>

</body>
</html>
