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
		   
	  return true; 
    }       
</script>


<%if request("Tipo") = 2 then %>
<script>
	window.onload = function Emergente() {
			$('#EditInfo').modal('show');
		}
</script> 
<%end if%>

<!-- Para el Responsive -->
<script type="text/javascript">

$(document).ready(function(){
      var height = $(window).height();
      $('#frm5').height(height-125);
});
function Resp(){
	$(document).ready(function(){
      var height = $(window).height();
     $('#frm5').height(height-125);
});
}

</script>

</head>

<body onresize="Resp();">
<!-- #include file="MenuFileManager.asp" -->

<%if Session("SICAD_Active") = 1 then%>
<br />&nbsp;
<div class="alert-info" style="padding:5px; height:50px; margin-top:-25px;">
	<div style="float:left; vertical-align:middle; width:30%; padding-left:30px; padding-top:10px;" ><em>Vista detallada de las notificaciones</em></div>

	<!--Para buscar -->
	<div align="center" style="width:20%; padding-left:15px; padding-right:15px; float:left; position:relative; padding-top:1px;">&nbsp;</div>

<!-- Para cerrar y ayuda -->
    <div align="right" style="padding-right:15px; width:50%; float:right; padding-top:5px;">
      <a href="#" class="btn btn-sm btn-info" title="Ayuda">?</a> &nbsp;&nbsp;
      <a href="Main.asp" class="btn btn-sm btn-danger">cerrar</a> &nbsp;&nbsp;
       <!-- <input type="button" onclick="tableToExcel('TestTable', 'CADIDO Global')" value="Exportar CADIDO a Excel" class="btn btn-sm  btn-outline-success"> -->
    </div>
</div>

<% Cols = 0
   Cols = datediff("YYYY", "2017-01-01", date) + 1
%>

<%
		set cnn = server.CreateObject("ADODB.CONNECTION")
		set rst = server.CreateObject("ADODB.RECORDSET")
		set rst2 = server.CreateObject("ADODB.RECORDSET")
		set rst3 = server.CreateObject("ADODB.RECORDSET")
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
<br />&nbsp;
<!--<table width="95%" border="0" cellspacing="0" cellpadding="0" align="center" class="table-hover">
  <tr style="border-bottom:#000 1px solid; background-color:#fff;">
    <td><strong>Módulo</strong></td>
    <%for k = 1 to Cols%>
    	<td align="center"><strong><%=k+2016%></strong></td>
    <%next%>
     <td align="center"><strong>Total</strong></td>
  </tr>
  <tr>
  	<td>Oficios <strong>sin atender</strong> en su correspondencia</td>
    <%Suma = 0%>
	<%for k = 1 to Cols%>
    <%rst.open "select count(Oficio) as SinAtender from Oficios where Turnado = '" & session("IDUsuario") & "' and Resuelto = 0 and EnSeguimiento = 0 and Fecha1 like '" & 2016+k & "%'", cnn%>
    	<td align="center"><a href="BuscaCorr.asp?Anio=<%=2016+k%>"><%=rst.fields("SinAtender")%></a></td>
        <%Suma = Suma + CInt(rst.fields("SinAtender"))%>
    <%rst.close%>
    <%next%>
   <td align="center"><strong><%=Suma%></strong></td>
  </tr>
  
  <tr>
  	<td>Oficios <strong>en seguimiento</strong> en su correspondencia</td>
    <%Suma = 0%>
	<%for k = 1 to Cols%>
    <%rst.open "select count(Oficio) as SinAtender from Oficios where Turnado = '" & session("IDUsuario") & "' and Resuelto = 0 and EnSeguimiento = 1 and Fecha1 like '" & 2016+k & "%'", cnn%>
    	<td align="center"><a href="BuscaCorr.asp?Anio=<%=2016+k%>"><%=rst.fields("SinAtender")%></a></td>
        <%Suma = Suma + CInt(rst.fields("SinAtender"))%>
    <%rst.close%>
    <%next%>
    <td align="center"><strong><%=Suma%></strong></td>
  </tr>
  
</table> -->

</div>

<%
'Declaraciones de los tipos de datos
		set cnn = server.CreateObject("ADODB.CONNECTION")
		set rstSeccion = server.CreateObject("ADODB.RECORDSET")
		set rstSerie = server.CreateObject("ADODB.RECORDSET")
		set rstSubSerie = server.CreateObject("ADODB.RECORDSET")
		set rstExpediente = server.CreateObject("ADODB.RECORDSET")
		Archivo = request.ServerVariables("APPL_PHYSICAL_PATH") & "/config.txt"
		set ConFile = createobject ("scripting.filesystemobject")
		set Fichero = ConFile.OpenTextFile(Archivo)
		TextoFichero = Fichero.ReadAll()
						
		Fichero.Close()
		
						
		strConexion = TextoFichero
		cnn.open strConexion
	
		rstSeccion.CursorLocation = 2
		rstSeccion.CursorType = 0
		rstSeccion.LockType = 3
		
		rstSerie.CursorLocation = 2
		rstSerie.CursorType = 0
		rstSerie.LockType = 3
		
		rstSubSerie.CursorLocation = 2
		rstSubSerie.CursorType = 0
		rstSubSerie.LockType = 3

		rstExpediente.CursorLocation = 2
		rstExpediente.CursorType = 0
		rstExpediente.LockType = 3
		
		'rstSeccion.open "SELECT * FROM ar_Seccion",cnn
		
		%>

<!-- Aqui empieza el TreeView -->
<div class="container TextoNormal">
    <div class="panel panel-default">
        <div class="panel-heading" style="background-color:#ff; height:5px;" ><h5>&nbsp;</h5></div>
        <div class="panel-body">
            <!-- TREEVIEW CODE -->
            <ul class="treeview" >
            <%if session("Corr") = 1 then %>
            <li style="border-left:#000 8px solid; border-bottom:#000 1px solid;">
            	<a href="#">Control de Correspondencia</a>
                <ul>
                	<li>
                    	<a href="#">Oficios sin atender</a>
                        <ul>
                            <%
								'Para que los directores vean todo
								if session("NivelCorr") = 2 then
								rst.open "SELECT count(Folio) as Pendientes, Nombre, ApellidoP, ApellidoM, idEmpleados FROM oficios inner join empleados on (Resuelto = 0) and (enSeguimiento = 0) and (Oficios.area = " & session("Area") & ") and (Turnado = idEmpleados) group by Turnado ORDER BY Nombre, ApellidoP, ApellidoP;", cnn
								end if
								if session("NivelCorr") = 3 then
								rst.open "SELECT count(Folio) as Pendientes, Nombre, ApellidoP, ApellidoM, idEmpleados FROM oficios inner join empleados on (Resuelto = 0) and (enSeguimiento = 0) and (Turnado = '" & session("IDUsuario") & "') and (Turnado = idEmpleados) group by Turnado ORDER BY Nombre, ApellidoP, ApellidoP;", cnn
								end if
								
								if rst.eof and rst.bof then
								else
								do while not rst.eof
							%>
                            	<li>
                                <a href="#"><%=rst.fields("Nombre")%>&nbsp;<%=rst.fields("ApellidoP")%>&nbsp;<%=rst.fields("ApellidoM")%>&nbsp;(<%=rst.fields("Pendientes")%>)</a>
                                <ul>
                                	<%
										for k = 1 to Cols
										rst2.open "SELECT count(Folio) as Pendientes, Nombre, ApellidoP, ApellidoM FROM oficios inner join empleados on (Resuelto = 0) and (enSeguimiento = 0) and (Oficios.area = " & session("Area") & ") and (Turnado = idEmpleados) and (Turnado = '" & rst.fields("idEmpleados") & "') and (Fecha1 like '" & 2016+k & "%') group by Turnado;", cnn
									%>
                                    <li>
                                    	<%if rst2.eof and rst2.bof then
										else %>
                                        <a href="#">EJERCICIO <%=2016+k%>&nbsp;(<%=rst2.fields("Pendientes")%>)</a>
                                        <ul>
                                       		<li>
                                            	<%
													rst3.open "select * from Oficios where Turnado = '" & rst.fields("idEmpleados") & "' and resuelto = 0 and enseguimiento = 0 and fecha1 like '" & 2016+k & "%' ",cnn
												
												if rst3.eof and rst3.bof then
												else
												%>
                                                	<table width="61%" border="0" cellspacing="0" cellpadding="0">
                                                    <tr><td>
 													<%do while not rst3.eof%>
                                                    <img src="Imagenes/printPliegoBlank.png" width="50" height="25" />
													<div class="tooltip1"><a href="ShowFolio.asp?Folio=<%=rst3.fields("Folio")%>"><%=right("00000" & rst3.fields("Folio"),5)%></a>
                                                		<span class="tooltiptext" style="line-height:1.2em; font-size:12px;">
                                                    		<%response.Write("Asunto:<br>" & rst3.fields("Asunto") & "<br>&nbsp;")%>	
                                                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                              <tr>
                                                                <td>Fecha de Recepcion</td>
                                                                <td>Remitente</td>
                                                              </tr>
                                                              <tr>
                                                              	<td><%=rst3.fields("Fecha1")%></td>
                                                                <td><%=rst3.fields("Remitente")%></td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><br />Dependencia</td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><%=rst3.fields("Depe")%></td>
                                                              </tr>
                                                            </table>

	                                                	</span>
                                            		</div>
                                                    <%
														rst3.movenext
													  loop
													 rst3.close
													%>
                                                    </td></tr>
                                                    </table>
                                                <%
													end if
												%>
                                                
                                            </li>
                                        </ul>
                                        <%end if%>
                                    </li>
                                      <%rst2.close%>
                                    <%next%>
                                </ul>
                                </li>
                                <%
									rst.movenext
								  loop
								  
								  end if
								  rst.close
								%>
                            </li>
                        </ul>
                    </li>
                    <li>
                    	<a href="#">Oficios en seguimiento</a>
                        <ul>
                            <%
								
								'Para que los directores vean todo
								if session("NivelCorr") = 2 then
								rst.open "SELECT count(Folio) as Pendientes, Nombre, ApellidoP, ApellidoM, idEmpleados FROM oficios inner join empleados on (Resuelto = 0) and (enSeguimiento = 1) and (Oficios.area = " & session("Area") & ") and (Turnado = idEmpleados) group by Turnado ORDER BY Nombre, ApellidoP, ApellidoP;", cnn
								end if
								if session("NivelCorr") = 3 then
								rst.open "SELECT count(Folio) as Pendientes, Nombre, ApellidoP, ApellidoM, idEmpleados FROM oficios inner join empleados on (Resuelto = 0) and (enSeguimiento = 1) and (Turnado = '" & session("IDUsuario") & "') and (Turnado = idEmpleados) group by Turnado ORDER BY Nombre, ApellidoP, ApellidoP;", cnn
								end if
								
								if rst.eof and rst.bof then
								else
								do while not rst.eof
							%>
                            	<li>
                                <a href="#"><%=rst.fields("Nombre")%>&nbsp;<%=rst.fields("ApellidoP")%>&nbsp;<%=rst.fields("ApellidoM")%>&nbsp;(<%=rst.fields("Pendientes")%>)</a>
                                <ul>
                                	<%
										for k = 1 to Cols
										rst2.open "SELECT count(Folio) as Pendientes, Nombre, ApellidoP, ApellidoM FROM oficios inner join empleados on (Resuelto = 0) and (enSeguimiento = 1) and (Oficios.area = " & session("Area") & ") and (Turnado = idEmpleados) and (Turnado = '" & rst.fields("idEmpleados") & "') and (Fecha1 like '" & 2016+k & "%') group by Turnado;", cnn
									%>
                                    
                                    	<%if rst2.eof and rst2.bof then
										else %><li>
                                        <a href="#">EJERCICIO <%=2016+k%>&nbsp;(<%=rst2.fields("Pendientes")%>)</a>
                                        <ul>
                                       		<li>
                                            	<%
													rst3.open "select * from Oficios where Turnado = '" & rst.fields("idEmpleados") & "' and resuelto = 0 and enseguimiento = 1 and fecha1 like '" & 2016+k & "%' ",cnn
												
												if rst3.eof and rst3.bof then
												else
												%>
                                                	<table width="61%" border="0" cellspacing="0" cellpadding="0">
                                                    <tr><td>
 													<%do while not rst3.eof%>
                                                    <img src="Imagenes/printPliegoBlank.png" width="50" height="25" />
													<div class="tooltip1"><a href="ShowFolio.asp?Folio=<%=rst3.fields("Folio")%>"><%=right("00000" & rst3.fields("Folio"),5)%></a>
                                                		<span class="tooltiptext" style="line-height:1.2em; font-size:12px;">
                                                    		<%response.Write("Asunto:<br>" & rst3.fields("Asunto") & "<br>&nbsp;")%>	
                                                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                              <tr>
                                                                <td>Fecha de Recepcion</td>
                                                                <td>Remitente</td>
                                                              </tr>
                                                              <tr>
                                                              	<td><%=rst3.fields("Fecha1")%></td>
                                                                <td><%=rst3.fields("Remitente")%></td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><br />Dependencia</td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><%=rst3.fields("Depe")%></td>
                                                              </tr>
                                                            </table>

	                                                	</span>
                                            		</div>
                                                    <%
														rst3.movenext
													  loop
													 rst3.close
													%>
                                                    </td></tr>
                                                    </table>
                                                <%
													end if
												%>
                                                
                                            </li>
                                        </ul> </li>
                                        <%end if%>
                                   
                                      <%rst2.close%>
                                    <%next%>
                                </ul>
                                </li>
                                <%
									rst.movenext
								  loop
								 
								  end if 
								  rst.close
								%>
                            </li>
                        </ul>
                    </li>
                </ul>
            </li>
            <%end if%> <!-- Fin del NivelCorr > 1 -->
            <%if Session("NivelMinutario") > 1 then %>
            <div style="height:30px" ></div>
            <li style="border-left:#090 8px solid; border-bottom:#090 1px solid;">
            	<a href="#"> Minutario de Oficios</a>
                <ul>
                	<li>
                    	<a href="#">Oficios reservados sin asignar</a>
                           <ul>
                        	<%
								if session("NivelMinutario") = 2 then
									rst.open "SELECT count(idDetalleMinutarios) as Pendientes, Seguimiento FROM DetalleMinutarios where (Fase = 0) and (AreaSolicitante = " & session("Area") & ") group by Seguimiento ORDER BY Seguimiento", cnn
								end if
								if session("NivelMinutario") = 3 then
									rst.open "SELECT count(idDetalleMinutarios) as Pendientes, Seguimiento, idEmpleados FROM DetalleMinutarios inner join empleados on (Fase = 0) and (Seguimiento = '" & session("Usuario") & "') and (Solicitante = idEmpleados) group by Seguimiento ORDER BY Seguimiento", cnn
								end if
							%>
                            <%
							if rst.eof and rst.bof then
							else
							do while not rst.eof%>
                            <li>
                            	<a href="#"><%=rst.fields("Seguimiento")%>&nbsp;(<%=rst.fields("Pendientes")%>)</a>
                                <ul>
                                	    <%
											for k = 1 to Cols
											rst2.open "SELECT COUNT(idDetalleMinutarios) as Pendientes FROM detalleminutarios d where areaSolicitante = " & session("Area") & " and fase = 0 and Seguimiento = '" & rst.fields("Seguimiento") & "' and Fecha like '" & 2016+k & "%';", cnn
										%>
                                    
                                       <%if cint(rst2.fields("Pendientes")) = 0 then
										else %>
                                        <li>
                                        	<a href="#">EJERCICIO <%=2016+k%>&nbsp;(<%=rst2.fields("Pendientes")%>)</a>
                                            <ul>
                                            	<li>
                                            	<%
													rst3.open "select * from detalleminutarios where seguimiento = '" & rst.fields("Seguimiento") & "' and fase = 0 and fecha like '" & 2016+k & "%';",cnn
												
												if rst3.eof and rst3.bof then
												else
												%>
                                                	<table width="61%" border="0" cellspacing="0" cellpadding="0">
                                                    <tr><td>
 													<%do while not rst3.eof%>
                                                    <img src="Imagenes/printPliegoBlank.png" width="50" height="25" />
													<div class="tooltip1"><a href="javascript:void(0);" onclick="ShowMinSinAsignar(<%=rst3.fields("idDetalleMinutarios")%>,<%=rst3.fields("Fase")%>);"><%=right("0000" & rst3.fields("Consecutivo"),4)%>/<%=rst3.fields("Anio")%></a>
                                                		<span class="tooltiptext" style="line-height:1.2em; font-size:12px;">
                                                    		<%response.Write("Asunto:<br>" & rst3.fields("Asunto") & "<br>&nbsp;")%>	
                                                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                              <tr>
                                                                <td>Fecha de Solicitud</td>
                                                                <td>Destinatario</td>
                                                              </tr>
                                                              <tr>
                                                              	<td><%=rst3.fields("Fecha")%></td>
                                                                <td><%=rst3.fields("Destinatario")%></td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><br />Dependencia</td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><%=rst3.fields("DepeDestino")%></td>
                                                              </tr>
                                                            </table>

	                                                	</span>
                                            		</div>
                                                    <%
														rst3.movenext
													  loop
													 rst3.close
													%>
                                                    </td></tr>
                                                    </table>
                                                <%
													end if
												%>
                                                
                                            </li>
                                            </ul>
                                        </li>
                                        <%end if%>
                                    
                                    <% 
											rst2.close
											next 
									%>
                                </ul>
                            </li>
                            <%
								rst.movenext
							loop
							end if
							rst.close
							
							%>
                        </ul>

                    </li>
                    <li>
                    	<a href="#">Oficios sin acuse</a>
                        <ul>
                        	<%
								if session("NivelMinutario") = 2 then
									rst.open "SELECT count(idDetalleMinutarios) as Pendientes, Seguimiento, idEmpleados FROM DetalleMinutarios inner join empleados on (Fase = 1) and (AreaSolicitante = " & session("Area") & ") and (Solicitante = idEmpleados) group by Seguimiento ORDER BY Seguimiento", cnn
								end if
								if session("NivelMinutario") = 3 then
									rst.open "SELECT count(idDetalleMinutarios) as Pendientes, Seguimiento, idEmpleados FROM DetalleMinutarios inner join empleados on (Fase = 1) and (Seguimiento = '" & session("Usuario") & "') and (Solicitante = idEmpleados) group by Seguimiento ORDER BY Seguimiento", cnn
								end if
							%>
                            <%
							if rst.eof and rst.bof then
							else
							do while not rst.eof%>
                            <li>
                            	<a href="#"><%=rst.fields("Seguimiento")%>&nbsp;(<%=rst.fields("Pendientes")%>)</a>
                                <ul>
                                	    <%
											for k = 1 to Cols
											rst2.open "SELECT COUNT(idDetalleMinutarios) as Pendientes FROM detalleminutarios d where areaSolicitante = " & session("Area") & " and fase = 1 and Seguimiento = '" & rst.fields("Seguimiento") & "' and Fecha like '" & 2016+k & "%';", cnn
										%>
                                    
                                       <%if cint(rst2.fields("Pendientes")) = 0 then
										else %>
                                        <li>
                                        	<a href="#">EJERCICIO <%=2016+k%>&nbsp;(<%=rst2.fields("Pendientes")%>)</a>
                                            <ul>
                                            	<li>
                                            	<%
													rst3.open "select * from detalleminutarios where seguimiento = '" & rst.fields("Seguimiento") & "' and fase = 1 and fecha like '" & 2016+k & "%';",cnn
												
												if rst3.eof and rst3.bof then
												else
												%>
                                                	<table width="61%" border="0" cellspacing="0" cellpadding="0">
                                                    <tr><td>
 													<%do while not rst3.eof%>
                                                    <img src="Imagenes/printPliegoBlank.png" width="50" height="25" />
													<div class="tooltip1"><a href="javascript:void(0);" onclick="ShowMinSinAsignar(<%=rst3.fields("idDetalleMinutarios")%>,<%=rst3.fields("Fase")%>);"><%=right("0000" & rst3.fields("Consecutivo"),4)%>/<%=rst3.fields("Anio")%></a>
                                                		<span class="tooltiptext" style="line-height:1.2em; font-size:12px;">
                                                    		<%response.Write("Asunto:<br>" & rst3.fields("Asunto") & "<br>&nbsp;")%>	
                                                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                              <tr>
                                                                <td>Fecha de Solicitud</td>
                                                                <td>Destinatario</td>
                                                              </tr>
                                                              <tr>
                                                              	<td><%=rst3.fields("Fecha")%></td>
                                                                <td><%=rst3.fields("Destinatario")%></td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><br />Dependencia</td>
                                                              </tr>
                                                              <tr>
                                                              	<td colspan="2"><%=rst3.fields("DepeDestino")%></td>
                                                              </tr>
                                                            </table>

	                                                	</span>
                                            		</div>
                                                    <%
														rst3.movenext
													  loop
													 rst3.close
													%>
                                                    </td></tr>
                                                    </table>
                                                <%
													end if
												%>
                                                
                                            </li>
                                            </ul>
                                        </li>
                                        <%end if%>
                                    
                                    <% 
											rst2.close
											next 
									%>
                                </ul>
                            </li>
                            <%
								rst.movenext
							loop
							end if
							rst.close
							
							%>
                        </ul>
                        
                    </li>
                </ul>
            </li>
            <%end if%> <!-- Fin NivelMinutario > 1 -->
            <%if Session("NivelPliego") < 4 then%>
            <div style="height:30px" ></div>
            <li style="border-left:#900 8px solid; border-bottom:#900 1px solid;">
            	<a href="#">Pliegos y Permisos</a>
                <ul>
                	<li>
                    	<a href="#">Autorizaciones pendientes</a>
                        <ul>
                        	---[ EN PROCESO DE ACTUALIZACION, La información estará disponible en breve ]---
                        </ul>
                    </li>
                    <li>
                    	<a href="#">Envios pendientes a Capital Humano</a>
                        <ul>
                        	---[ EN PROCESO DE ACTUALIZACION, La información estará disponible en breve ]---
                        </ul>
                    </li>
                </ul>
            </li>
            <%end if%> <!-- Fin NivelPliego < 4 -->
            </ul>
            <!-- TREEVIEW CODE -->
            
        </div>
    </div>
<!-- Aqui finaliza el TrewView-->
<!-- #include file="JSFileManager.asp" -->



<!-- Para agregar informe -->
  <div class="modal fade" id="NewInfo" data-keyboard="false" data-backdrop="static">
    <div class="modal-dialog modal-xl">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">nuevo documento</h4>
          <!--<button type="button" class="close" data-dismiss="modal">&times;</button>-->
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
          
          <iframe src="NewNormateca.asp" frameborder="0" width="100%" height="380" ></iframe>

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
        	<button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button> 
        </div>
        
      </div>
    </div>
</div> <!-- Fin agregar informe --> 


<!-- Para Editar -->
  <div class="modal fade" id="EditInfo" data-keyboard="false" data-backdrop="static">
    <div class="modal-dialog modal-xl">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">editar documento</h4>
          <!--<button type="button" class="close" data-dismiss="modal">&times;</button>-->
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
          
          <iframe src="NewNormateca.asp?Origen=Edit&IDDoc=<%=request("IDDoc")%>&TituloDoc=<%=request("TituloDoc")%>&idPadre=<%=request("idPadre")%>&idHijo=<%=request("idHijo")%>&idNieto=<%=request("idNieto")%>&FechaPub=<%=request("Fechaüb")%>&RutaDoc=<%=request("RutaDoc")%>&FuentePub=<%=request("FuentePub")%>" frameborder="0" width="100%" height="380" ></iframe>

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
        	<a href="MainNormateca.asp" class="btn btn-danger" target="_top">cerrar</a>
         <!--   <button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button> -->
        </div>
        
      </div>
    </div>
</div> <!-- Fin Editar --> 


<%end if%>

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


</body>

</html>