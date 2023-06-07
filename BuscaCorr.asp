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
	<script type="text/javascript" language="javascript" src="SearchFiles/jquery-3.3.1.js"></script>
	<script type="text/javascript" language="javascript" src="SearchFiles/jquery.dataTables.min.js"></script>
	
	<script type="text/javascript" class="init">
		$(document).ready(function() {
			$('#example').DataTable();
		} );
	</script>
<!-- #include file="CSSFileManager.asp" -->

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
	<div style="float:left; vertical-align:middle; width:70%; padding-left:30px; padding-top:10px;" ><em>Revisar correspondencia</em></div>

	<!--Para buscar -->
	<div align="center" style="width:20%; padding-left:15px; padding-right:15px; float:left; position:relative; padding-top:1px;">&nbsp;</div>

<!-- Para cerrar y ayuda -->
    <div align="right" style="padding-right:15px; width:10%; float:right; padding-top:5px;">
        <a href="#" class="btn btn-sm btn-info" title="Ayuda">?</a> &nbsp;&nbsp;
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

<% rst.open "SELECT distinct fecha1 as C1 from oficios group by mid(fecha1,1,4)", cnn 
	do while not rst.eof
%> 	
	    <%if mid(Anio,1,4) = mid(rst.fields("C1"), 7,4) then%>
        	<img src="Imagenes/Selected.png" width="20" height="20" />
        <%else%>
        	<img src="Imagenes/UnSelected.png" width="20" height="20" />
        <%end if%>
        <a class="enlaces" href="BuscaCorr.asp?Anio=<%=mid(rst.fields("C1"), 7,4)%>" ><%=mid(rst.fields("C1"), 7,4)%></a><br />
    

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
     <div style="width:85%; float:right; padding:15px; display:none" id="Listado">
     
     <%
		ConsultaSQL = "SELECT * FROM Oficios "
		WhereSQL = ""
		CadComplete = false
		OrderBy = "ORDER BY Folio DESC"
		
		if Session("NivelCorr") = 3 then
			WhereSQL = "WHERE Turnado = '" & UCase(Session("idUsuario")) & "' and Fecha1 like '" & Anio & "%' "  'Para usuario restringido
		end if
		
		if Session("NivelCorr") = 2 then
			WhereSQL = "WHERE area = " & Session("Area") & " and Fecha1 like '" & Anio & "%' "
		end if
		
		if Session("NivelCorr") = 1 then
			WhereSQL = "WHERE Fecha1 like '" & Anio & "%' "  'Para usuario restringido
		end if
		
		'response.Write(ConsultaSQL & WhereSQL & OrderBy)
		rst.open ConsultaSQL & WhereSQL & OrderBy, cnn
	 %>
     

				<table id="example" class="table-hover TextoNormal" style="width:100%;">
					<thead>
						<tr>
						  <th>&nbsp;</th>
							<th>Status</th>
							<th>Folio</th>
							<th>Remitente</th>
							<th>Responsable</th>
                            <th>Fecha de registro</th>
                            
						</tr>
					</thead>
					<tbody>
						<tr>
							<%do while not rst.eof
                                    if rst.fields("Resuelto") = 1 then
                                        Texto = "ATENDIDO"
                                        Clase = "#c3e6cb"
                                        Estilo = ""
                                    else
                                        Texto = "SIN ATENDER"
                                        Clase = "#ffeeba"
                            
                                            if rst.fields("EnSeguimiento") = 1 then
                                                Texto = "EN SEGUIMIENTO"
                                                Clase = "#d1ecf1"
                                            end if
                                    end if
                            %>
						  <td bgcolor="<%=Clase%>">&nbsp;</td>
							<td><%=Texto%></td>
							<td>
                                <div class="tooltip1"> <img src="Imagenes/DetailsSearch.png" width="19" height="19" />
                                    <span class="tooltiptext">
                                    	<%=rst.fields("oficio")%>: <%=rst.fields("Asunto")%>	
                                    </span>
                                </div>
                            	<a href="ShowFolio.asp?Folio=<%=rst.fields("Folio")%>"><%=right("00000" & rst.fields("Folio"),5)%></a>
                            </td>
							
						  <td><%=rst.fields("Remitente")%></td>
							<td>	<%if rst.fields("Turnado") = "" then %>
        								SIN ASIGNAR
        							<%else%>
   							  <%
										rst2.open "select * from empleados where idempleados = " & rst.fields("Turnado"), cnn
										if rst2.eof and rst2.bof then
										else
											response.Write(rst2.fields("Nombre") & " " & rst2.fields("ApellidoP") & " " & rst2.fields("ApellidoM"))
										end if
										rst2.close
									%>
       								<%end if%></td>
                            <td><%=rst.fields("Fecha1")%></td>
                          
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

<%end if%>

</body>
</html>
