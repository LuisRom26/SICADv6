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
	  $('#divDatos').height(height-170);
	  $('#OficioNuevo').height(height-220);
});
function Resp(){
	$(document).ready(function(){
      var height = $(window).height();
     $('#divContenidos').height(height-120);
	 $('#divDatos').height(height-170);
	 $('#OficioNuevo').height(height-220);
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
	<div style="float:left; vertical-align:middle; width:70%; padding-left:30px; padding-top:10px;" ><em>Nueva correspondencia entrante</em></div>

	<!--Para buscar -->
	<div align="center" style="width:20%; padding-left:15px; padding-right:15px; float:left; position:relative; padding-top:1px;">&nbsp;</div>

<!-- Para cerrar y ayuda -->
    <div align="right" style="padding-right:15px; width:10%; float:right; padding-top:5px;">
        <a href="#" class="btn btn-sm btn-info" title="Ayuda">?</a> &nbsp;&nbsp;
      <a href="Main.asp" class="btn btn-sm btn-danger">cerrar</a> &nbsp;&nbsp;
       <!-- <input type="button" onclick="tableToExcel('TestTable', 'CADIDO Global')" value="Exportar CADIDO a Excel" class="btn btn-sm  btn-outline-success"> -->
    </div>
</div>


<div id="divContenidos"  align="left" style="overflow:auto; padding-right: 15px; padding-top: 15px; padding-left: 0px; padding-bottom: 0px; border-right: #fff 1px solid; border-top: #fff 1px solid; border-left: #fff 1px solid; border-bottom: #fff 1px solid; scrollbar-arrow-color : #999999; scrollbar-face-color : #cccccc; scrollbar-track-color :#3333333; position: relative; left: 0px; top: 0px; width: 100%; float:left;;">

    <!-- Recoleccion de datos -->
    <div id="divDatos"  align="left" style="overflow:auto; padding: 15px; border-right: #fff 1px solid; border-top: #fff 1px solid; border-left: #fff 1px solid; border-bottom: #fff 1px solid; scrollbar-arrow-color : #999999; scrollbar-face-color : #cccccc; scrollbar-track-color :#3333333; position: relative; left: 0px; top: 0px; width: 40%; float:left;;">
    <form id="form1" name="form1" method="post" action="GuardaOficio.asp?Seccion=6&Parent=NewCorr.asp&SeccPar=3">
        <table style="padding-left:15px;" width="97%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="2"><h5>Datos del oficio</h5></td>
            </tr>
          <tr>
            <td width="39%">&nbsp;</td>
            <td width="61%">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" valign="top" >Número de oficio.</td>
            <td>
              <input name="txtNoOficio" type="text"  id="txtNoOficio" size="45" class="form-control" required="required" />
            </td>
          </tr>
          <tr>
            <td height="30" valign="top" >Dependencia</td>
            <td>              
              <input name="tags" type="text"  id="tags" size="45" class="form-control" required="required"/>
            </td>
          </tr>
          <tr>
            <td height="30" valign="top">Deptartamento</td>
            <td>             
              <input name="txtDepartamento" type="text"  id="txtDepartamento" size="45" class="form-control" required="required"/>
            </td>
          </tr>
          <tr>
            <td height="30" valign="top">Remitente</td>
            <td>              
              <input name="txtRemitente" type="text" id="txtRemitente" size="45" class="form-control" required="required"/>
            </td>
          </tr>
          <tr>
            <td height="30" valign="top">Fecha recepcion</td>
            <td>              
              <input name="txtFecha" type="text"  id="txtFecha" value="<%=formatdatetime(date,1)%>" size="45" readonly="readonly" class="form-control" required="required"/>
            </td>
          </tr>
          <tr>
            <td height="30" valign="top">Destino</td>
            <td>              
              <input name="txtDestino" type="text"  id="txtDestino" size="45" class="form-control" required="required"/>
            </td>
          </tr>
          <tr>
            <td valign="top">Asunto</td>
            <td>              
              <textarea name="txtAsunto" cols="45" rows="5"  id="txtAsunto" class="form-control" required="required"></textarea>
            </td>
          </tr>
          <tr>
            <td valign="top">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="2" valign="top"><h5>Turnado a</h5></td>
            </tr>
          <tr>
          <%
		     rst.open "select * from AreasContraloria where Activa = 1 order by idAreasContraloria", cnn
		  %>
            <td colspan="2" valign="top"><label for="cmbTurnado"></label>
              <select name="cmbTurnado" id="cmbTurnado" class="form-control">
                <%do while not rst.eof%>
                	<option value="<%=rst.fields("idAreasContraloria")%>"><%=rst.fields("NombreAreaC")%></option>
                <% rst.movenext
				   loop 
				%>
              </select></td>
            </tr>
          <tr>
            <td valign="top">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="2" valign="top"><h5>Anotaciones</h5></td>
          </tr>
          <tr>
            <td colspan="2" valign="top"><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td width="50%"><table width="200">
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup1" value="1" id="CheckboxGroup1_0" />
                        Urgente</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup1" value="2" id="CheckboxGroup1_1" />
                        Atencion</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup1" value="3" id="CheckboxGroup1_2" />
                        Comentario</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup1" value="4" id="CheckboxGroup1_3" />
                        Conocimiento</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup1" value="5" id="CheckboxGroup1_4" />
                        Seguimiento</label>
                    </td>
                  </tr>
                </table></td>
                <td width="50%"><table width="200">
                  
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup2" value="6" id="CheckboxGroup2_0" />
                        Asistir</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup2" value="7" id="CheckboxGroup2_1" />
                        Analizar</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup2" value="8" id="CheckboxGroup2_2" />
                        Respuesta</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup2" value="9" id="CheckboxGroup2_3" />
                        Informar resolucion</label>
                    </td>
                  </tr>
                  <tr>
                    <td>
                      <label>
                        <input type="checkbox" name="CheckboxGroup2" value="10" id="CheckboxGroup2_4" />
                        Archivo</label>
                    </td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table></td>
            </tr>
          <tr>
            <td valign="top">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td valign="top">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="2" align="center" valign="top"><input class="btn btn-outline-info" type="submit" name="cmbGuardar" id="cmbGuardar" value="Guardar" /></td>
            </tr>
        </table>
	 </form>
	 </div>
     <!-- FIN Recoleccion de datos -->
	
     <!-- Dato adjunto -->
     <div style="width:60%; float:right; padding:15px;">
	     <h5>Documento adjunto</h5><br />
         <iframe src="CapturaOficio.asp" width="100%" scrolling="no" frameborder="0" allowtransparency="true" name="OficioNuevo" id="OficioNuevo"></iframe>
      </div>
     <!-- FIN Dato Adjunto -->

</div>
<%end if%>

</body>
</html>
