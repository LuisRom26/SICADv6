<!-- #include file="BarraMenu2.asp" -->

<body>
  <%
		set cnn = server.CreateObject("ADODB.CONNECTION")
		set rst = server.CreateObject("ADODB.RECORDSET")
		Archivo = request.ServerVariables("APPL_PHYSICAL_PATH") & "/config.txt"
		set ConFile = createobject ("scripting.filesystemobject")
		set Fichero = ConFile.OpenTextFile(Archivo)
		TextoFichero = Fichero.ReadAll()
						
		Fichero.Close()
		
						
		strConexion = TextoFichero
		cnn.open strConexion
	
		ValidateUser = request.Form("txtUser")
		ValidatePass = request.Form("txtPass")
		
				
		rst.CursorLocation = 2
		rst.CursorType = 0
		rst.LockType = 3
		'RESPONSE.Write("SELECT * FROM Empleados WHERE NoControl = '" & ValidateUser & "' and usrPass = '" & ValidatePass & "'")
		rst.open "SELECT * FROM Empleados WHERE NoControl = '" & ValidateUser & "' and usrPass = '" & ValidatePass & "'" , cnn
		
		if rst.eof and rst.bof then
%>


<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" background="Imagenes/FondoLogin2.png" style="background-repeat:no-repeat">
  <tr>
    <td height="300">
    <form id="form2" name="form2" method="post" action="ValidateUser2.asp">
   <div style="padding: 2em; margin:1em 0 1em 4em; width:480px;"><table width="99%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="50%">&nbsp;</td>
    <td width="50%">&nbsp;</td>
  </tr>
  <tr>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><label for="txtUser"></label>
      <input name="txtUser" type="text" id="txtUser" class="form-control"/></td>
    <td align="right"><label for="txtPass"></label>
      <input name="txtPass" type="password" id="txtPass" class="form-control"/></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><input type="hidden" name="hiddenAlto" id="hiddenAlto" /></td>
    </tr>
  <tr>
    <td colspan="2" align="right"><input type="submit" name="cmdEnviar2" id="cmdEnviar2" value="ingresar" class="btn btn-success" /></td>
  </tr>
  <tr>
    <td colspan="2" align="center"></td>
  </tr>
  <tr>
    <td colspan="2" align="left">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" align="center" class="TextoNormalVerde">

<div class="alert alert-danger">
    los datos proporcionados son incorrectos o el usuario está deshabilitado
</div>

</td>
  </tr>
  <tr>
    <td colspan="2" align="left">&nbsp;</td>
  </tr>
   </table>
</div>
      </form>
     </td>
  </tr>
  <tr>
    <td height="100">&nbsp;</td>
  </tr>
</table>

<%else 
		Session("SICAD_Active") = 1
		Session("Usuario") = rst.fields("Nombre") & " " & rst.fields("ApellidoP")  & " " & rst.fields("ApellidoM")
		Session("NivelUsuario")  = rst.fields("PuestoFuncional")
		
		Session("NivelCorr") = rst.fields("NivelCorr")
		Session("Corr") = rst.fields("Corr")
		Session("Area") = rst.fields("Area")
		Session("Depto") = rst.fields("Depto")
		
		Session("Opciones") = rst.fields("Opciones")
		Session("OpcionesV4") = rst.fields("OpcionesV4")
		Session("IDUsuario") = rst.fields("idEmpleados")
		session("CTRLUsuario") = rst.fields("NoControl")
		
		Session("NivelPliego") = rst.fields("NivelPliego")
		Session("NivelVehiculo") = rst.fields("NivelVehiculo")
		Session("NivelMinutario") = rst.fields("NivelMinutario")
		Session("NivelInventario") = rst.fields("NivelInventario")
		Session("AdminMinutario") = rst.fields("AdminMinutario")
		Session("Nivel5M") = rst.fields("Nivel5M")
		Session("NivelMonitoreo") = rst.fields("NivelMonitoreo")
		
		Session("PermisoEditarEtica") = 0
		Session("PermisoEditarCI") = 0
		
		Session("AreaSICOP") = rst.fields("AreaSICOP")
		Session("UserSICOP") = rst.fields("UserSICOP")
		
		


		
		response.Redirect("Main.asp")

 end if
%>
  


<table width="1024" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="60%" colspan="2" align="center" valign="top" class="TextoPiePagina" >© Derechos Reservados 2006-2018 :: Contraloría General  del Estado de Colima<br />
Calzada Pedro A. Galván Sur # 454 :: Colonia Centro :: CP   28000 :: Colima, Colima, México<br />
Teléfono +52 (312) 314-4475 / FAX +52 (312)   312-8988<br />
www.col.gob.mx/contraloria - despacho.cg@gobiernocolima.gob.mx<br />
<br /></td>
  </tr>
</table>
<script type="text/javascript">
 var height = $(window).height();
 document.getElementById("hiddenAlto").value = height;
 alert(height);
</script>
</body>
</html>
