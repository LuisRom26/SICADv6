<%if Session("SICAD_Active") = 0 then %>
<script>
	window.onload = function Emergente() {
			$('#ErrorLogin').modal('show');
		}
</script> 
<%else%>

<!-- para validar los persmisos de usuario -->
<%
			Opc = split(session("OpcionesV4"),",")
			Opcion1 = 0 'Nueva correspondencia
			Opcion2 = 0 'Minutario de Oficios
			Opcion3 = 0 'Revisar correspondencia
			Opcion4 = 0 'Informes 5 al Millar
			Opcion5 = 0 'Administracion de bienes
			Opcion6 = 0 'Bienes por Usuario
			Opcion7 = 0 'Servicios vehiculares
			Opcion9 = 0 'Pliegos y permisos
			Opcion15 = 0 'Archivos SISER
			Opcion18 = 0 'Autorizacion de pliegos y permisos
			Opcion10 = 0 'Administracion de la Normateca
			Opcion12 = 0 'Administracion Comités de Ética
			Opcion13 = 0 'Administracion Control Interno Institucional
			Opcion19 = 0 'Archivo Historico
 			
			for k = 0 to ubound(Opc)
				if Opc(k) = "1" then
					Opcion1 = 1
				end if
				if Opc(k) = "2" then
					Opcion2 = 1
				end if
				if Opc(k) = "3" then
					Opcion3 = 1
				end if
				if Opc(k) = "4" then
					Opcion4 = 1
				end if
				if Opc(k) = "6" then
					Opcion6 = 1
				end if
				if Opc(k) = "9" then
					Opcion9 = 1
				end if
				if Opc(k) = "15" then
					Opcion15 = 1
				end if
				if Opc(k) = "18" then
					Opcion18 = 1
				end if
				if Opc(k) = "10" then
					Opcion10 = 1
				end if
				if Opc(k) = "12" then
					Opcion12 = 1
				end if
				if Opc(k) = "13" then
					Opcion13 = 1
				end if
				if Opc(k) = "7" then
					Opcion7 = 1
				end if
				if Opc(k) = "5" then
					Opcion5 = 1
				end if
				if Opc(k) = "19" then
					Opcion19 = 1
				end if
			next
%>


<iframe src="KeepAlive.asp" width="250" height="55" frameborder="0" style="display:none;"></iframe>
<div class="wrapper">
<!--Navigation Start-->
<nav class="navigation" style="border-bottom:#ccc 1px solid;">
  <ul>
    <li class="active">
      <a style="cursor:pointer">Correspondencia</a>
      <ul class="children sub-menu">
        <li>
          <%if Opcion1 = "1" then%>
      			<a href="NewCorr.asp"> Nueva correspondencia</a>
          <%else%>
        		<a data-toggle="modal" data-target="#ErrorPermiso" style="cursor:pointer" >Nueva correspondencia</a>
          <%end if%>
        </li>
        <li>
          <%if Opcion2 = "1" then%>
	      		<a href="BuscaMin.asp" >Minutario de Oficios</a>
          <%else%>
        		<a data-toggle="modal" data-target="#ErrorPermiso" style="cursor:pointer">Minutario de Oficios</a>
          <%end if%>
        </li>
        <li>
          <%if Opcion3 = "1" then%>
	      		<a href="BuscaCorr.asp">Revisar correspondencia</a>
          <%else%>
        		<a data-toggle="modal" data-target="#ErrorPermiso" style="cursor:pointer" >Revisar correspondencia</a>
          <%end if%>
        </li>
        
      </ul>
    </li>
    
    <li>
      <a href="CloseSession.asp">salir</a>
    </li>
  </ul>
</nav>

<div style="width:25%; float:right; background-color:#343a40; height:50px;"> &nbsp;
    <div align="right" style="padding-right:15px; float:left; background-color:#343a40; width:89%; height:50px; padding-top:6px;"  class="TextoNormalBlanco">
        <%=Session("Usuario")%><br /><%=Session("NivelUsuario")%>
    </div>
    <div align="left" style="float:right; background-color:#343a40; width:8%; height:50px; padding-top:6px;"  class="TextoNormalBlanco">
        <img src="Imagenes/IconUser.png" width="27" height="27" />
    </div>
</div>
<!--Navigation End-->
</div>


<%end if%>

 <!-- Para error de login -->
  <div class="modal fade" id="ErrorLogin" data-keyboard="false" data-backdrop="static">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">aviso</h4>
          <!--<button type="button" class="close" data-dismiss="modal">&times;</button>-->
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
          <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="TextoNormal16" align="center">Su sesión no es válida o ha expirado, por favor ingresa nuevamente al sistema<br><br><a href="default.asp" class="btn btn-success" >iniciar sesión</a> </td>
    </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          
        </div>
        
      </div>
    </div>
</div>  

<!-- Error de permisos -->
 <!-- Para error en los permisos -->
  <div class="modal fade" id="ErrorPermiso">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">Aviso</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
          <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="TextoNormal16" align="center">El usuario actual no tiene los privilegios necesarios para ingresar al módulo seleccionado.</td>
    </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <button type="button" class="btn btn-danger" data-dismiss="modal"><span class="TextoBoton">cerrar</span></button>
        </div>
        
      </div>
    </div>
  </div>    

<!-- Fin Error de Permisos -->

<!-- Para error de login -->
  <div class="modal fade" id="ErrorLogin" data-keyboard="false" data-backdrop="static">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">aviso</h4>
          <!--<button type="button" class="close" data-dismiss="modal">&times;</button>-->
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
          <table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td class="TextoNormal16" align="center">Su sesión no es válida o ha expirado, por favor ingresa nuevamente al sistema<br><br><a href="default.asp" class="btn btn-success" >iniciar sesión</a> </td>
    </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          
        </div>
        
      </div>
    </div>
</div>  

