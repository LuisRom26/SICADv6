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
      $('#divPrincipal').height(height-110);
});
function Resp(){
	$(document).ready(function(){
      var height = $(window).height();
     $('#divPrincipal').height(height-110);
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
	
		ValidateUser = Session("txtUser")
		ValidatePass = Session("txtPass")
		
			
		rst.CursorLocation = 2
		rst.CursorType = 0
		rst.LockType = 3
		rst2.CursorLocation = 2
		rst2.CursorType = 0
		rst2.LockType = 3
	
'---------------------------------------------------------------------------------------------
'	CREAR LA CADENA DE PENDIENTES SEGUN EL USUARIO ACTIVO
'---------------------------------------------------------------------------------------------

'Validaciones de correspondencia:
CadPend = "<br>Estos son lo elementos que necesitan su atención: <br><br>"
if session("NivelCorr") <> 0 then
	'session("SQL") = "select count(Oficio) as SinAtender from Oficios where Turnado = '" & session("IDUsuario") & "' and Resuelto = 0 and EnSeguimiento = 0"
	if session("NivelCorr") = 2 then
	rst2.open "select count(Oficio) as SinAtender from Oficios where Area = " & Session("Area") & " and Resuelto = 0 and EnSeguimiento = 0", cnn
	else
	rst2.open "select count(Oficio) as SinAtender from Oficios where Turnado = '" & session("IDUsuario") & "' and Resuelto = 0 and EnSeguimiento = 0", cnn
	end if
		if cint(rst2.fields("SinAtender")) > 0 then
			CadPend = CadPend & "<strong>" & CStr(rst2.fields("SinAtender")) & "</strong> Oficios <strong>sin atender</strong> en su correspondencia<br>"
		end if
	rst2.close
	
	if session("NivelCorr") = 2 then
	rst2.open "select count(Oficio) as EnSeguimiento from Oficios where Area = '" & session("Area") & "' and Resuelto = 0 and EnSeguimiento = 1", cnn
	else
	rst2.open "select count(Oficio) as EnSeguimiento from Oficios where Turnado = '" & session("IDUsuario") & "' and Resuelto = 0 and EnSeguimiento = 1", cnn
	end if
		
		if cint(rst2.fields("EnSeguimiento")) > 0 then
			CadPend = CadPend & "<strong>" & CStr(rst2.fields("EnSeguimiento")) & "</strong> Oficios <strong>en seguimiento</strong> en su correspondencia<br>"
		end if
	rst2.close
end if

'Para las validaciones de los minutarios
	
	
if Session("NivelMinutario") > 0 then
	
	if session("NivelMinutario") = 2 or session("NivelMinutario") = 1 then
		rst2.open "select count(Fase) as SinAsignar from DetalleMinutarios where Fase = 0 and AreaSolicitante = '" & session("Area") & "'", cnn
	end if
	if session("NivelMinutario") = 3 then
		rst2.open "select count(Fase) as SinAsignar from DetalleMinutarios where Fase = 0 and Seguimiento = '" & session("Usuario") & "'", cnn
	end if
		if Cint(rst2.fields("SinAsignar")) > 0 then
			CadPend = CadPend & "<strong>" & CStr(rst2.fields("SinAsignar")) & "</strong> Oficios reservados <strong>sin asignar</strong> en el minutario<br>"
		end if
	rst2.close
	
	'session("SQL") = "select count(Fase) as SinAsignar from DetalleMinutarios where Fase = 1 and Seguimiento = '" & session("Usuario") & "'"
	if session("NivelMinutario") = 2 or session("NivelMinutario") = 1 then
		rst2.open "select count(Fase) as SinAsignar from DetalleMinutarios where Fase = 1 and AreaSolicitante = '" & session("Area") & "'", cnn
	end if
	if session("NivelMinutario") = 3 then
		rst2.open "select count(Fase) as SinAsignar from DetalleMinutarios where Fase = 1 and Seguimiento = '" & session("Usuario") & "'", cnn
	end if
		if Cint(rst2.fields("SinAsignar")) > 0 then
			CadPend = CadPend & "<strong>" & CStr(rst2.fields("SinAsignar")) & "</strong> Oficios <strong>sin acuse</strong> en el minutario<br>"
		end if
	rst2.close
end if



Session("CadPend") = CadPend

'---------------------------------------------------------------------------------------------
'	FIN DE CREAR LA CADENA DE PENDIENTES SEGUN EL USUARIO ACTIVO
'---------------------------------------------------------------------------------------------
%>

<!-- Para las notificaciones -->
<script>
$(document).ready(function () {
	    var messages = [
        [ 'uno', 'alert-info','Bienvenido(a) <strong><%=session("Usuario")%></strong><%=Session("CadPend")%><br><div style="text-align:right"><a href="MainNotify.asp">ver detalles ...</a></div>']
      ];
	   for(i=0;i<messages.length;i++){
		var message = messages[Math.floor(Math.random() * messages.length)];
        $('.notify').append('<div id="'+message[0]+'" class="alert '+message[1]+' notify2"><button type="button" class="close">×</button>'+message[2]+'</div>');
			$('#'+message[0]).delay(500).fadeIn( "slow");
			$('#'+message[0]).delay(20000).fadeOut( "slow");
    }
	$(document).on('click', '.close', function () {$(this).parent().hide();});
});
</script>
<style>
.notify {
    position: fixed;
	z-index: 99999999;
	width: 450px;
	left: 30px;
	bottom: 30px;
}
.notify .alert{
	box-shadow: 0px 2px 5px -1px #000;
   display: none;
}
</style>
<!-- Fin de las Notificaciones -->

</head>

<body onresize="Resp();">
<!-- #include file="MenuFileManager.asp" -->

<%if Session("SICAD_Active") = 1 then%>
<br />&nbsp;
<div id="divPrincipal" style="width:100%; background-image:url(Imagenes/FondoMain.png); background-repeat:no-repeat; background-position:bottom right; padding:15px;">
&nbsp;<%=Session("SQL")%>
</div>

    <div class="container">
    <div class="row">
            <div class="notify"></div>
	</div>
    </div>
    
<%end if%>

</body>
</html>
