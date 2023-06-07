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
<title></title>

<!-- para los Hiddens-->
<script type="text/javascript"> 
    function idMin(ID) { 
	   var varID = ID;
	   document.getElementById("idMin").value = varID;
	   
	   $('#EliminaAdjunto').modal('show');

      return true; 
    }       
</script>


 
<!--<link type="text/css" href="http://jquery-ui.googlecode.com/svn/tags/1.7/themes/redmond/jquery-ui.css" rel="stylesheet" /> -->
<link type="text/css" href="css/jquery-ui.css" rel="stylesheet" />
<link href="CSS/Estilos.css" media="screen" type="text/css" rel="stylesheet" />
<link href="CSS/EstilosClick.css" media="screen" type="text/css" rel="stylesheet" />
<link href="CSS/impresora.css" media="print" type="text/css" rel="stylesheet" />

<!--Para el calendario -->
<link rel="stylesheet" href="CSS/calendar.css" media="screen">

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


<script type="text/javascript"> 
    $(function() {         
        $('#dialog5').dialog({ 
            autoOpen: false, 
            width: 400 
        }); 
    });
 
    function MostrarDialog5(idPac) { 
      var varIdPac = idPac;
	  
	   document.getElementById("idMin").value = varIdPac;
	  
	  $('#dialog5').dialog('option', 'modal', true).dialog('open'); 
      
      return true; 
    }       
</script>


</head>

<body>
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

		rst.open "select * from DetalleMinutarios where idDetalleMinutarios=" & request("IDMin"), cnn
		
	NumOf = rst.fields("NumOficio")
	Ar = rst.fields("ArchivoTMP") 
	
	NumOfNew = ""
	for k = 1 to len(NumOf)
		if mid(NumOf, k, 1) = "/" then
			NumOfNew = NumOfNew & "-"
		else
			NumOfNew = NumOfNew & mid(NumOf, k, 1)
		end if
	next
	
	

%>


<div align="center"><embed src="../Minutario/<%=NumOfNew%><%=right(Ar,4)%>" width="100%" height="300" ></div>
<br />
<div align="center"><a onclick="idMin('<%=request("IDMin")%>')", style="cursor:pointer" class="btn btn-outline-danger">Eliminar adjunto</a></div>

<script type="text/javascript">
function submitform()
{
     var theForm = document.forms['formX'];
     if (!theForm) {
         theForm = document.formX;
     }
     theForm.submit();
}
</script> 


 <!-- Para eliminar un adjunto -->
  <div class="modal fade" id="EliminaAdjunto">
    <div class="modal-dialog">
      <div class="modal-content">
      
        <!-- Modal Header -->
        <div class="modal-header">
          <h4 class="modal-title">pregunta</h4>
          <button type="button" class="close" data-dismiss="modal">&times;</button>
        </div>
        
        <!-- Modal body -->
        <div class="modal-body">
        <form id="formX" name="formX" method="post" action="EliminaAcuse.asp"><table width="380" border="0" cellspacing="0" cellpadding="0" align="center">

  <tr>
    <td width="389" align="center">¿Está usted seguro de eliminar este acuse?</td>
  </tr>
  <tr>
    <td>
      <input type="hidden" name="idMin" id="idMin" />
    </td>
  </tr>
  <tr>
  	<td>&nbsp;</td>
  </tr>
  <tr>
    <td align="center"><a class="btn btn-outline-success" href="javascript: submitform()">sí, eliminar</a>&nbsp;&nbsp;&nbsp;<a class="btn btn-outline-danger" data-dismiss="modal">no, regresar</a></td>
  </tr>
  <tr>
    <td colspan="2">&nbsp;</td>
  </tr>
</table></form>

        </div>
        
        <!-- Modal footer -->
        <div class="modal-footer">
          <button type="button" class="btn btn-danger" data-dismiss="modal">cerrar</button>
        </div>
        
      </div>
    </div>
  </div>


</body>
</html>
