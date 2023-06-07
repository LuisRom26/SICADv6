<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%	Session("rDestino") = "BuscaMin.asp"
	session("rFrame") = 4
	Session("frameNivel4") = "TabsCorrespondencia"
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Reporte de acuses pendientes</title>
</head>
<body>
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
   
<script type="text/javascript" src="js/jspdf.debug.js"></script>
<script type="text/javascript">
	
	
        var pdf = new jsPDF();
        
        
// Because of security restrictions, getImageFromUrl will
// not load images from other domains.  Chrome has added
// security restrictions that prevent it from loading images
// when running local files.  Run with: chromium --allow-file-access-from-files --allow-file-access
// to temporarily get around this issue.
var getImageFromUrl = function(url, callback) {
    var img = new Image();

    img.onError = function() {
        alert('Cannot load image: "'+url+'"');
    };
    img.onload = function() {
        callback(img);
    };
    img.src = url;
}


		<%
			set cnn = server.CreateObject("ADODB.CONNECTION")
			set rst = server.CreateObject("ADODB.RECORDSET")
			set rst2 = server.CreateObject("ADODB.RECORDSET")
			Archivo = request.ServerVariables("APPL_PHYSICAL_PATH") & "/SICAD5/config.txt"
			set ConFile = createobject ("scripting.filesystemobject")
			set Fichero = ConFile.OpenTextFile(Archivo)
			TextoFichero = Fichero.ReadAll()
							
			Fichero.Close()

			strConexion = TextoFichero
			cnn.open strConexion
				
			rst.CursorLocation = 2
			rst.CursorType = 0
			rst.LockType = 3
			rst2.CursorLocation = 2
			rst2.CursorType = 0
			rst2.LockType = 3
					
			rst.open "SELECT * FROM detalleminutarios d inner join empleados on (Solicitante=idEmpleados) and (fase=1) and (Solicitante=" & request("cmbResponsable") & ")", cnn
			
			if rst.eof and rst.bof then
				response.Redirect("SinPendientes.asp")
			end if
			
			
			TotBie = 0
			do while not rst.eof
			    TotBie = TotBie+1
				rst.movenext
			loop
			rst.movefirst
			
			Pags = TotBie/29
			EnteroPags = int(Pags)
			if Pags > EnteroPags then Pags = EnteroPags + 1
			
			Cadena = rst.fields("Nombre")& " " & rst.fields("ApellidoP") & " " & rst.fields("ApellidoM")
		%>
// Since images are loaded asyncronously, we must wait to create
// the pdf until we actually have the image.
// If we already had the jpeg image binary data loaded into
// a string, we create the pdf without delay.
	
	var createPDF = function(imgData) {
    var doc = new jsPDF('p','in','letter');
	var lines;
			
		var sizes = [10, 16, 20];
		var fonts = [['Courier',''],['Helvetica',''], ['Times','Italic']];
		var font, size;
		var loremipsum;
		var verticalOffset = 0.5;
		var Media=5.5;
	

    // This is a modified addImage example which requires jsPDF 1.0+
    // You can check the former one at examples/js/basic.js

	// ------------ ENCABEZADO DEL DOCUMENTO ---------------
	
    	
		

		font = fonts[0];
		size = sizes[0];
		//alert(font);
		
		
		
		loremipsum = '';
		
		lines = doc.setFont(font[0], font[1])
					.setFontSize(size)
					.splitTextToSize(loremipsum, 10)
		
		verticalOffset = (lines.length + 0.5) * size / 72
		doc.text(0.5, (verticalOffset + size / 72)-0.1, lines)
		
		doc.addImage(imgData, 'JPG', 0, 0, 8.5, 11, 'Logo'); // Cache the image using the alias 'monkey'
		
		doc.setFontSize(8);
		doc.text(0.2, 4.4, lines);
		
		
		

		
		doc.setFontSize(8);
		doc.setFontType('bold');
		doc.text(0.5, 10.65, 'Total de oficios sin acuse de recibido: <%=TotBie%>\nPágina 1 de <%=round(Pags,0)%>');
		
		doc.setFontSize(12);
		doc.setFontType('bold');
		doc.text(4.0, 1.15, '\u0020\u0020<%=date%>\n\u0020<%=time%>');
		
		
		doc.setFontSize(9);
		doc.setFontType('bold');
		//doc.text(5.4, 1.25, '\u0020\u0020<%=Cadena%>');
		
		
		
		doc.setFontType('bold');
		doc.setFontSize(8);
		
		
		
		<%
		  Pos = 0
		  Pos2 = 0
		  Cadena = rst.fields("Nombre")& " " & rst.fields("ApellidoP") & " " & rst.fields("ApellidoM")
		  Pos = 88.5-(len(Cadena)/2)
		  Pos = Pos * 0.08
		  
		  Cadena2 = formatdatetime(date, 1)
  		  Pos2 = 53.5-(len(Cadena2)/2)
		  Pos2 = Pos2 * 0.08

		  
		%>
		
		
		
		doc.text(<%=Pos%>, 1.25, '<%=Cadena%>');

		
		
		<%'=rst.fields("a")%>
		<%Li = 0.15
		 Reg = 1
		 PagTot=1
		%>
		
		<%do while not rst.eof%>
				<%
			ServSol = ucase(rst.fields("Asunto"))
			Cad = ""
			for k = 1 to len(ServSol)
				if asc(mid(ServSol,k,1)) <> 13 and asc(mid(ServSol,k,1)) <> 10  then
					Cad = Cad & mid(ServSol,k,1)
				else
					if asc(mid(ServSol,k,1)) = 13 then
						Cad = Cad & "\"
					end if
					if asc(mid(ServSol,k,1)) = 10 then
						Cad = Cad & "n"
					end if
				end if
				
			next 
		 %>

		  doc.text(0.25, 1.65+<%=Li%>, '<%=rst.fields("NumOficio")%>');
		  doc.text(1.25, 1.65+<%=Li%>, '<%=left(rst.fields("Destinatario"),43)%>');
		  doc.text(4.25, 1.65+<%=Li%>, '<%=left(rst.fields("DepeDestino"),43)%>');
		  doc.text(1.25, 1.65+<%=Li%>, '\n<%=left(Cad,105)%>');
		  <%
		    if Reg = 29 then
			PagTot = PagTot+1
		   %>
 		     doc.addPage();
			 doc.addImage(imgData, 'JPG', 0, 0, 8.5, 11, 'Logo'); // Cache the image using the alias 'monkey'
			 
			 	doc.setFontSize(8);
				doc.setFontType('bold');
				doc.text(0.5, 10.65, 'Total de oficios sin acuse de recibido: <%=TotBie%>\nPágina <%=PagTot%> de <%=Pags%>');

		
				<%
		  Pos = 0
		  Pos2 = 0
		  Cadena = rst.fields("Nombre")& " " & rst.fields("ApellidoP") & " " & rst.fields("ApellidoM")
		  Pos = 88.5-(len(Cadena)/2)
		  Pos = Pos * 0.08
		  
		  Cadena2 = formatdatetime(date, 1)
  		  Pos2 = 53.5-(len(Cadena2)/2)
		  Pos2 = Pos2 * 0.08

		  
		%>
		
		
		
		doc.text(<%=Pos%>, 1.25, '<%=Cadena%>');
		doc.setFontSize(12);
		doc.setFontType('bold');
		doc.text(4.0, 1.15, '\u0020\u0020<%=date%>\n\u0020<%=time%>');

				doc.setFontType('bold');
				doc.setFontSize(8);
		
				
		  <%	 
		       Reg = 1
			   Li=0
		   end if
		  %>
		<%
		  rst.movenext
		  Li = Li + (0.15*2)
		  Reg = Reg + 1
		  TotalBienes = TotalBienes+1
		  loop
		%>
		



	
		
		
		
		
		
		
doc.output('datauri');
}

getImageFromUrl('imagenes/FondoPendientesMinutario.jpg', createPDF);

		

		//pdf.save('ER-<%=generadordeclaves(6)%>.pdf');

</script>




</body>

</html>


