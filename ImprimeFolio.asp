<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <title>Formato de Control de Correspondencia</title>
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

   
<script type="text/javascript" src="JS/jspdf.debug.js"></script>
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
					
			rst.open "select * from Oficios inner join areascontraloria on (folio = " & request("ID") & ") and (AreaInicial=idAreasContraloria)", cnn
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
		
		loremipsum =  '<%=Cad%>';
		
		lines = doc.setFont(font[0], font[1])
					.setFontSize(size)
					.splitTextToSize(loremipsum, 7.5)
		
		verticalOffset = (lines.length + 0.5) * size / 72
		doc.text(0.5, (verticalOffset + size / 72)-0.1, lines)
		
		
		
		doc.addImage(imgData, 'JPG', 0, 0, 8.5, 11, 'Logo'); // Cache the image using the alias 'monkey'
		
		doc.setFontSize(10);
		doc.text(0.5, 5, lines);
		
		
		

		doc.setFontSize(16);
		doc.setFontType('bold');
		doc.text(6.9, 1.2, '<%=right("00000" &  rst.fields("Folio"), 5)%>');
		
		
		
		doc.setFontType('bold');
		doc.setFontSize(10);
		doc.text(0.5, 1.9, 'OFICIO No. \u0020\u0020\u0020\u0020\u0020\u0020<%=ucase(rst.fields("Oficio"))%>\nDEPENDENCIA \u0020\u0020\u0020\u0020\u0020<%=ucase(rst.fields("Depe"))%>\nDEPARTAMENTO \u0020\u0020\u0020\u0020<%=ucase(rst.fields("Departamento"))%>\nREMITENTE \u0020\u0020\u0020\u0020\u0020\u0020\u0020<%=ucase(rst.fields("Remitente"))%>\nDIRIGIDO A \u0020\u0020\u0020\u0020\u0020\u0020<%=ucase(rst.fields("Destino"))%>\nRECEPCION \u0020\u0020\u0020\u0020\u0020\u0020\u0020<%=ucase(formatdatetime(rst.fields("Fecha1"),1))%> ');
		
		doc.setFontType('bold');
		doc.setFontSize(10);
		doc.text(4.9, 3.5, '<%=ucase(rst.fields("NombreAreaC"))%>');
		
		<%
		cad = ""
		n = 1
		dim Com(20)
		
		for k = 1 to 15
			Com(k) = 0
		next
		
		for k = 1 to len(rst.fields("Comentario"))
			if mid(rst.fields("Comentario"), k, 1) <> "," then
				cad = cad & mid(rst.fields("Comentario"), k, 1)
			else
			  if cad = "" then
			  	 cad = "0"
			  end if
			  if cad = " " then
			  	 cad = "0"
			  end if
					'response.Write("N=" & n & " Cad=" & cad & ". ")
					Com(int(trim(cad))) = int(trim(cad))
					n = n + 1
			
			Cad = ""
			end if
		next
		%>
		
		<%if com(1) <> 0 then%>
			doc.text(0.5, 3.4, '[ x ] URGENTE');
		<%else%>
			doc.text(0.5, 3.4, '[   ] URGENTE');
		<%end if%>
		
		<%if com(2) <> 0 then%>
			doc.text(0.5, 3.6, '[ x ] ATENCION');
		<%else%>
			doc.text(0.5, 3.6, '[   ] ATENCION');
		<%end if%>
		
		<%if com(3) <> 0 then%>
			doc.text(0.5, 3.8, '[ x ] COMENTARIO');
		<%else%>
			doc.text(0.5, 3.8, '[   ] COMENTARIO');
		<%end if%>
		
		<%if com(4) <> 0 then%>
			doc.text(0.5, 4, '[ x ] CONOCIMIENTO');
		<%else%>
			doc.text(0.5, 4, '[   ] CONOCIMIENTO');
		<%end if%>
		
		<%if com(5) <> 0 then%>
			doc.text(0.5, 4.2, '[ x ] SEGUIMIENTO');
		<%else%>
			doc.text(0.5, 4.2, '[   ] SEGUIMIENTO');
		<%end if%>
		
		//A PARTIR DE AQUI ES LA SEGUNDA COLUMNA
		
		<%if com(6) <> 0 then%>
			doc.text(2.5, 3.4, '[ x ] ASISTIR');
		<%else%>
			doc.text(2.5, 3.4, '[   ] ASISTIR');
		<%end if%>
		
		<%if com(7) <> 0 then%>
			doc.text(2.5, 3.6, '[ x ] ANALIZAR');
		<%else%>
			doc.text(2.5, 3.6, '[   ] ANALIZAR');
		<%end if%>
		
		<%if com(8) <> 0 then%>
			doc.text(2.5, 3.8, '[ x ] RESPUESTA');
		<%else%>
			doc.text(2.5, 3.8, '[   ] RESPUESTA');
		<%end if%>
		
		<%if com(9) <> 0 then%>
			doc.text(2.5, 4, '[ x ] INFORMAR RESOLUCION');
		<%else%>
			doc.text(2.5, 4, '[   ] INFORMAR RESOLUCION');
		<%end if%>
		
		<%if com(10) <> 0 then%>
			doc.text(2.5, 4.2, '[ x ] ARCHIVO');
		<%else%>
			doc.text(2.5, 4.2, '[   ] ARCHIVO');
		<%end if%>
		
		
		
		doc.setFontType('italic');
		doc.setFontSize(6);
		doc.text(2.9, 10.8, "Documento controlado electrónicamente, cualquier copia impresa es un documento no controlado [SICAD v3.01b]");
		doc.setFontSize(7);
		doc.setFontType('normal');
		//doc.text(0.1, 10.6, "Hola mundo");
		
		
		
		
doc.output('datauri');
}

getImageFromUrl('imagenes/FondoCorrespondencia.jpg', createPDF);

		

		//pdf.save('ER-<%=generadordeclaves(6)%>.pdf');

</script>




</body>

</html>


