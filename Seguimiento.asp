<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.CharSet = "utf-8"%>
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
				
					rst.CursorLocation = 2
					rst.CursorType = 0
					rst.LockType = 3
					
		'Para el log de actividades
		rst.open "select * from Monitoreo where idMonitoreo = 1", cnn
			rst.addnew
				rst.fields("usrID") = Session("IDUsuario")
				rst.fields("usrNombre") = Session("Usuario")
				rst.fields("usrArea") = Session("Area")
				rst.fields("Comentario") =  "El usuario escribió un comentario de seguimiento en el folio No. " & request("Folio") & " del módulo de correspondencia.<br><br><em><strong>Texto del comentario:</strong></em><br>"  & request("txtSeguimiento")
				rst.fields("Fecha") = date
				rst.fields("Hora") = time
				rst.fields("Modulo") = "SEGUIMIENTO CORRESPONDENCIA"
			rst.update
		rst.close

					
						rst.open "UPDATE Oficios SET DioSeguimiento = '" & session("Usuario") & "', Seguimiento = '" & request("txtSeguimiento") & "', EnSeguimiento = 1, FechaSeguimiento = '" & year(date()) & "-" & month(date()) & "-" & day(date()) & "' WHERE Folio =  " & request("Folio")  , cnn
			
Set rst = Nothing
%>

		<!-- Para los redireccionamientos -->
		<script language="Javascript">
			window.open('BuscaCorr.asp', '_self');
 		</script> 

