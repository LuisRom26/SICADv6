<!-- #include file="BarraMenu2.asp" -->

<body>
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
    <td colspan="2" align="center"><input type="hidden" name="hiddenAlto" id="hiddenAlto" />&nbsp;</td>
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

<div class="alert alert-info">
    introduce tus datos de acceso</div>

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

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" class="TextoPiePagina">© Derechos Reservados 2006-2018 :: Contraloría General  del Estado de Colima<br />
Calzada Pedro A. Galván Sur # 454 :: Colonia Centro :: CP   28000 :: Colima, Colima, México<br />
Teléfono +52 (312) 314-4475 / FAX +52 (312)   312-8988<br />
www.col.gob.mx/contraloria - despacho.cg@gobiernocolima.gob.mx<br /></td>
  </tr>
</table>
<script type="text/javascript">
 var height = $(window).height();
 document.getElementById("hiddenAlto").value = height;
 
</script>
</body>
</html>
