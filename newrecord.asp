<!-- #include file="includes/Param.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>SITS - Security Incident Tracking System</title>
<SCRIPT LANGUAGE="JavaScript">
		function changeImages() {
			if (document.images ) {
				for (var i=0; i<changeImages.arguments.length; i+=2) {
					document[changeImages.arguments[i]].src = changeImages.arguments[i+1];
					}
				}
		}
function gettext(){ 
var sel = document.selection; 
if (sel!=null) { 
    var rng = sel.createRange(); 
    if (rng!=null) 
	  myRemote = launch("http://uioap009.andespetro.com/SITS/docs/"+rng.text, "SITS", "height=400,width=400,channelmode=0,dependent=0,directories=1,fullscreen=0,location=1,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0", "myWindow");      
} 
} 

function launch(newURL, newName, newFeatures, orgName) {
  var remote = open(newURL, newName, newFeatures);
  if (remote.opener == null)
    remote.opener = window;
  remote.opener.name = orgName;
  return remote;
}		
</SCRIPT>
<link rel="stylesheet" type="text/css" href="http://apps01.andespetro.com/includes/css/default.css">
</head>

<body>	  
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="245">
    <tr>
    <td valign="top" width="8"><!-- ########### INCLUDE Sample Navigation ########### --><!-- #include file="includes/navigation.htm" --></td>
	
	
    <td bgcolor="#FFFFFF" valign="top" align="left" width="12"><font size="4" color="#000080">&nbsp;&nbsp;</font></td>
    <td valign="top" width="868">


</SCRIPT>
<link rel="stylesheet" type="text/css" href="http://apps01.andespetro.com/includes/css/default.css">
<script language="javascript" src="includes/cal2.js"></script>
<script language="javascript" src="includes/cal_conf2.js"></script>
<script language="javascript" src="includes/fechas.js"></script>
<title>Project Tracking Tool</title>
<meta http-equiv="Content-Language" content="es-ec">
<link rel="stylesheet" type="text/css" href="http://apps01.andespetro.com/includes/css/default.css">

<%
Dim sql, Conex, RS, valida, Project_Code
action = request("action")

Fecha = Request.Form("Fecha")
Provincia = Request.Form("Provincia")
Ciudad = Request.Form("Ciudad")
Estado = Request.Form("Estado")
Clasificacion = Request.Form("Clasificacion")
CalificacionInfo = Request.Form("CalificacionInfo")
Nivel = Request.Form("Nivel")
Asunto = Request.Form("Asunto")
Nombre = Request.Form("Nombre")
Lugar = Request.Form("Lugar")
Autor = "unknown"
Descripcion = Request.Form("Descripcion")
Placa = Request.Form("Placa")
Id = Request.Form("Id")
Fuente = Request.Form("Fuente")
CalificacionFuente = Request.Form("CalificacionFuente")
Anexos = Request.Form("Anexos")
Telefono = Request.Form("Telefono")
Seguimiento = Request.Form("Seguimiento")
Foto = Request.Form("Foto")
Direccion = Request.Form("Direccion")
Marca = Request.Form("Marca")
Telfuente = Request.Form("TelFuente")
Color = Request.Form("Color")
Auditoria = UserID & " - " & Now()


if action = "" then
%>
<p>&nbsp;</p>

<p class="banner_blue">Crear un Nuevo Record</p>

<form method="POST" name="FrontPage_Form1" action="newrecord.asp?action=save" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" >
<table border="1" cellspacing="0" style="border-width:0; border-collapse: collapse" bordercolor="#111111" width="100%" bordercolorlight="#FFFFFF" cellpadding="0" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
  <tr>
    <td width="16%"><span lang="es-ec">Fecha :</span></td>
    <td width="16%">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars="/-" b-value-required="TRUE" i-minimum-length="10" i-maximum-length="11" --><input name="FECHA" type="text" onBlur="checkdate(this)" size="12" maxlength="11" readonly="true">
    <a href="javascript:showCal('FechaCalR');"><img border="0" src="http://apps01.andespetro.com/includes/Calendar/date.gif" align="middle" width="19" height="17" alt="Select a Date"></a>    
    </td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
  </tr>
  <tr>
    <td width="16%"><span lang="es-ec">Provincia :</span></td>
    <td width="16%">
    <select size="1" name="Provincia">
    <option>Azuay</option>
    <option>Bolívar</option>
    <option>Cañar</option>
    <option>Carchi</option>
    <option>Cotopaxi</option>
    <option>Chimborazo</option>
    <option>El Oro</option>
    <option>Esmeraldas</option>
    <option>Francisco de Orellana</option>
    <option>Galápagos</option>
    <option>Guayas</option>
    <option>Imbabura</option>
    <option>Loja</option>
    <option>Los Ríos</option>
    <option>Manabí</option>
    <option>Morona Santiago</option>
    <option>Napo</option>
    <option>Pastaza</option>
    <option>Pichincha</option>
    <option>Sucumbios</option>
    <option>Tungurahua</option>
    <option>Zamora Chinchipe</option>
    </select></td>
    <td width="17%"><span lang="es-ec">Clasificación :</span></td>
    <td width="17%">
    <select size="1" name="CLASIFICACION">
    <option>Activista</option>
    <option>Candidato</option>
    <option>Chantajista/Extorsionador</option>
    <option>Comunidad</option>
    <option>Contactos</option>
    <option>Contratistas</option>
    <option>Delincuencia Común</option>
    <option>Delincuencia Organizada</option>
    <option>Denunciante Contra Compañía</option>
    <option>Dirigente</option>
    <option>Empleados AdPetro</option>
    <option>Fuerza Pública</option>
    <option>Grupo Ecologista</option>
    <option>Grupo Subversivo</option>
    <option>Líderes</option>
    <option>Ministerios</option>
    <option>Municipalidades</option>
    <option>ONG</option>
    <option>Organización Gubernamental</option>
    <option>Organización Indígena</option>
    <option>Pandilleros</option>
    <option>Partidos Políticos</option>
    <option>Prefecturas</option>
    <option>Sospechosos</option>
    <option>Vehículos</option>
    </select></td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
  </tr>
  <tr>
    <td width="16%"><span lang="es-ec">Ciudad :</span></td>
    <td width="16%">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="50" --><input name="Ciudad" type="text"  size="20" maxlength="50"></td>
    <td width="17%">Idoneidad :</td>
    <td width="17%">
    <select size="1" name="CalificacionInfo">
    <option>Idóneo</option>
    <option>No Idóneo</option>
    </select></td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
  </tr>
  <tr>
    <td width="16%"><span lang="es-ec">Estado :</span></td>
    <td width="16%">
    <select size="1" name="ESTADO">
    <option>Abierto</option>
    <option>Cerrado</option>
    <option>En Proceso</option>
    </select></td>
    <td width="17%"><span lang="es-ec">Nivel :</span></td>
    <td width="17%">
    <select size="1" name="NIVEL">
    <option>Bajo</option>
    <option>Medio</option>
    <option>Alto</option>
    </select></td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
  </tr>
  <tr>
    <td width="16%">Historical Record :</td>
    <td width="84%" colspan="5">
    <!--webbot bot="Validation" s-display-name="Seguimiento" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#_" i-maximum-length="2500" --><textarea rows="3" name="Seguimiento" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="16%"><span lang="es-ec">Nombre</span> :</td>
    <td width="84%" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="500" --><input name="NOMBRE" type="text"  size="80" maxlength="500"></td>
  </tr>
  <tr>
    <td width="16%">ID :</td>
    <td width="16%">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="ID" type="text"  size="10" maxlength="150"></td>
    <td width="17%">Teléfono :</td>
    <td width="17%">
    <!--webbot bot="Validation" s-display-name="Telefono" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="Telefono" type="text"  size="10" maxlength="150"></td>
    <td width="17%">Foto :</td>
    <td width="17%">
    <!--webbot bot="Validation" s-display-name="Foto" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="." i-maximum-length="150" --><input name="Foto" type="text"  size="10" maxlength="150"></td>
  </tr>
  <tr>
    <td width="16%"><span lang="es-ec">Dirección :</span></td>
    <td width="84%" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#_" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2500" --><textarea rows="3" name="DIRECCION" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="16%">Marca Veh<span lang="es-ec">í</span>culo<span lang="es-ec"> 
    :</span></td>
    <td width="16%">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="MARCA" type="text"  size="10" maxlength="150"></td>
    <td width="17%">Color :</td>
    <td width="17%">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="COLOR" type="text"  size="10" maxlength="150"></td>
    <td width="17%">
    <span lang="es-ec">Placa : </span></td>
    <td width="17%">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="PLACA" type="text"  size="10" maxlength="150"></td>
  </tr>
  <tr>
    <td width="16%"><span lang="es-ec">Descripción :</span></td>
    <td width="84%" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#_.:" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2500" --><textarea rows="3" name="DESCRIPCION" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="16%">Fuente :</td>
    <td width="84%" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="200" --><input name="FUENTE" type="text"  size="80" maxlength="200"></td>
  </tr>
  <tr>
    <td width="16%">
    <span lang="es-ec">Tel Fuente : </span></td>
    <td width="16%">
    <!--webbot bot="Validation" s-display-name="Telefono" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="TelFuente" type="text"  size="10" maxlength="150"></td>
    <td width="17%">Lugar<br><span lang="es-ec">de Contacto </span>:</td>
    <td width="17%">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#._-@!:;,/\?=)(%" i-maximum-length="2500" --><textarea rows="3" name="LUGAR" cols="25"></textarea></td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
  </tr>
  <tr>
    <td width="16%">Calificación Fuente :</td>
    <td width="16%">
    <select size="1" name="CalificacionFuente">
    <option>Confiable</option>
    <option>Casi Confiable</option>
    <option>No Confiable</option>
    </select></td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
  </tr>
  <tr>
    <td width="16%">Anexos :</td>
    <td width="84%" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="." i-maximum-length="2500" --><textarea rows="2" name="Anexos" cols="63"></textarea><input type="button" onclick="gettext()" value="Ver" name="B1"></td>
  </tr>
  <tr>
    <td width="16%"><input type="submit" value="Grabar" name="Add"></td>
    <td width="16%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
    <td width="17%">&nbsp;</td>
  </tr>
  </table>


</form>

<%
end if
%>
<%
if action = "save" then

'Obtengo el próximo código del Projecto
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "select seq_record_code.NextVal from Dual"
RS.Open sql, Conex, 1
	Codigo = RS.Fields("NEXTVAL")
Conex.Close
Set RS = Nothing
Set Conex = Nothing

'Grabo los datos
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "insert into RECORD(FECHA,PROVINCIA,CIUDAD,ESTADO,CLASIFICACION,CALIFICACIONINFO,NIVEL,ASUNTO,NOMBRE,LUGAR,AUTOR,DESCRIPCION,PLACA,ID,FUENTE,CALIFICACIONFUENTE,ANEXOS,CODIGO,TELEFONO,SEGUIMIENTO,FOTO,DIRECCION,MARCA,COLOR,TELFUENTE,AUDITORIA)" 
sql = sql &" values('" & Fecha & "','" & Provincia & "','" & Ciudad & "','" & Estado & "','"
sql = sql & Clasificacion & "','" & CalificacionInfo & "','" & Nivel & "','" & Asunto & "','" & Nombre & "','" & Lugar & "','"
sql = sql & Autor & "','" & Descripcion & "','" & Placa & "','" & Id & "','" & Fuente & "','" & CalificacionFuente & "','" & Anexos & "'," & Codigo & ",'" & Telefono & "','" & Seguimiento & "','" & Foto & "','" & Direccion & "','" & Marca & "','" & Color & "','" & Telfuente & "','" & Auditoria & "')"
RS.Open sql, Conex, 1
Conex.Close
Set RS = Nothing
Set Conex = Nothing

'response.write sql

response.redirect "updaterecord.asp?action=update&Codigo=" & Codigo

end if
%>


	  </table>
	  
	  <hr>
		  <p class="footer" align="center">
		  <!-- ########### INCLUDE FOOTER ########### Contact is specific to site-->
		  <!-- #include file="includes/contact.htm" --><br>
		  <!-- #include virtual = "/includes/htm/footer.htm" -->
		  </p>

</body>

</html>