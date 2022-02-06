<!-- #include file="includes/Param.asp" -->
<!-- #include file="includes/CambioFecha.asp" -->

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
	  myRemote = launch("https://uioap009/SITS/docs/"+rng.text, "SITS", "height=400,width=400,channelmode=0,dependent=0,directories=1,fullscreen=0,location=1,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0", "myWindow");      
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

Cod = Request("Codigo")

if action = "update" then

'Obtengo los datos de este incidente
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "select * from RECORD where CODIGO = " & Cod
RS.Open sql, Conex, 1

Fecha = CambioFecha(RS.Fields("FECHA"))
Provincia = RS.Fields("PROVINCIA")
Ciudad = RS.Fields("CIUDAD")
Estado = RS.Fields("ESTADO")
Clasificacion = RS.Fields("CLASIFICACION")
CalificacionInfo = RS.Fields("CALIFICACIONINFO")
Nivel = RS.Fields("NIVEL")
Asunto = RS.Fields("ASUNTO")
Nombre = RS.Fields("NOMBRE")
Lugar = RS.Fields("LUGAR")
Autor = RS.Fields("AUTOR")
Descripcion = RS.Fields("DESCRIPCION")
Placa = RS.Fields("PLACA")
ID = RS.Fields("ID")
Fuente = RS.Fields("FUENTE")
CalificacionFuente = RS.Fields("CALIFICACIONFUENTE")
Anexos = RS.Fields("ANEXOS")
Telefono = RS.Fields("TELEFONO")
Seguimiento = RS.Fields("SEGUIMIENTO")
Foto = RS.Fields("FOTO")
Direccion = RS.Fields("DIRECCION")
Marca = RS.Fields("MARCA")
Color = RS.Fields("COLOR")
TelFuente = RS.Fields("TELFUENTE")
Auditoria = RS.Fields("AUDITORIA")


Conex.Close
Set RS = Nothing
Set Conex = Nothing
%>
<p>&nbsp;</p>

<p class="banner_blue">Actualizar Record Codigo: <%=Cod%> - Last Update: <%=Auditoria%></p>

<form method="POST" name="FrontPage_Form1" action="updaterecord.asp?action=save&Codigo=<%=Cod%>" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" >
<table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="652">
  <tr>
    <td width="84" height="22">Fecha :</td>
    <td width="189" height="22" colspan="2">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars="/-" b-value-required="TRUE" i-minimum-length="10" i-maximum-length="11" --><input name="FECHA" type="text" onChange="checkdate(this)" onBlur="checkdate(this)" size="12" maxlength="11" value="<%=Fecha%>" readnoly="true">
    <a href="javascript:showCal('FechaCalRU');"><img border="0" src="http://apps01.andespetro.com/includes/Calendar/date.gif" align="middle" width="19" height="17" alt="Select a Date"></a>    
    </td>
    <td width="337" height="22" align="right" colspan="5">&nbsp;
    </td>
  </tr>

  <tr>
    <td width="84" height="22">Provincia :</td>
    <td width="189" height="22" colspan="2">
    <select size="1" name="Provincia">
    <option selected><%=Provincia%></option>
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
    <td width="100" height="22" align="right" colspan="3">
    <p>Clasificación :</td>
    <td width="237" height="22" colspan="2">
    <select size="1" name="CLASIFICACION">
    <option selected><%=Clasificacion%></option>
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
  </tr>
  <tr>
    <td width="84" height="22">Cuidad :</td>
    <td width="189" height="22" colspan="2">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="50" --><input name="Ciudad" type="text"  size="20" maxlength="50" value="<%=Ciudad%>"></td>
    <td width="100" height="22" align="right" colspan="3">
    <p><span lang="es-ec">Idoneidad</span> :</td>
    <td width="237" height="22" colspan="2">
    <select size="1" name="CalificacionInfo">
    <option selected><%=CalificacionInfo%></option>
    <option>Idóneo</option>
    <option>No Idóneo</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" height="22">Estado :</td>
    <td width="189" height="22" colspan="2">
    <select size="1" name="ESTADO">
    <option selected><%=Estado%></option>
    <option>Abierto</option>
    <option>Cerrado</option>
    <option>En Proceso</option>
    </select></td>
    <td width="100" height="22" align="right" colspan="3">
    <p>Nivel :</td>
    <td width="237" height="22" colspan="2">
    <select size="1" name="NIVEL">
    <option selected><%=Nivel%></option>
    <option>Bajo</option>
    <option>Medio</option>
    <option>Alto</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    Historical Record :</td>
    <td width="568" align="left" height="22" colspan="7">
    <!--webbot bot="Validation" s-display-name="Seguimiento" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" i-maximum-length="2500" --><textarea rows="3" name="Seguimiento" cols="68"><%=Seguimiento%></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36"><span lang="es-ec">Nombre</span> :</td>
    <td width="568" height="36" colspan="7">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="500" --><input name="NOMBRE" type="text"  size="80" maxlength="500" value="<%=Nombre%>"></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    ID : </td>
    <td width="114" align="left" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="ID" type="text"  size="15" maxlength="150" value="<%=ID%>"></td>
    <td width="114" align="left" height="22" colspan="2">
    Teléfono : </td>
    <td width="114" align="left" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="Telefono" type="text"  size="15" maxlength="150" value="<%=Telefono%>"></td>
    <td width="6" align="left" height="22" colspan="2">
    <a target="_blank" href="docs/<%=Foto%>">Foto</a> : </td>
    <td width="220" align="left" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-." i-maximum-length="150" --><input name="Foto" type="text"  size="15" maxlength="150" value="<%=Foto%>"></td>
  </tr>
  <tr>
    <td width="84" height="36">Dirección :</td>
    <td width="568" height="36" colspan="7">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#_" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2500" --><textarea rows="3" name="DIRECCION" cols="68"><%=Direccion%></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    Marca : </td>
    <td width="114" align="left" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="Marca" type="text"  size="15" maxlength="150" value="<%=Marca%>"></td>
    <td width="114" align="left" height="22" colspan="2">
    Color : </td>
    <td width="114" align="left" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="Color" type="text"  size="15" maxlength="150" value="<%=Color%>"></td>
    <td width="6" align="left" height="22" colspan="2">
    Placa : </td>
    <td width="220" align="left" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="PLACA" type="text"  size="15" maxlength="150" value="<%=Placa%>"></td>
  </tr> 
  <tr>
    <td width="84" height="36"><span lang="es-ec">Descripción :</span></td>
    <td width="568" height="36" colspan="7">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#_:." b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2500" --><textarea rows="3" name="DESCRIPCION" cols="68"><%=Descripcion%></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    Fuente :</td>
    <td width="568" align="left" height="22" colspan="7">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="200" --><input name="FUENTE" type="text"  size="80" maxlength="200" value="<%=Fuente%>"></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <span lang="es-ec">Tel Fuente : </span></td>
    <td width="171" align="left" height="22" colspan="3">
    <!--webbot bot="Validation" s-display-name="Telefono" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="TelFuente" type="text"  size="15" maxlength="150" value="<%=TelFuente%>"></td>
    <td width="198" align="left" height="22" colspan="3">
    Lugar <span lang="es-ec">de Contacto </span>:</td>
    <td width="199" align="left" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#._-@!:;,/\?=)(%" i-maximum-length="2500" --><textarea rows="3" name="LUGAR" cols="25"><%=Lugar%></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    Calificación Fuente :</td>
    <td width="568" align="left" height="22" colspan="7">
    <select size="1" name="CalificacionFuente">
    <option selected><%=CalificacionFuente%></option>
    <option>Confiable</option>
    <option>Casi Confiable</option>
    <option>No Confiable</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    Anexos :</td>
    <td width="568" align="left" height="22" colspan="7">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="." i-maximum-length="2500" --><textarea rows="2" name="Anexos" cols="63"><%=Anexos%></textarea><input type="button" onclick="gettext()" value="Ver" name="B1"></td>
  </tr>
  </table>
  <p><input type="submit" value="Grabar" name="Add"></p>
</form>


<%
end if
%>
<%
if action = "save" then

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
Autor = Request.Form("Autor")
Descripcion = Request.Form("Descripcion")
Placa = Request.Form("Placa")
Id = Request.Form("Id")
Fuente = Request.Form("Fuente")
CalificacionFuente = Request.Form("CalificacionFuente")
Anexos = Request.Form("Anexos")
Telefono = Request.Form("Telefono")
Seguimiento = Request.Form("Seguimiento")
Foto = Request.Form("FOTO")
Direccion = Request.Form("DIRECCION")
Marca = Request.Form("MARCA")
Color = Request.Form("COLOR")
TelFuente = Request.Form("TELFUENTE")
Auditoria = UserID & " - " & Now()

'Grabo los datos
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "update RECORD set FECHA='" & Fecha & "',PROVINCIA='" & Provincia & "',CIUDAD='" & Ciudad & "',ESTADO='" & Estado & "',CLASIFICACION='"
sql = sql & Clasificacion & "',CALIFICACIONINFO='" & CalificacionInfo & "',NIVEL='" & Nivel & "',ASUNTO='" & Asunto & "',NOMBRE='" & Nombre & "',LUGAR='" & Lugar & "',AUTOR='"
sql = sql & Autor & "',DESCRIPCION='" & Descripcion & "',PLACA='" & Placa & "',ID='" & Id & "',FUENTE='" & Fuente & "',CALIFICACIONFUENTE='" & CalificacionFuente & "',ANEXOS='" & Anexos & "',TELEFONO='" & Telefono & "',SEGUIMIENTO='" & Seguimiento & "',FOTO='" & Foto & "',DIRECCION='" & Direccion & "',MARCA='" & MARCA & "',COLOR='" & Color & "',TELFUENTE='" & TelFuente & "',AUDITORIA='" & Auditoria &"' WHERE CODIGO =" & Cod
RS.Open sql, Conex, 1
Conex.Close
Set RS = Nothing
Set Conex = Nothing

'response.write sql
response.redirect "searchrecord.asp"

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