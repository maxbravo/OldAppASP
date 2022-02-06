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

if action = "view" then

'Obtengo los datos de este incidente
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "select * from RECORD where CODIGO = " & Cod
RS.Open sql, Conex, 1

Fecha = RS.Fields("FECHA")
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
Anexo = RS.Fields("ANEXOS")
Telefono = RS.Fields("TELEFONO")
Seguimiento = RS.Fields("SEGUIMIENTO")
Foto = RS.Fields("FOTO")
Direccion = RS.Fields("DIRECCION")
Marca = RS.Fields("MARCA")
Color = RS.Fields("COLOR")
TelFuente = RS.Fields("TELFUENTE")


Conex.Close
Set RS = Nothing
Set Conex = Nothing

%>
<p>&nbsp;</p>

<p class="banner_blue">Ver Incidente Record : <%=Cod%></p>

<form method="POST" name="FrontPage_Form1" action="updateincident.asp?action=save&Codigo=<%=Cod%>" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" >
<table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="652">
  <tr>
    <td width="84" height="22"><b>Fecha :</b></td>
    <td width="189" height="22"><%=Fecha%></td>
    <td width="337" height="22" align="right" colspan="2">
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" height="22"><b>Provincia :</b></td>
    <td width="189" height="22"><%=Provincia%>
    &nbsp;</td>
    <td width="100" height="22" align="right">
    <p><b>Clasificación :</b></td>
    <td width="237" height="22"><%=Clasificacion%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" height="22"><b>Cuidad :</b></td>
    <td width="189" height="22"><%=Ciudad%>
    &nbsp;</td>
    <td width="100" height="22" align="right">
    <p><b>Calificación Información :</b></td>
    <td width="237" height="22"><%=CalificacionInfo%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" height="22"><b>Estado :</b></td>
    <td width="189" height="22"><%=Estado%></td>
    <td width="100" height="22" align="right">
    <p><b>Nivel :</b></td>
    <td width="237" height="22"><%=Estado%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" height="36"><b><span lang="es-ec">Nombre</span> :</b><p>
    <span lang="es-ec"></span></td>
    <td width="568" height="36" colspan="3"><%=Nombre%> (<a target="_blank" href="docs/<%=Foto%>">Foto</a>)
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <b>
    <span lang="es-ec">ID : </span></b></td>
    <td width="568" align="left" height="22" colspan="3"><%=ID%> &nbsp;</td>
  </tr>
  <tr>
    <td width="84" height="36"><b>Dirección :</b></td>
    <td width="568" height="36" colspan="3"><%=Direccion%></td>
  </tr>
  <tr>
    <td width="84" height="36"><b>Lugar :</b></td>
    <td width="568" height="36" colspan="3"><%=Lugar%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" height="36"><b><span lang="es-ec">Descripción :</span></b></td>
    <td width="568" height="36" colspan="3"><%=Descripcion%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <span lang="es-ec"><b>Marca :</b></span></td>
    <td width="568" align="left" height="22" colspan="3"><%=Marca%></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <span lang="es-ec"><b>Color :</b></span></td>
    <td width="568" align="left" height="22" colspan="3"><%=Color%></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <b>
    <span lang="es-ec">Placa : </span></b></td>
    <td width="568" align="left" height="22" colspan="3"><%=Placa%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <b>
    <span lang="es-ec">Teléfono : </span></b></td>
    <td width="568" align="left" height="22" colspan="3"><%=TELEFONO%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <b>Fuente :</b></td>
    <td width="568" align="left" height="22" colspan="3"><%=Fuente%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <span lang="es-ec"><b>Teléfono Fuente : </b></span></td>
    <td width="568" align="left" height="22" colspan="3"><%=TelFuente%></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <b>Calificación Fuente :</b></td>
    <td width="568" align="left" height="22" colspan="3"><%=CalificacionFuente%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <b>Historical Record :</b></td>
    <td width="568" align="left" height="22" colspan="3"><%=SEGUIMIENTO%>
    &nbsp;</td>
  </tr>  
  <tr>
    <td width="84" align="left" height="22">
    <b>Anexos :</b></td>
    <td width="568" align="left" height="22" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="." i-maximum-length="2500" --><textarea rows="2" name="Anexos" cols="63"><%=Anexo%></textarea><input type="button" onclick="gettext()" value="Ver" name="B1"></td>
  </tr>
  </table>
  <p>&nbsp;</p>
</form>


<%
end if
%>


	  </table>
	  
	  <hr>


</body>

</html>