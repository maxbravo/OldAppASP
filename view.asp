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

<% 
If UserID="EREY" or  UserID = "MAVILA" or UserID="MBRAVO" or UserID="MSAONA" or UserID="JPAZMINO" or UserID="JABAD" or UserID="GBRAVO" or UserID="CMAFLA" or UserID="" or UserID="GOBANDO1" or UserID="PEGAS"  or UserID="JCASTELL"  or UserID="JMORALES"  or UserID="FGARCIA1"  or UserID="YLAMAR"  or UserID="JECHEVER"  or UserID = "CVILLAGO" or UserID = "MMOLINA" Then
%>


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
	sql = "select * from INCIDENT where CODIGO = " & Cod
	RS.Open sql, Conex, 1

	Fecha = RS.Fields("FECHA")
	Hora = RS.Fields("HORA")
	Provincia = RS.Fields("PROVINCIA")
	Ciudad = RS.Fields("CIUDAD")
	Clasificacion = RS.Fields("CLASIFICACION")
	Tipo = RS.Fields("TIPO")
	Tipo1 = RS.Fields("TIPO1")
	Nivel = RS.Fields("NIVEL")
	Lugar = RS.Fields("LUGAR")
	Autor = RS.Fields("AUTOR")
	Descripcion = RS.Fields("DESCRIPCION")
	Costo = RS.Fields("COSTO")
	Personasaec = RS.Fields("PERSONASAEC")
	Personascontract = RS.Fields("PERSONASCONTRACT")
	Personasthird = RS.Fields("PERSONASTHIRD")
	Vehiculos = RS.Fields("VEHICULOS")
	Placa = RS.Fields("PLACA")
	Fuente = RS.Fields("FUENTE")
	Fuente1 = RS.Fields("FUENTE1")
	Clasificacionfuente = RS.Fields("CLASIFICACIONFUENTE")
	Razon = RS.Fields("RAZON")
	Seguimiento = RS.Fields("SEGUIMIENTO")
	Status = RS.Fields("STATUS")
	Anexos = RS.Fields("ANEXOS")
	Referencia = RS.Fields("REFERENCIA")

	Reportadoa = RS.Fields("REPORTADOA")
	Acciones = RS.Fields("ACCIONES")
	Danos = RS.Fields("DANOS")
	Material = RS.Fields("MATERIAL")
	Requieres = RS.Fields("REQUIERES")
	Requierei = RS.Fields("REQUIEREI")

	Conex.Close
	Set RS = Nothing
	Set Conex = Nothing

%>
<p>&nbsp;</p>

<p class="banner_blue">Ver Incidente Código : <%=Cod%></p>

<form method="POST" name="FrontPage_Form1" action="updateincident.asp?action=save&Codigo=<%=Cod%>" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" >
<table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="652">
  <tr>
    <td width="84" height="22"><b>Referencia :</b></td>
    <td width="526" height="22" colspan="5"><%=Referencia%></td>
</td>
  </tr>

  <tr>
    <td width="84" height="22"><b>Fecha :</b></td>
    <td width="189" height="22"><%=Fecha%></td>
    <td width="100" height="22" align="right">
    <p><b>Hora :</b></td>
    <td width="79" height="22"><%=Hora%></td>
    <td width="79" height="22">
    <b>Status :</b></td>
    <td width="79" height="22"><%=Status%></td>
  </tr>

  <tr>
    <td width="84" height="22"><b>Clasificación Incidente :</b></td>
    <td width="189" height="22"><%=Tipo%></td>
    <td width="100" height="22" align="right">
    <p><b><span lang="es-ec">Gravedad</span> :</b></td>
    <td width="237" height="22" colspan="3"><%=Nivel%></td>
  </tr>

  <tr>
    <td width="84" height="22"><b>Provincia :</b></td>
    <td width="189" height="22"><%=Provincia%></td>
    <td width="100" height="22" align="right">
    <p><span lang="es-ec"><b>Sector</b></span><b> :</b></td>
    <td width="237" height="22" colspan="3"><%=Ciudad%></td>
  </tr>

  <tr>
    <td width="84" height="22"><b>Lugar Específico:</b></td>
    <td width="526" height="22" colspan="5"><%=Lugar%></td>
  </tr>

  <tr>
    <td width="84" height="22"><span lang="es-ec"><b>Elaborado por :</b></span></td>
    <td width="526" height="22" colspan="5"><%=Razon%></td>
  </tr>

  <tr>
    <td width="84" height="22"><span lang="es-ec"><b>Reportado a :</b></span></td>
    <td width="526" height="22" colspan="5"><%=Reporatadoa%></td>
  </tr>

  <tr>
    <td width="84" height="22">
    <b>Reportado por :</b></td>
    <td width="189" height="22"><%=Fuente%></td>
    <td width="100" height="22" align="right">
    <b>C<span lang="es-ec">onfiabilidad</span> Fuente :</b></td>
    <td width="237" height="22" colspan="3"><%=Clasificacion%></td>
  </tr>

  <tr>
    <td width="84" height="22"><b>Clasificación&nbsp; Información:</b></td>
    <td width="189" height="22"><%=Clasificacion%></td>
    <td width="100" height="22" align="right">
    <p><b>Calificación&nbsp; Información :</b></td>
    <td width="237" height="22" colspan="3"><%=Tipo1%></td>
  </tr>
  <tr>
    <td width="610" height="22" colspan="6"><hr></td>
  </tr>
  <tr>
    <td width="84" height="36"><b>Descripción :</b></td>
    <td width="568" height="36" colspan="5"><%=Descripcion%></td>
  </tr>
  <tr>
    <td width="84" height="22">
    <b><span lang="es-ec">Notificación</span> <span lang="es-ec">Seguimiento :</span></b></td>
    <td width="568" height="22" colspan="5"><%=Seguimiento%></td>
  </tr>
  <tr>
    <td width="84" height="22">
    <b><span lang="es-ec">Acciones Tomadas:</span></b></td>
    <td width="568" height="22" colspan="5"><%=Acciones%></td>
  </tr>
  <tr>
    <td width="84" height="36"><b>Personas AdPetro :</b></td>
    <td width="568" height="36" colspan="5"><%=PersonasAEC%></td>
  </tr>
  <tr>
    <td width="84" height="36"><b>Personas Contratistas :</b></td>
    <td width="568" height="36" colspan="5"><%=Personascontract%></td>
  </tr>
  <tr>
    <td width="84" height="36"><b>Personas Terceros :</b></td>
    <td width="568" height="36" colspan="5"><%=Personasthird%>
    &nbsp;</td>
  </tr>
  <tr>
    <td width="84" height="36"><b>Vehiculos :</b></td>
    <td width="568" height="36" colspan="5"><%=Vehiculos%></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <p><b>Placa :</b></td>
    <td width="568" align="left" height="22" colspan="5"><%=Placa%></td>
  </tr>
  <tr>
    <td width="84" height="22">
    <b><span lang="es-ec">Daños Ocasionados :</span></b></td>
    <td width="568" height="22" colspan="5"><%=Danos%></td>
  </tr>
  <tr>
    <td width="84" height="22">
    <b><span lang="es-ec">Dueño Equipo/Material :</span></b></td>
    <td width="568" height="22" colspan="5"><%=Material%></td>
  </tr>
  <tr>
    <td width="84" height="22">
    <b>Requiere Seguimiento :</b></td>
    <td width="189" height="22"><%=Requieres%></td>
    <td width="100" height="22" align="right">
    <b><span lang="es-ec">Requiere Investigación</span> :</b></td>
    <td width="237" height="22" colspan="3"><%=Requierei%></td>
  </tr>  
  <tr>
    <td width="84" align="left" height="22">
    <b>Anexos :</b></td>
    <td width="568" align="left" height="22" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="." i-maximum-length="500" --><textarea  rows="2" name="Anexos" cols="62"><%=Anexos%></textarea><input type="button" onclick="gettext()" value="Ver" name="B1"></td>
  </tr>
  <tr>
    <td width="84" align="left" height="14">
    <b>Costo :</b></td>
    <td width="568" align="left" height="14" colspan="5"><%=Costo%></td>
  </tr>
  </table>
  <p>&nbsp;</p>
	  </table>

<%
	'Busco Records Asociados
	Set Conex = Server.CreateObject ("ADODB.Connection")
	Conex.Open ConString
	Set RS = Server.CreateObject("ADODB.RecordSet")
	sql = "select * from record where codigo in(select record from incidenterecord where incidente = " & Cod & ")"
	RS.Open sql, Conex, 1
%>

&nbsp;<p class="banner_blue">
Records Asociados al Incidente Código: <%=Cod%></p>
<table border="1" cellspacing="0" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-width: 0; border-collapse:collapse" bordercolor="#111111" cellpadding="0">
  <tr>
    <td width="10%">
    <p class="banner_green">Código</td>
    <td width="20%">
    <p class="banner_green">Fecha</td>
    <td width="550%">
    <p class="banner_green">Nombre</td>
    <td width="5%">
    <p align="center" class="banner_green">View</td>
  </tr>
<%
	Do While not RS.EOF
%>
  <tr>
    <td bgcolor="#E7FEE0" width="10%"><%=RS.Fields("CODIGO")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" width="20%"><%=RS.Fields("FECHA")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" width="55%"><%=RS.Fields("Nombre")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" align="center" width="10%"><a target="_blank" href="viewr.asp?action=view&codigo=<%=RS.Fields("CODIGO")%>">
    <img src="images/view.gif" border="0"></a></td>
  </tr>
<%
	RS.MoveNext
	Loop
%>

</table>
<%
	Conex.Close
	Set RS = Nothing
	Set Conex = Nothing
end if
%>
	  
	  <hr>

<%
End If
%>  
</body>

</html>