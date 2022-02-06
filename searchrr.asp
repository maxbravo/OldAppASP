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
</SCRIPT>
<link rel="stylesheet" type="text/css" href="http://apps01.andespetro.com/includes/css/default.css">
</head>

<body>	  
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="860">
    <tr>
    <td valign="top"><!-- ########### INCLUDE Sample Navigation ########### --><!-- #include file="includes/navigation.htm" --></td>
	
	
    <td bgcolor="#FFFFFF" valign="top" align="left"><font size="4" color="#000080">&nbsp;&nbsp;</font></td>
    <td valign="top" width="860">


</SCRIPT>
<link rel="stylesheet" type="text/css" href="http://apps01.andespetro.com/includes/css/default.css">
<script language="javascript" src="includes/cal2.js"></script>
<script language="javascript" src="includes/cal_conf2.js"></script>
<script language="javascript" src="includes/ValidarFrmProject.js"></script>
<title>Project Tracking Tool</title>
<meta http-equiv="Content-Language" content="es-ec">
<link rel="stylesheet" type="text/css" href="http://apps01.andespetro.com/includes/css/default.css">

<%
Dim sql, Conex, RS, valida, Project_Code
action = request("action")

Cod = Request("Codigo")

Codigo = Request.Form("Codigo")
Provincia = Request.Form("Provincia")
Ciudad = Request.Form("Ciudad")
Estado = Request.Form("Estado")
Nombre = Request.Form("Nombre")
ID = Request.Form("ID")
Lugar = Request.Form("Lugar")
Clasificacion = Request.Form("Clasificacion")
CalificacionInfo = Request.Form("CalificacionInfo")
Nivel = Request.Form("Nivel")
Placa = Request.Form("Placa")
Telefono = Request.Form("Telefono")
FechaD = Request.Form("FechaD")
FechaH = Request.Form("FechaH")
Asunto = Request.Form("Asunto")
Fuente = Request.Form("Fuente")
TelFuente = Request.Form("TelFuente")
Descripcion = Request.Form("Descripcion")

if action = "" then
%>
<p>&nbsp;</p>

<p class="banner_blue">Buscar Record Por:</p>


<form method="POST" action="searchrr.asp?action=search" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
  <table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="652">
  <tr>
    <td width="84" height="32"><span lang="es-ec">Fecha Desde :</span></td>
    <td width="189" height="32">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars="/-" i-maximum-length="11" --><input name="FechaD" type="text" onBlur="checkdate(this)" size="12" maxlength="11" readonly="true">
    <a href="javascript:showCal('FechaCalD');"><img border="0" src="http://apps01.andespetro.com/includes/Calendar/date.gif" align="middle" width="19" height="17" alt="Select a Date"></a>    
    </td>
    <td width="100" height="32" align="right">
    <span lang="es-ec">Fecha Hasta :</span></td>
    <td width="237" height="32">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars="/-" i-maximum-length="11" --><input name="FechaH" type="text" onBlur="checkdate(this)" size="12" maxlength="11" readonly="true">
    <a href="javascript:showCal('FechaCalH');"><img border="0" src="http://apps01.andespetro.com/includes/Calendar/date.gif" align="middle" width="19" height="17" alt="Select a Date"></a>    
    </td>
  </tr>
  <tr>
    <td width="84" height="22">Provincia :</td>
    <td width="189" height="22">
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
    <option selected>All</option>
    </select></td>
    <td width="100" height="22" align="right">
    <p>Clasificación :</td>
    <td width="237" height="22">
    <select size="1" name="Clasificacion">
    <option selected>All</option>
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
    <td width="189" height="22">
    <input name="Ciudad" type="text"  size="20" maxlength="50"></td>
    <td width="100" height="22" align="right">
    <p>Idoneidad :</td>
    <td width="237" height="22">
    <select size="1" name="CalificacionInfo">
    <option selected>All</option>
    <option>Idóneo</option>
    <option>No Idóneo</option>
    <option>No Confiable</option>
    <option>Rumor</option>
    <option>Casi Confiable</option>
    <option>Confiable</option>
    <option>Verídica</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" height="22">Estado :</td>
    <td width="189" height="22">
    <select size="1" name="ESTADO">
    <option selected>All</option>
    <option>Abierto</option>
    <option>Cerrado</option>
    <option>En Proceso</option>
    </select></td>
    <td width="100" height="22" align="right">
    <p>Nivel :</td>
    <td width="237" height="22">
    <select size="1" name="Nivel">
    <option>Bajo</option>
    <option>Medio</option>
    <option>Alto</option>
    <option selected>All</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" height="22"><span lang="es-ec">Nombre : </span></td>
    <td width="189" height="22">
    <input name="Nombre" type="text"  size="20" maxlength="50"></td>
    <td width="100" height="22" align="right">
    <span lang="es-ec">Placa :</span></td>
    <td width="237" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="PLACA" type="text"  size="15" maxlength="150"></td>
  </tr>
  <tr>
    <td width="84" height="22"><span lang="es-ec">ID :</span></td>
    <td width="189" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="ID" type="text"  size="15" maxlength="150"></td>
    <td width="100" height="22" align="right">
    <span lang="es-ec">Teléfono :</span></td>
    <td width="237" height="22">
    <!--webbot bot="Validation" s-display-name="Telefono" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="Telefono" type="text"  size="15" maxlength="150"></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    Fuente :</td>
    <td width="568" align="left" height="22" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="200" --><input name="FUENTE" type="text"  size="80" maxlength="200"></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <span lang="es-ec">Tel Fuente : </span></td>
    <td width="171" align="left" height="22" colspan="3">
    <!--webbot bot="Validation" s-display-name="Telefono" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-" i-maximum-length="150" --><input name="TelFuente" type="text"  size="15" maxlength="150"></td>
  </tr>
  <tr>
    <td width="84" height="22"><span lang="es-ec">Lugar :</span></td>
    <td width="526" height="22" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" i-maximum-length="2500" --><textarea rows="3" name="LUGAR" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="89" height="22"><span lang="es-ec">Descripción :</span></td>
    <td width="526" height="22" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" i-maximum-length="2500" --><textarea rows="3" name="DESCRIPCION" cols="68"></textarea></td>
  </tr>    
  </table>
  
  <p><input type="submit" value="Buscar" name="B1"></p>
</form>

<%
end if
%>
<%
if action = "search" then
'Update de los datos
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "select CODIGO,FECHA,NOMBRE from RECORD where 1=1 "
if Nombre <> "" Then
 sql = sql & " and UPPER(NOMBRE) like '%" & UCase(Nombre) & "%' " 
End If
if Ciudad <> "" Then
 sql = sql & " and UPPER(CIUDAD) like '%" & UCase(Ciudad) & "%' " 
End If
if Provincia <> "All" then
 sql = sql & " and PROVINCIA ='" & Provincia & "' " 
end if
if Estado <> "All" then
 sql = sql & " and ESTADO ='" & Estado & "' " 
end if
if Clasificacion  <> "All" then
 sql = sql & " and CLASIFICACION = '" & Clasificacion & "' " 
end if
if CalificacionInfo  <> "All" then
 sql = sql & " and CALIFICACIONINFO = '" & CalificacionInfo & "' "    
end if
if Nivel  <> "All" then
 sql = sql & " and NIVEL = '" & Nivel & "' " 
end if
if Placa  <> "" then
 sql = sql & " and PLACA = '" & Placa & "' " 
end if
if ID  <> "" then
 sql = sql & " and ID = '" & ID & "' " 
end if
if Telefono  <> "" then
 sql = sql & " and TELEFONO = '" & Telefono & "' " 
end if
if Lugar  <> "" then
 sql = sql & " and LUGAR = '" & Lugar & "' " 
end if
if (FechaD  <> "" and FechaH <> "") then
 sql = sql & " and FECHA >= '" & FechaD & "' and FECHA <= '" & FechaH & "' "
end if
if Asunto <> "" Then
 sql = sql & " and UPPER(ASUNTO) like '%" & UCase(Asunto) & "%' " 
End If
if Fuente <> "" Then
 sql = sql & " and UPPER(FUENTE) like '%" & UCase(Fuente) & "%' " 
End If
if TelFuente <> "" Then
 sql = sql & " and UPPER(TELFUENTE) like '%" & UCase(TelFuente) & "%' " 
End If
if Descripcion <> "" Then
 sql = sql & " and UPPER(DESCRIPCION) like '%" & UCase(Descripcion) & "%' " 
End If
RS.Open sql, Conex, 1
%>


&nbsp;<p class="banner_blue">
Resultados de la Búsqueda</p>
<table border="1" cellspacing="0" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-width: 0; border-collapse:collapse" bordercolor="#111111" cellpadding="0">
  <tr>
    <td width="20%">
    <p class="banner_green">Código</td>
    <td width="20%">
    <p class="banner_green">Fecha</td>
    <td width="50%">
    <p class="banner_green">Nombre</td>
    <td width="10%">
    <p align="center" class="banner_green">Ver</td>
  </tr>
<%
Do While not RS.EOF
%>
  <tr>
    <td bgcolor="#E7FEE0" width="20%"><%=RS.Fields("CODIGO")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" width="20%"><%=RS.Fields("FECHA")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" width="50%"><%=RS.Fields("Nombre")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" align="center" width="10%"><a target="_blank" href="viewr.asp?action=view&codigo=<%=RS.Fields("CODIGO")%>">
    <img src="images/view.gif" border="0"></a></td>
  </tr>
<%
RS.MoveNext
Loop
%>

</table>
<p>
<%
Conex.Close
Set RS = Nothing
Set Conex = Nothing
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