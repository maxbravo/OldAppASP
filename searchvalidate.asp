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
<script language="javascript" src="includes/fechas.js"></script>
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
Clasificacion = Request.Form("Clasificacion")
Tipo = Request.Form("Tipo")
Tipo1 = Request.Form("Tipo1")
Nivel = Request.Form("Nivel")
Referencia = Request.Form("Referencia")
Incidente = Request.Form("Incidente")
Status = Request.Form("Status")
Descripcion = Request.Form("Descripcion")
Lugar = Request.Form("Lugar")
Personas = Request.Form("Personasaec")
Vehiculos = Request.Form("Vehiculos")
Placa = Request.Form("Placa")
FechaD = Request.Form("FechaD")
FechaH = Request.Form("FechaH")

if action = "" then
%>


<p class="banner_blue">Buscar Incidente Por:</p>


<form method="POST" action="searchvalidate.asp?action=search" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
  <table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="652" height="241">
  <tr>
    <td width="84" height="32"><span lang="es-ec">No. Incidente :</span></td>
    <td width="189" height="32">
    <!--webbot bot="Validation" s-data-type="Integer" s-number-separators="x" i-maximum-length="5" --><input name="Incidente" type="text"  size="5" maxlength="5"></td>
    <td width="100" height="32" align="right">
    <span lang="es-ec">Status :</span></td>
    <td width="237" height="32">
    <select size="1" name="STATUS">
    <option selected>Validado</option>
    </select></td>
  </tr>

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
    <td width="84" height="32">Provincia :</td>
    <td width="189" height="32">
    <select size="1" name="Provincia">
    <option>Azuay</option>
    <option selected>All</option>
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
    <td width="100" height="32" align="right">
    <p>Clasificación del Incidente :</td>
    <td width="237" height="32">
    <select size="1" name="Tipo">
    <option selected>All</option>
    <option>Abuso de confianza</option>
    <option>Accidente de transito</option>
    <option>Acciones activistas</option>
    <option>Acoso sexual</option>
    <option>Agresión física</option>
    <option>Amenaza de bomba</option>
    <option>Amenazas otras</option>
    <option>Apoyo a ejecutivo</option>
    <option>Asalto armado a ejecutivo</option>
    <option>Asalto bancos</option>
    <option>Asalto C. Comercial</option>
    <option>Asalto carretera</option>
    <option>Asalto y robo vehículos</option>
    <option>Asalto y robo</option>
    <option>Asociación ilícita</option>
    <option>Atentado al pudor</option>
    <option>Bloqueo cierre de vías</option>    
    <option>Bomba panfletaria</option>
    <option>Choque/accidente vehicular</option>    
    <option>Derrame Petroleo</option>
    <option>Destrucción bienes</option>
    <option>Disparo accidental</option>
    <option>Disparo arma de fuego</option>
    <option>Disturbios</option>
    <option>Emergencia médica</option>
    <option>Estupro</option>
    <option>Evasión</option>
    <option>Extorsión</option>
    <option>Falsificación</option>
    <option>Fraude</option>
    <option>Heridas / Lesiones</option>
    <option>Homicidios</option>
    <option>Huelga</option>
    <option>Hurtos</option>
    <option>Impedimento de patrullaje</option>    
    <option>Incautación</option>
    <option>Incendio</option>
    <option>Instalaciones eléctricas</option>
    <option>Intento a entrar en red</option>
    <option>Intento de asalto</option>
    <option>Intento de robo</option>
    <option>Intento de robo de cobre</option>
    <option>Intento de soborno</option>    
    <option>Intento homicidio</option>
    <option>Intento ingreso locación</option>
    <option>Intento secuestro</option>
    <option>Intento violación</option>
    <option>Intimidación / Amenaza</option>
    <option>Invasiones</option>
    <option>Llamada Telefono Sospechosa</option>
    <option>Mal uso de información</option>
    <option>Manifestaciones</option>
    <option>Muerte</option>
    <option>Objeto Sospechoso</option>
    <option>Otros robos</option>	
    <option>Paquetes Sospechoso</option>
    <option>Paralización de trabajos</option>        
	  <option>Paros</option>
    <option>Peculado</option>
	  <option>Pérdida</option>
    <option>Persona desaparecida</option>
    <option>Presencia persona sospechosa</option>
    <option>Raptos</option>
    <option>Robo accesorios de vehículos</option>
    <option>Robo de información</option>
    <option>Robo de tubería</option>
    <option>Robo domicilio</option>
    <option>Robo vehículo</option>
    <option>Robo</option>
    <option>Saboteo Oleoducto</option>
    <option>Secuestros</option>
    <option>Seguimiento sospechoso</option>
    <option>Simulacros</option>
    <option>Soborno</option>
    <option>Sospecha de Robo / Hurto</option>
    <option>Suicidio</option>
    <option>Tenencia explosivos</option>
    <option>Tenencia ilícita de armas</option>
    <option>Tenencia moneda falsa</option>
    <option>Trata de blancas</option>
    <option>Varios</option>
    <option>Violaciones</option>
    </select></td>
  </tr>

  <tr>
    <td width="84" height="32"><span lang="es-ec">Sector</span> :</td>
    <td width="189" height="32">
    <select size="1" name="Ciudad">
    <option selected>All</option>
	<option>Aguas Negras</option>
	<option>Alóag</option>
	<option>Ambato</option>
	<option>Archidona</option>
	<option>Auca</option>
	<option>Baeza</option>
	<option>Baños</option>
	<option>Buenos Amigos</option>
	<option>Campamento Reventador</option>
	<option>Cantagallo</option>
	<option>Cayambe</option>
	<option>Chiritza</option>
	<option>Cononaco</option>
	<option>Cosanga</option>
	<option>Cotundo</option>
	<option>Cuembi</option>
	<option>Cuenca</option>
	<option>Cuyabeno</option>
	<option>Dayuma</option>
	<option>El Chaco</option>
	<option>El Coca</option>
	<option>El Dorado de Cascales</option>
	<option>El Eno</option>
	<option>Esmeraldas</option>
	<option>Estación Bombeo El Salado</option>
	<option>Dureno</option>
	<option>General Farfán</option>
	<option>Guataraco</option>
	<option>Guayllabamba</option>
	<option>Guayaquil</option>
	<option>Hormigueros</option>
	<option>Ibarra</option>
	<option>Kupi</option>
	<option>La Bermeja</option>
	<option>La Bonita</option>
	<option>La Independencia</option>
	<option>La Joya de los Sachas</option>
	<option>Lago Agrio</option>
	<option>Latacunga</option>
	<option>Limoncocha</option>
	<option>Loja</option>
	<option>Loreto</option>
	<option>Lumbaqui</option>
	<option>Macas</option>
	<option>Machala</option>
	<option>Manta</option>
	<option>Nantus</option>
	<option>Narupa</option>
	<option>Nueva Frontera</option>
	<option>Nueva Santa Rosa</option>
	<option>Otavalo</option>
	<option>Pacayacu</option>
	<option>Palma Roja</option>
	<option>Pañacocha</option>
	<option>Papallacta</option>
	<option>Paz y Bien</option>
	<option>Pifo</option>
	<option>Pindo</option>
	<option>Pinuña Negra</option>
	<option>Pompeya</option>
	<option>Poza Honda</option>
	<option>Puerto El Carmen</option>
	<option>Pto. Nuevo</option>
	<option>Puerto Quito</option>
	<option>Putumayo</option>
	<option>Puyo</option>
	<option>Quevedo</option>
	<option>Quininde</option>
	<option>Quito 1</option>
	<option>Quito 2</option>
	<option>Quito 3</option>
	<option>Quito 4</option>
	<option>Restrepo</option>
	<option>Riera</option>
	<option>Rodrigo Borja</option>
	<option>San Carlos</option>
	<option>Sangolquí</option>
	<option>San José</option>
	<option>San Miguel de Los Bancos</option>
	<option>Sansahuari</option>
	<option>San Sebastián del Coca</option>
	<option>Sarayacu</option>
	<option>Sardinas</option>
	<option>Santa Elena</option>
	<option>Santa Rosa</option>
	<option>Santo Domingo</option>
	<option>Shiripuno</option>
	<option>Shell</option>
	<option>Shushufindi</option>
	<option>Singüe</option>
	<option>Tabacundo</option>
	<option>Taracea</option>
	<option>Tarapoa</option>
	<option>Tase</option>
	<option>Tena</option>
	<option>Tetetes</option>
	<option>Tiguano</option>
	<option>Tipishca</option>
	<option>Tiputini</option>
	<option>Tulcán</option>
	<option>Valle de Tumbaco</option>
	<option>Valle de los Chillos</option>
    </select></td>
    <td width="100" height="32" align="right">
    <p>Calificación&nbsp; Información :</td>
    <td width="237" height="32">
    <select size="1" name="Tipo1">
    <option>No confiable</option>
    <option selected>All</option>
    <option>Rumor</option>
    <option>Casi Confiable</option>
    <option>Confiable</option>
    <option>Verídica</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" height="48">Clasificación de la Información<span lang="es-ec">
    </span>:</td>
    <td width="189" height="48">
    <select size="1" name="Clasificacion">
    <option>Abierta</option>
    <option selected>All</option>
    <option>Reservada</option>
    <option>Secreta</option>
    <option>Secretísima</option>
    </select></td>
    <td width="100" height="48" align="right">
    <p><span lang="es-ec">Gravedad</span> :</td>
    <td width="237" height="48">
    <select size="1" name="Nivel">
    <option selected>All</option>
    <option>Alto</option>
    <option>Medio</option>
    <option>Bajo</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" height="53">Descripción :</td>
    <td width="568" height="53" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="2000" --><textarea rows="3" name="Descripcion" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" height="52">Lugar :</td>
    <td width="568" height="52" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" i-maximum-length="500" --><textarea rows="3" name="Lugar" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36">Personas :</td>
    <td width="568" height="36" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" i-maximum-length="2000" --><textarea rows="2" name="PersonasAEC" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36">Vehiculos :</td>
    <td width="568" height="36" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" i-maximum-length="2000" --><textarea rows="2" name="Vehiculos" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <p>Placa :</td>
    <td width="568" align="left" height="22" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="250" --><input name="Placa" type="text"  size="80" maxlength="250"></td>
  </tr>
  <tr>
    <td width="84" height="22"><span lang="es-ec">Referencia :</span></td>
    <td width="526" height="22" colspan="3">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" i-maximum-length="300" --><input name="Referencia" type="text"  size="80" maxlength="300"></td>
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
sql = "select CODIGO,FECHA,HORA,REFERENCIA from INCIDENT where 1=1 " 
if Ciudad <> "All" then
 sql = sql & " and CIUDAD ='" & Ciudad & "' " 
end if
if Provincia <> "All" then
 sql = sql & " and PROVINCIA ='" & Provincia & "' " 
end if
if Clasificacion <> "All" then
 sql = sql & " and CLASIFICACION ='" & Clasificacion & "' " 
end if
if Tipo  <> "All" then
 sql = sql & " and TIPO = '" & Tipo & "' " 
end if
if Tipo1  <> "All" then
 sql = sql & " and TIPO1 = '" & Tipo1 & "' "    
end if
if Nivel  <> "All" then
 sql = sql & " and NIVEL = '" & Nivel & "' " 
end if
if Referencia  <> "" then
 sql = sql & " and UPPER(REFERENCIA) like '%" & UCase(Referencia) & "%' " 
end if
if Incidente  <> "" then
 sql = sql & " and CODIGO = '" & Incidente & "' "    
end if

 sql = sql & " and STATUS = 'Validado' "    

if Descripcion  <> "" then
 sql = sql & " and UPPER(DESCRIPCION) like '%" & UCase(Descripcion) & "%' " 
end if
if Lugar  <> "" then
 sql = sql & " and UPPER(LUGAR) like '%" & UCase(Lugar) & "%' " 
end if
if Vehiculos  <> "" then
 sql = sql & " and UPPER(VEHICULOS) like '%" & UCase(Vehiculos) & "%' " 
end if
if Placa  <> "" then
 sql = sql & " and UPPER(PLACA) like '%" & UCase(Placa) & "%' " 
end if 
if FechaD  <> "" then
 sql = sql & " and FECHA >= '" & FechaD & "' " 
end if
if FechaH  <> "" then
 sql = sql & " and FECHA <= '" & FechaH & "' " 
end if
if Personas  <> "" then
 sql = sql & " and ( UPPER(PERSONASAEC) like '%" & UCase(Personas) & "%' " 
 sql = sql & " or UPPER(PERSONASCONTRACT) like '%" & UCase(Personas) & "%' " 
 sql = sql & " or UPPER(PERSONASTHIRD) like '%" & UCase(Personas) & "%') " 
end if 

RS.Open sql, Conex, 1
%>


&nbsp;<p class="banner_blue">
Resultados de la Búsqueda</p>
<table border="1" cellspacing="0" width="100%" id="AutoNumber1" bordercolorlight="#FFFFFF" bordercolordark="#FFFFFF" style="border-width: 0; border-collapse:collapse" bordercolor="#111111" cellpadding="0">
  <tr>
    <td width="55%">
    <p class="banner_green">Referencia</td>
    <td width="15%">
    <p class="banner_green">Fecha</td>
    <td width="10%">
    <p class="banner_green">Hora</td>
    <td width="10%">
    <p align="center" class="banner_green">Edit</td>
    <td width="10%">
    <p align="center" class="banner_green">Delete</td>
  </tr>
<%
Do While not RS.EOF
%>
  <tr>
    <td bgcolor="#E7FEE0" width="55%"><%=RS.Fields("REFERENCIA")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" width="15%"><%=RS.Fields("FECHA")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" width="10%"><%=RS.Fields("HORA")%>&nbsp;</td>
    <td bgcolor="#E7FEE0" align="center" width="10%"><a target="_blank" href="updateincident.asp?action=update&codigo=<%=RS.Fields("CODIGO")%>">
    <img src="images/edit.gif" border="0"></a></td>
    <td bgcolor="#E7FEE0" align="center" width="10%"><a href="searchvalidate.asp?action=delete&codigo=<%=RS.Fields("CODIGO")%>">
    <img src="images/delete.gif" border="0"></a></td>
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
<%
if action = "delete1" then
'Update de los datos
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "delete INCIDENT where CODIGO = " & Cod
RS.Open sql, Conex, 1
Conex.Close
Set RS = Nothing
Set Conex = Nothing
response.redirect "index.asp"
end if
%>
<%
if action = "delete" then
%>
<form method="POST" action="searchvalidate.asp?action=delete1&codigo=<%=Cod%>">
&nbsp;<p>Está seguro que desea eliminar el record de incidente número <%=Cod%> ?</p>
<p>&nbsp;<input type="submit" value="Eliminar" name="Eliminar"> </p>
</form>
<%
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