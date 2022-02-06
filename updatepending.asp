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
	  myRemote = launch("https://uioap009/SITS/docs/temp/"+rng.text, "SITS", "height=400,width=400,channelmode=0,dependent=0,directories=1,fullscreen=0,location=1,menubar=0,resizable=1,scrollbars=0,status=0,toolbar=0", "myWindow");      
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
    <td valign="top" width="8"><!-- ########### INCLUDE Sample Navigation ########### --><!-- #include file="includes/navigation2.htm" --></td>
	
	
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
sql = "select * from INCIDENT where CODIGO = " & Cod
RS.Open sql, Conex, 1

Fecha = CambioFecha(RS.Fields("FECHA"))
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
'Por la eliminacion de este campo
Fuente1 = Fuente 
Clasificacionfuente = RS.Fields("CLASIFICACIONFUENTE")
Lideres = RS.Fields("LIDERES")
Organizacion = RS.Fields("ORGANIZACION")
Radio = RS.Fields("RADIO")
Status = RS.Fields("STATUS")
Anexos = RS.Fields("ANEXOS")
Referencia = RS.Fields("REFERENCIA")
Razon = RS.Fields("Razon")
Seguimiento = RS.Fields("Seguimiento")
Auditoria = RS.Fields("Auditoria")

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
<b><p align="center"><i><font size="5" color="#009933">INCIDENT TRACKING SYSTEM</font></i></p>
</b>

<p class="banner_blue">Modificar Incidente de Código : <%=Cod%> - Last Update: <%=Auditoria%></p>

<form method="POST" name="FrontPage_Form1" action="updatepending.asp?action=save&Codigo=<%=Cod%>" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" >
<table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="660" height="585">
  <tr>
    <td width="84" height="22"><span lang="es-ec">Referencia :</span></td>
    <td width="526" height="22" colspan="5">
    <!--webbot bot="Validation" s-display-name="Referencia" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="300" --><input name="Referencia" type="text"  size="80" maxlength="300" value="<%=Referencia%>"></td>
  </tr>
  <tr>
    <td width="84" height="22">Fecha :</td>
    <td width="176" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars="-/" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="11" --><input name="Fecha" type="text" onBlur="checkdate(this)" size="11" maxlength="11" value="<%=Fecha%>" tabindex="1" readonly="true">
    <a href="javascript:showCal('FechaCalU');"><img border="0" src="http://apps01.andespetro.com/includes/Calendar/date.gif" align="middle" width="19" height="17" alt="Select a Date"></a>    
    </td>
    <td width="93" height="22" align="right">
    <p>Hora :</td>
    <td width="78" height="22">
    <!--webbot bot="Validation" s-display-name="Hora" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars=":" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="5" --><input name="Hora" type="text"  size="10" maxlength="5" value="<%=Hora%>"></td>
    <td width="108" height="22">
    <p align="right">Status :</td>
    <td width="79" height="22">
    <select size="1" name="STATUS">
    <option selected><%=Status%></option>
    <option>Borrador</option>
    <option>Validado</option>
    </select></td>
  </tr>

  <tr>
    <td width="84" height="32"><span lang="es-ec">Tipo Incidente :</span></td>
    <td width="176" height="32">
    <select size="1" name="Tipo">
    <option selected><%=Tipo%></option>
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
    <td width="93" height="32" align="right">
    <p><span lang="es-ec">Gravedad</span> :</td>
    <td width="265" height="32" colspan="3">
    <select size="1" name="Nivel">
    <option selected><%=Nivel%></option>
    <option>Alto</option>
    <option>Medio</option>
    <option>Bajo</option>
    </select></td>
  </tr>

  <tr>
    <td width="84" height="32">Provincia :</td>
    <td width="176" height="32">
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
    <td width="93" height="32" align="right">
    <p><span lang="es-ec">Sector</span> :</td>
    <td width="265" height="32" colspan="3">
    <select size="1" name="Ciudad">
    <option selected><%=Ciudad%></option>    
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
  </tr>

  <tr>
    <td width="84" height="52">Lugar Específico:</td>
    <td width="526" height="52" colspan="5">
    <!--webbot bot="Validation" s-display-name="Lugar" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="500" --><textarea rows="3" name="Lugar" cols="68"><%=Lugar%></textarea></td>
  </tr>

  <tr>
    <td width="84" height="32"><span lang="es-ec">Elaborado por :</span></td>
    <td width="526" height="32" colspan="5">
    <!--webbot bot="Validation" s-display-name="Elaborado Por" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="250" --><textarea rows="1" name="Razon" cols="68"><%=Razon%></textarea></td>
  </tr>

  <tr>
    <td width="84" height="32"><span lang="es-ec">Reportado a :</span></td>
    <td width="526" height="32" colspan="5">
    <!--webbot bot="Validation" s-display-name="Reportado A" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="250" --><input name="Reportadoa" type="text"  size="29" maxlength="250" value="<%=Reportadoa%>"></td>
  </tr>

  <tr>
    <td width="84" height="32">Reportado por :</td>
    <td width="176" height="32">
    <!--webbot bot="Validation" s-display-name="Reportado Por" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="250" --><input name="Fuente" type="text"  size="29" maxlength="250" value="<%=Fuente%>"></td>
    <td width="93" height="32" align="right">
    C<span lang="es-ec">onfiabilidad</span> Fuente :</td>
    <td width="265" height="32" colspan="3">
    <select size="1" name="Clasificacionfuente">
    <option selected><%=Clasificacionfuente%></option>
    <option>No confiable</option>
    <option>Casi Confiable</option>
    <option>Confiable</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" height="32">Clasificación&nbsp; Información:</td>
    <td width="176" height="32">
    <select size="1" name="Clasificacion">
    <option selected><%=Clasificacion%></option>
    <option>Abierta</option>
    <option>Reservada</option>
    <option>Secreta</option>
    <option>Secretísima</option>
    </select></td>
    <td width="93" height="32" align="right">
    Calificación&nbsp; Información <span lang="es-ec">:</span></td>
    <td width="265" height="32" colspan="3">
    <select size="1" name="Tipo1">
    <option selected><%=Tipo1%></option>
    <option>Casi Confiable</option>
    <option>Confiable</option>
    <option>No confiable</option>
    <option>Rumor</option>
    <option>Verídica</option>
    </select></td>
  </tr>
  <tr>
    <td width="652" align="left" height="32" colspan="6">
    <hr>
    <p></td>
  </tr>
  <tr>
    <td width="84" height="52">Descripción :</td>
    <td width="568" height="52" colspan="5">
    <!--webbot bot="Validation" s-display-name="Descripción" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_." b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2000" --><textarea rows="3" name="Descripcion" cols="68"><%=Descripcion%></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Notificación/Seguimiento :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-display-name="Notificación Seguimiento" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="2000" --><textarea rows="2" name="Seguimiento" cols="68"><%=Seguimiento%></textarea></td>
  </tr>
  
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Acciones Tomadas :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-display-name="Acciones Tomadas" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="2000" --><textarea rows="2" name="Acciones" cols="68"><%=Acciones%></textarea></td>
  </tr>   
  
  <tr>
    <td width="84" height="36">Personas AdPetro :</td>
    <td width="568" height="36" colspan="5">
    <!--webbot bot="Validation" s-display-name="Personas Andes Petroleum" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="2000" --><textarea rows="2" name="PersonasAEC" cols="68"><%=PersonasAEC%></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36">Personas Contracts :</td>
    <td width="568" height="36" colspan="5">
    <!--webbot bot="Validation" s-display-name="Personas Contratistas" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" i-maximum-length="2000" --><textarea rows="2" name="Personascontract" cols="68"><%=Personascontract%></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36">Personas Third :</td>
    <td width="568" height="36" colspan="5">
    <!--webbot bot="Validation" s-display-name="Personas Terceras" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="2000" --><textarea rows="2" name="Personasthird" cols="68"><%=Personasthird%></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36">Vehiculos :</td>
    <td width="568" height="36" colspan="5">
    <!--webbot bot="Validation" s-display-name="Vehiculos" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_" i-maximum-length="2000" --><textarea rows="2" name="Vehiculos" cols="68"><%=Vehiculos%></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <p>Placa :</td>
    <td width="568" align="left" height="22" colspan="5">
    <!--webbot bot="Validation" s-display-name="Placa" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#_" i-maximum-length="250" --><input name="Placa" type="text"  size="80" maxlength="250" value="<%=Placa%>"></td>
  </tr>
  
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Daños Ocasionados :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-display-name="Daños Ocasionados" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="2000" --><textarea rows="2" name="Danos" cols="68"><%=Danos%></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Dueño Equipo/Material :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-display-name="Dueño Equipo Material" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-" i-maximum-length="2000" --><textarea rows="2" name="Material" cols="68"><%=Material%></textarea></td>
  </tr>    
  <tr>
    <td width="84" height="32">Requiere Seguimiento :</td>
    <td width="176" height="32">
    <select size="1" name="Requieres">
    <option selected><%=Requieres%></option>
    <option>Si</option>
    <option>No</option>
    </select></td>
    <td width="93" height="32" align="right">
    Requiere Investigación <span lang="es-ec">:</span></td>
    <td width="265" height="32" colspan="3">
    <select size="1" name="Requierei">
    <option selected><%=Requierei%></option>
    <option>Si</option>
    <option>No</option>
    </select></td>
  </tr>    
  
  <tr>
    <td width="84" align="left" height="16">
    Anexos :</td>
    <td width="568" align="left" height="16" colspan="5">
    <!--webbot bot="Validation" s-display-name="Anexos" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="." i-maximum-length="2500" --><textarea rows="2" name="Anexos" cols="63"><%=Anexos%></textarea><input type="button" onclick="gettext()" value="Ver" name="B1"><br>
    <a target="_blank" href="file://uioap009/SITStemp$">copiar archivos anexos</a></td>
  </tr>
  <tr>
    <td width="84" height="22">Costo :</td>
    <td width="568" height="22" colspan="5">
    <!--webbot bot="Validation" s-display-name="Costo" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="15" --><input name="Costo" type="text"  size="10" maxlength="15" value="<%=Costo%>"></td>
  </tr>
  </table>
  <p><input type="submit" value="Grabar" name="Add"></p>
</form>
<form method="POST" action="asociarrecord.asp?Incidente=<%=Cod%>" name="Recor">
&nbsp;</form>


<%
end if
%>
<%
if action = "save" then

'Obtengo los valores del formulario
Fecha = Request.Form("FECHA")
Hora = Request.Form("HORA")
Provincia = Request.Form("PROVINCIA")
Ciudad = Request.Form("CIUDAD")
Clasificacion = Request.Form("CLASIFICACION")
Tipo = Request.Form("TIPO")
Tipo1 = Request.Form("TIPO1")
Nivel = Request.Form("NIVEL")
Lugar = Request.Form("LUGAR")
Autor = Request.Form("AUTOR")
Descripcion = Request.Form("DESCRIPCION")
Costo = Request.Form("COSTO")
Personasaec = Request.Form("PERSONASAEC")
Personascontract = Request.Form("PERSONASCONTRACT")
Personasthird = Request.Form("PERSONASTHIRD")
Vehiculos = Request.Form("VEHICULOS")
Placa = Request.Form("PLACA")
Fuente = Request.Form("FUENTE")
Fuente1 = Request.Form("FUENTE1")
Clasificacionfuente = Request.Form("CLASIFICACIONFUENTE")
Lideres = Request.Form("LIDERES")
Organizacion = Request.Form("ORGANIZACION")
Radio = Request.Form("R1")
Status = Request.Form("STATUS")
Anexos = Request.Form("ANEXOS")
Referencia = Request.Form("REFERENCIA")
Razon = Request.Form("Razon")
Seguimiento = Request.Form("Seguimiento")
Auditoria = UserID & " - " & Now()

Reportadoa = Request.Form("Reportadoa")
Acciones = Request.Form("Acciones")
Danos = Request.Form("Danos")
Material = Request.Form("Material")
Requieres = Request.Form("Requieres")
Requierei = Request.Form("Requierei")

'Grabo los datos
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "update INCIDENT set "
sql = sql & "FECHA = '" & Fecha & "',"
sql = sql & "HORA = '" & Hora & "',"
sql = sql & "PROVINCIA = '" & Provincia & "',"
sql = sql & "CIUDAD = '" & Ciudad & "',"
sql = sql & "CLASIFICACION = '" & Clasificacion & "',"
sql = sql & "TIPO = '" & Tipo & "',"
sql = sql & "TIPO1 = '" & Tipo1 & "',"
sql = sql & "NIVEL = '" & Nivel & "',"
sql = sql & "LUGAR = '" & Lugar & "',"
sql = sql & "DESCRIPCION = '" & Descripcion & "',"
sql = sql & "COSTO = " & Costo & ","
sql = sql & "PERSONASAEC = '" & Personasaec & "',"
sql = sql & "PERSONASCONTRACT = '" & Personascontract & "',"
sql = sql & "PERSONASTHIRD = '" & Personasthird & "',"
sql = sql & "VEHICULOS = '" & Vehiculos & "',"
sql = sql & "PLACA = '" & Placa & "',"
sql = sql & "FUENTE = '" & Fuente & "',"
sql = sql & "FUENTE1 = '" & Fuente1 & "',"
sql = sql & "CLASIFICACIONFUENTE = '" & Clasificacionfuente & "',"
sql = sql & "RAZON = '" & Razon & "',"
sql = sql & "AUDITORIA = '" & Auditoria & "',"
sql = sql & "SEGUIMIENTO = '" & Seguimiento & "',"
sql = sql & "STATUS = '" & Status & "',"

sql = sql & "REPORTADOA = '" & Reportadoa & "',"
sql = sql & "ACCIONES = '" & Acciones & "',"
sql = sql & "DANOS = '" & Danos & "',"
sql = sql & "MATERIAL = '" & Material & "',"
sql = sql & "REQUIERES = '" & Requieres & "',"
sql = sql & "REQUIEREI = '" & Requierei & "',"

sql = sql & "ANEXOS = '" & Anexos & "', REFERENCIA='" & Referencia & "' where CODIGO = " & Cod

	if Status = "Validado" then

		'ENVIO DE EMAIL PARA EDUARDO
	   strHost = "mailhost2.andespetro.com"
	   Set Mail = Server.CreateObject("Persits.MailSender")
	   ' enter valid SMTP host
	   Mail.Host = strHost
	   Mail.From = "ECSecuritySITS@andespetro.com"
	   Mail.AddAddress "erey@andespetro.com"
	   ' message subject
	   Mail.Subject = "SITS: A New Incident has been registered."
	   ' message body
	   Mail.IsHTML = True
	   body = "A New Incident has been registered.<br><b>Reference: </b> " & Referencia & "<br><b>Date: </b>" & Fecha & " " & Hora & "<br><br><a target='_blank' href=https://uioap009/sits/updateincident.asp?action=update&codigo=" & Cod & ">If you want to validate this draft, please click here.</a>"
	   Mail.Body = body
	   ' send message
	   Mail.Send 
	   Set Mail = Nothing

		'ENVIO DE EMAIL PARA TODOS
	   strHost = "mailhost2.andespetro.com"
	   Set Mail = Server.CreateObject("Persits.MailSender")
	   ' enter valid SMTP host
	   Mail.Host = strHost
	   Mail.From = "ECSecuritySITS@andespetro.com"
	   Mail.AddAddress "ECSecuritySITS@andespetro.com"
	   ' message subject
	   Mail.Subject = "SITS: A New Incident has been registered."
	   ' message body
	   Mail.IsHTML = True
	   body = "A New Incident has been registered.<br><b>Reference: </b> " & Referencia & "<br><b>Date: </b>" & Fecha & " " & Hora & "<br><b>Localion: </b> " & Provincia & " - " & Ciudad & "<br><b>Registered by: </b> " & Auditoria & "<br><br><a target='_blank' href=https://uioap009/sits/view.asp?action=view&codigo=" & Cod & ">If you want to view this new incident, please click here.</a>"
	   Mail.Body = body
	   ' send message
	   Mail.Send 
	   Set Mail = Nothing

	 End If

RS.Open sql, Conex, 1
Conex.Close
Set RS = Nothing
Set Conex = Nothing

'response.write Status
response.redirect "indexpending.asp"

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