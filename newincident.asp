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
<script language="javascript" src="includes/cal2.js"></script>
<script language="javascript" src="includes/cal_conf2.js"></script>

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
Hora = Request.Form("Hora")
Provincia = Request.Form("Provincia")
Ciudad = Request.Form("Ciudad")
Clasificacion = Request.Form("Clasificacion")
Tipo = Request.Form("Tipo")
Tipo1 = Request.Form("Tipo1")
Nivel = Request.Form("Nivel")
Lugar = Request.Form("Lugar")
Autor = Request.Form("Autor")
Descripcion = Request.Form("Descripcion")
Costo = Request.Form("Costo")
Personasaec = Request.Form("Personasaec")
Personascontract = Request.Form("Personascontract")
Personasthird = Request.Form("Personasthird")
Vehiculos = Request.Form("Vehiculos")
Placa = Request.Form("Placa")
Fuente = Request.Form("Fuente")
Fuente1 = Request.Form("Fuente1")
'Por la eliminacion del campo Fuente1
Fuente1 = Fuente
Clasificacionfuente = Request.Form("Clasificacionfuente")
Lideres = Request.Form("Lideres")
Organizacion = Request.Form("Organizacion")
Radio = Request.Form("R1")
Status = Request.Form("Status")
Anexos = Request.Form("Anexos")
Referencia = Request.Form("Referencia")
Razon = Request.Form("Razon")
Seguimiento = Request.Form("Seguimiento")
Auditoria = UserID & " - " & Now()

Reportadoa = Request.Form("Reportadoa")
Acciones = Request.Form("Acciones")
Danos = Request.Form("Danos")
Material = Request.Form("Material")
Requieres = Request.Form("Requieres")
Requierei = Request.Form("Requierei")

if action = "" then
%>
<p>&nbsp;</p>

<p class="banner_blue">Crear un Nuevo Incidente</p>

<form method="POST" name="FrontPage_Form1" action="newincident.asp?action=save" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" >
<table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="652" height="745">
  <tr>
    <td width="84" height="22"><span lang="es-ec">Referencia :</span></td>
    <td width="526" height="22" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="300" --><input name="Referencia" type="text"  size="80" maxlength="300"></td>
  </tr>

  <tr>
    <td width="84" height="22">Fecha :</td>
    <td width="176" height="22">
    &nbsp;
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" s-allow-other-chars="/-" b-value-required="TRUE" i-maximum-length="11" --><input name="Fecha" type="text" onBlur="checkdate(this)" size="12" maxlength="11" readonly="true">
    <a href="javascript:showCal('FechaCal');"><img border="0" src="http://apps01.andespetro.com/includes/Calendar/date.gif" align="middle" width="19" height="17" alt="Select a Date"></a>    
    </td>
    <td width="80" height="22" align="right">
    <p align="left">Hora :</td>
    <td width="112" height="22">
    <!--webbot bot="Validation" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars=":" b-value-required="TRUE" i-maximum-length="5" --><input name="Hora" type="text"  size="10" maxlength="5"></td>
    <td width="79" height="22">
    <p align="right">Status :</td>
    <td width="79" height="22">
    <select size="1" name="STATUS">
    <option>Abierto</option>
    <option>Cerrado</option>
    </select></td>
  </tr>

  <tr>
    <td width="84" height="32"><span lang="es-ec">Tipo Incidente:</span></td>
    <td width="176" height="32">
    <select size="1" name="Tipo">
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
    <option>Atentado terrorista</option>
    <option>Bloqueo cierre de vías</option>    
    <option>Bomba panfletaria</option>
    <option>Choque/accidente vehicular</option>    
    <option>Derrame Combustible</option>
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
    <option>Presencia armada grupos ilegales</option>
    <option>Presencia persona sospechosa</option>
    <option>Raptos</option>
    <option>Robo accesorios de vehículos</option>
    <option>Robo cable cobre</option>
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
    <td width="80" height="32" align="right">
    Gravedad :</td>
    <td width="270" height="32" colspan="3">
    <select size="1" name="Nivel">
    <option>Alto</option>
    <option>Medio</option>
    <option>Bajo</option>
    </select></td>
  </tr>

  <tr>
    <td width="84" height="32">Provincia :</td>
    <td width="176" height="32">
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
    <td width="80" height="32" align="right">
    <p>Sector :</td>
    <td width="270" height="32" colspan="3">
    <select size="1" name="Ciudad">
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
	<option>Centro Biaña</option>
	<option>Centro Eno</option>
	<option>Chiritza</option>
	<option>Ciudad de Piñas</option>
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
	<option>Estrella del Aguarico</option>
	<option>Dureno</option>
	<option>General Farfán</option>
	<option>Guataraco</option>
	<option>Guayllabamba</option>
	<option>Guayaquil</option>
	<option>Hormigueros</option>
	<option>Ibarra</option>
	<option>Jesús del Gran Poder</option>
	<option>Kupi</option>
	<option>La Bermeja</option>
	<option>La Bonita</option>
	<option>La Independencia</option>
	<option>La Joya de los Sachas</option>
	<option>Las Mercedes</option>
	<option>Lago Agrio</option>
	<option>Latacunga</option>
	<option>Limoncocha</option>
	<option>Loja</option>
	<option>Loreto</option>
	<option>Lucha Bolivarense</option>
	<option>Lumbaqui</option>
	<option>Macas</option>
	<option>Machala</option>
	<option>Manta</option>
	<option>Nantus</option>
	<option>Narupa</option>
	<option>Nueva Frontera</option>
	<option>Nueva Manabí</option>
	<option>Nueva Santa Rosa</option>
	<option>Nueva Vida</option>
	<option>Otavalo</option>
	<option>Pacayacu</option>
	<option>Palma Roja</option>
	<option>Pañacocha</option>
	<option>Papallacta</option>
	<option>Parque Nacional Yasuni(secor CDP)</option>
	<option>Paz y Bien</option>
	<option>Pifo</option>
	<option>Pindo</option>
	<option>Pinuña Negra</option>
	<option>Pompeya</option>
	<option>Poza Honda</option>
	<option>Puerto El Carmen</option>
	<option>Putumayo</option>
	<option>Pto. Nuevo</option>
	<option>Puerto Quito</option>
	<option>Puerto Siney</option>
	<option>Puyo</option>
	<option>Quevedo</option>
	<option>Quininde</option>
	<option>Quito 1</option>
	<option>Quito 2</option>
	<option>Quito 3</option>
	<option>Quito 4</option>
	<option>Restrepo</option>
	<option>Rey de los Andes</option>
	<option>Riera</option>
	<option>Rodrigo Borja</option>
	<option>San Carlos</option>
	<option>Sangolquí</option>
	<option>San José</option>
	<option>San Miguel de Los Bancos</option>
	<option>Sansahuari</option>
	<option>San Pablo de Kantesiaya</option>
	<option>San Sebastián del Coca</option>
	<option>San Roque</option>
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
	<option>Tapir</option>
	<option>Taracea</option>
	<option>Tarapoa</option>
	<option>Tase</option>
	<option>Tena</option>
	<option>Tetetes</option>
	<option>Tierras Orientales</option>
	<option>Tiguano</option>
	<option>Tipishca</option>
	<option>Tiputini</option>
	<option>Tobeta</option>
	<option>Tulcán</option>
	<option>Valle de Tumbaco</option>
	<option>Valle de los Chillos</option>
	<option>Zamona</option>
    </select></td>
  </tr>

  <tr>
    <td width="84" height="52">Lugar Específico:</td>
    <td width="526" height="52" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="500" --><textarea rows="3" name="Lugar" cols="68"></textarea></td>
  </tr>

  <tr>
    <td width="84" height="32"><span lang="es-ec">Elaborado por :</span></td>
    <td width="526" height="32" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="250" --><textarea rows="1" name="Razon" cols="68"></textarea></td>
  </tr>

  <tr>
    <td width="84" height="32"><span lang="es-ec">Reportado a :</span></td>
    <td width="526" height="32" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="250" --><input name="Reportadoa"  type="text"  size="29" maxlength="250"></td>
  </tr>

  <tr>
    <td width="84" height="32">Reportado por :</td>
    <td width="176" height="32">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="250" --><input name="Fuente" type="text"  size="29" maxlength="250"></td>
    <td width="80" height="32" align="right">
    C<span lang="es-ec">onfiabilidad</span> Fuente :</td>
    <td width="270" height="32" colspan="3">
    <select size="1" name="Clasificacionfuente">
    <option>No confiable</option>
    <option>Casi Confiable</option>
    <option>Confiable</option>
    </select></td>
  </tr>
  <tr>
    <td width="84" height="32">Clasificación&nbsp; Información:</td>
    <td width="176" height="32">
    <select size="1" name="Clasificacion">
    <option>Abierta</option>
    <option>Reservada</option>
    <option>Secreta</option>
    <option>Secretísima</option>
    </select></td>
    <td width="80" height="32" align="right">
    <p>Calificación&nbsp; Información<span lang="es-ec"> :</span></td>
    <td width="270" height="32" colspan="3">
    <select size="1" name="Tipo1">
    <option>Casi Confiable</option>
    <option>Confiable</option>
    <option>No confiable</option>
    <option>Rumor</option>    
    <option>Verídica</option>
    </select></td>
  </tr>
  <tr>
    <td width="652" align="left" height="1" colspan="6">
    <hr>
    <p></td>
  </tr>
  <tr>
    <td width="84" height="52">Descripción :</td>
    <td width="568" height="52" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_.:" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2000" --><textarea rows="3" name="Descripcion" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Notificación/Seguimiento :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_" i-maximum-length="2000" --><textarea rows="2" name="Seguimiento" cols="68"></textarea></td>
  </tr>
  
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Acciones Tomadas :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_" i-maximum-length="2000" --><textarea rows="2" name="Acciones" cols="68"></textarea></td>
  </tr>
  
  <tr>
    <td width="84" height="36">Personas AdPetro :</td>
    <td width="568" height="36" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="_#-" i-maximum-length="2000" --><textarea rows="2" name="PersonasAEC" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" height="48">Personas Contratistas :</td>
    <td width="568" height="48" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_" i-maximum-length="2000" --><textarea rows="2" name="Personascontract" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36">Personas Terceros :</td>
    <td width="568" height="36" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#" i-maximum-length="2000" --><textarea rows="2" name="Personasthird" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" height="36">Vehiculos :</td>
    <td width="568" height="36" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="-#_" i-maximum-length="2000" --><textarea rows="2" name="Vehiculos" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    <p>Placa :</td>
    <td width="568" align="left" height="22" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_" i-maximum-length="250" --><input name="Placa" type="text"  size="80" maxlength="250"></td>
  </tr>
  
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Daños Ocasionados :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_" i-maximum-length="2000" --><textarea rows="2" name="Danos" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" align="left" height="31">
    <span lang="es-ec">Dueño Equipo/Material :</span></td>
    <td width="568" align="left" height="31" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="#-_" i-maximum-length="2000" --><textarea rows="2" name="Material" cols="68"></textarea></td>
  </tr>
  <tr>
    <td width="84" height="32">Requiere Seguimiento :</td>
    <td width="176" height="32">
    <select size="1" name="Requieres">
    <option>Si</option>
    <option>No</option>
    </select></td>
    <td width="80" height="32" align="right">
    <p>Requiere Investigación<span lang="es-ec"> :</span></td>
    <td width="270" height="32" colspan="3">
    <select size="1" name="Requierei">
    <option>Si</option>
    <option>No</option>
    </select></td>
  </tr>
      
  <tr>
    <td width="84" align="left" height="22">
    Anexos :</td>
    <td width="568" align="left" height="22" colspan="5">
    <!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" b-allow-whitespace="TRUE" s-allow-other-chars="." i-maximum-length="2500" --><textarea rows="2" name="Anexos" cols="63"></textarea><input type="button" onclick="gettext()" value="Ver" name="B1"></td>
  </tr>
  <tr>
    <td width="84" align="left" height="22">
    Costo :</td>
    <td width="568" align="left" height="22" colspan="5">
    <!--webbot bot="Validation" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="15" s-validation-constraint="Greater than or equal to" s-validation-value="0" --><input name="Costo" type="text"  size="10" maxlength="15"></td>
  </tr>
  </table>
  <p><input type="submit" value="Grabar" name="Add"></p>
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
sql = "select seq_incident_code.NextVal from Dual"
RS.Open sql, Conex, 1
Codigo = RS.Fields("NEXTVAL")
Conex.Close
Set RS = Nothing
Set Conex = Nothing

'Grabo los datos
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "insert into incident values('" & Fecha & "','" & Hora & "','" & Provincia & "','" & Ciudad & "','" & Clasificacion & "','"
sql = sql & Tipo & "','" & Tipo1 & "','" & Nivel & "','" & Lugar & "','none','" & Descripcion & "',"
sql = sql & Costo & ",'" & Personasaec & "','" & Personascontract & "','" & Personasthird & "','"
sql = sql & Vehiculos & "','" & Placa & "','" & Fuente & "','" & Fuente1 & "','" & Clasificacionfuente & "','" & Lideres & "','"
sql = sql & Organizacion & "','" & Radio & "','" & Status & "','" & Anexos & "'," & Codigo & ",'" & Referencia & "','" & Razon & "','" & Seguimiento & "','" & Auditoria & "','"
sql = sql & Reportadoa & "','" & Acciones & "','" & Danos & "','" & Material & "','" & Requieres & "','" & Requierei & "')"


RS.Open sql, Conex, 1
Conex.Close
Set RS = Nothing
Set Conex = Nothing

response.redirect "asociarrecord.asp?Incidente=" & Codigo

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