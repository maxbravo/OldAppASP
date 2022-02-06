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
    <p class="sub_header">&nbsp;</p>

    <p class="sub_header">Security Incident Tracking System - OLAP</p>

    <p class="banner_green">Incidentes</p>

    <p class="sub_header">
    <object classid="clsid:0002E552-0000-0000-C000-000000000046" id="PivotTableDeprecated1">
      <param name="XMLData" value="&lt;xml xmlns:x=&quot;urn:schemas-microsoft-com:office:excel&quot;&gt;
 &lt;x:PivotTable&gt;
  &lt;x:OWCVersion&gt;10.0.0.2621         &lt;/x:OWCVersion&gt;
  &lt;x:DisplayScreenTips/&gt;
  &lt;x:CubeProvider&gt;msolap.2&lt;/x:CubeProvider&gt;
  &lt;x:CacheDetails/&gt;
  &lt;x:ConnectionString&gt;Provider=MSOLAP.5;Cache Authentication=False;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=sits;Data Source=uiodb008\sqlprd02;Impersonation Level=Impersonate;Mode=ReadWrite;Protection Level=Pkt Privacy;Auto Synch Period=20000;Default Isolation Mode=0;Default MDX Visual Mode=0;MDX Compatibility=0;MDX Unique Name Style=0;Non Empty Threshold=0;SQLQueryMode=Calculated;Safety Options=1;Secured Cell Value=0;SQL Compatibility=0;Compression Level=0;Real Time Olap=False;Packet Size=4096&lt;/x:ConnectionString&gt;
  &lt;x:DataMember&gt;sits&lt;/x:DataMember&gt;
  &lt;x:Name&gt;OLAP SITS&lt;/x:Name&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Anexos&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Anexos].[Anexos]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Autor&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Autor].[Autor]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Ciudad&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Ciudad].[Ciudad]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Row&lt;/x:Orientation&gt;
   &lt;x:Position&gt;2&lt;/x:Position&gt;
   &lt;x:GroupedFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:SubtotalFormat Style='font-family:Arial;font-size:8pt'/&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Clasificacion&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Clasificacion].[Clasificacion]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Clasificacionfuente&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Clasificacionfuente].[Clasificacionfuente]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Codigo&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Codigo].[Codigo]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Costo&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Costo].[Costo]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Descripcion&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Descripcion].[Descripcion]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Year&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Fecha].[Year]&lt;/x:SourceName&gt;
   &lt;x:FilterCaption&gt;Fecha&lt;/x:FilterCaption&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
   &lt;x:Position&gt;2&lt;/x:Position&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Month&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Fecha].[Month]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
   &lt;x:Position&gt;3&lt;/x:Position&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Day&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Fecha].[Day]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
   &lt;x:Position&gt;4&lt;/x:Position&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Fuente&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Fuente].[Fuente]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Hora&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Hora].[Hora]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Lideres&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Lideres].[Lideres]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Lugar&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Lugar].[Lugar]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Nivel&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Nivel].[Nivel]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
   &lt;x:Position&gt;5&lt;/x:Position&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Organizacion&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Organizacion].[Organizacion]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Personasaec&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Personasaec].[Personasaec]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Personascontract&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Personascontract].[Personascontract]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Personasthird&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Personasthird].[Personasthird]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Placa&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Placa].[Placa]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Provincia&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Provincia].[Provincia]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Row&lt;/x:Orientation&gt;
   &lt;x:GroupedFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:SubtotalFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:Expanded/&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;TipoOrg&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Radio].[Radio]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
   &lt;x:Position&gt;6&lt;/x:Position&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Referencia&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Referencia].[Referencia]&lt;/x:SourceName&gt;
   &lt;x:GroupedFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:SubtotalFormat Style='font-family:Arial;font-size:8pt'/&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Status&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Status].[Status]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Column&lt;/x:Orientation&gt;
   &lt;x:GroupedFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:SubtotalFormat Style='font-family:Arial;font-size:8pt'/&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;ClasificacionIncidente&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Tipo].[Tipo]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
   &lt;x:Position&gt;7&lt;/x:Position&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;CalificaciónInformación&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Tipo1].[Tipo1]&lt;/x:SourceName&gt;
   &lt;x:Orientation&gt;Page&lt;/x:Orientation&gt;
   &lt;x:Position&gt;8&lt;/x:Position&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Vehiculos&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Vehiculos].[Vehiculos]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Incident&lt;/x:Name&gt;
   &lt;x:SourceName&gt;[Measures].[Incident Numbers]&lt;/x:SourceName&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Data&lt;/x:Name&gt;
   &lt;x:Orientation&gt;Column&lt;/x:Orientation&gt;
   &lt;x:Position&gt;-1&lt;/x:Position&gt;
   &lt;x:DataField/&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotField&gt;
   &lt;x:Name&gt;Incident&lt;/x:Name&gt;
   &lt;x:TotalNumber&gt;0&lt;/x:TotalNumber&gt;
   &lt;x:PLName&gt;Incident Numbers&lt;/x:PLName&gt;
   &lt;x:Orientation&gt;Data&lt;/x:Orientation&gt;
   &lt;x:ParentField&gt;[Measures].[Incident Numbers]&lt;/x:ParentField&gt;
  &lt;/x:PivotField&gt;
  &lt;x:PivotData&gt;
   &lt;x:Top&gt;0.0.0&lt;/x:Top&gt;
   &lt;x:TopOffset&gt;0&lt;/x:TopOffset&gt;
   &lt;x:Left&gt;0.0&lt;/x:Left&gt;
   &lt;x:LeftOffset&gt;0&lt;/x:LeftOffset&gt;
   &lt;x:InvertedRowMember&gt;!.Cotopaxi.Latacunga&lt;/x:InvertedRowMember&gt;
   &lt;x:SeqNum&gt;0&lt;/x:SeqNum&gt;
  &lt;/x:PivotData&gt;
  &lt;x:PivotView&gt;
   &lt;x:IsNotFiltered/&gt;
   &lt;x:TotalFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:HeaderFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:FieldLabelFormat Style='font-family:Arial;font-size:8pt'/&gt;
   &lt;x:Label Style='font-family:Arial;font-size:8pt;background:sandybrown'&gt;
    &lt;x:Caption&gt;OLAP SITS&lt;/x:Caption&gt;
   &lt;/x:Label&gt;
  &lt;/x:PivotView&gt;
 &lt;/x:PivotTable&gt;
&lt;/xml&gt;">
      <table width='100%' cellpadding='0' cellspacing='0' border='0' height='8'><tr><td bgColor='#336699' height='25' width='10%'>&nbsp;</td><td bgColor='#666666'width='85%'><font face='Tahoma' color='white' size='4'><b>&nbsp; Missing: Microsoft Office Web Components</b></font></td></tr><tr><td bgColor='#cccccc' width='15'>&nbsp;</td><td bgColor='#cccccc' width='500px'><br> <font face='Tahoma' size='2'>This page requires the Microsoft Office Web Components.<p align='center'> <a href='file:/files/owc/setup.exe'>Click here to install Microsoft Office Web Components.</a>.</p></font><p><font face='Tahoma' size='2'>This page also requires Microsoft Internet Explorer 4.01 (SP-1) or higher.</p><p align='center'><a href='http://www.microsoft.com/windows/ie/default.htm'> Click here to install the latest Internet Explorer</a>.</font><br>&nbsp;</td></tr></table></object>
    </p>

    <p class="sub_header">
    &nbsp;</p>

      </table>
	  
	  <hr>
		  <p class="footer" align="center">
		  <!-- ########### INCLUDE FOOTER ########### Contact is specific to site-->
		  <!-- #include file="includes/contact.htm" --><br>
		  <!-- #include virtual = "/includes/htm/footer.htm" -->
		  </p>

</body>

</html>