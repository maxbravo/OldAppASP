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
  &lt;x:ConnectionString&gt;Provider=MSOLAP.2;Password=v1entec1t0;Persist Security Info=True;User ID=sits;Data Source=uiodb002;Initial Catalog=sits;Client Cache Size=25;Auto Synch Period=10000&lt;/x:ConnectionString&gt;
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
   &lt;x:AllIncludeExclude&gt;Exclude&lt;/x:AllIncludeExclude&gt;
   &lt;x:IncludedMember&gt;
    &lt;x:Name&gt;2010&lt;/x:Name&gt;
    &lt;x:UniqueName&gt;[Fecha].[All Fecha].[2010]&lt;/x:UniqueName&gt;
   &lt;/x:IncludedMember&gt;
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
   &lt;x:Top&gt;0.0&lt;/x:Top&gt;
   &lt;x:TopOffset&gt;0&lt;/x:TopOffset&gt;
   &lt;x:Left&gt;0.0&lt;/x:Left&gt;
   &lt;x:LeftOffset&gt;0&lt;/x:LeftOffset&gt;
   &lt;x:SeqNum&gt;0&lt;/x:SeqNum&gt;
  &lt;/x:PivotData&gt;
  &lt;x:PivotView&gt;
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
    <object classid="clsid:0002E556-0000-0000-C000-000000000046" id="ChartSpaceDeprecated1">
      <param name="XMLData" value="&lt;xml xmlns:x=&quot;urn:schemas-microsoft-com:office:excel&quot;&gt;
 &lt;x:ChartSpace&gt;
  &lt;x:OWCVersion&gt;10.0.0.2621         &lt;/x:OWCVersion&gt;
  &lt;x:Width&gt;15240&lt;/x:Width&gt;
  &lt;x:Height&gt;10160&lt;/x:Height&gt;
  &lt;x:DataSource&gt;
   &lt;x:Type&gt;PivotList&lt;/x:Type&gt;
   &lt;x:Name&gt;PivotTableDeprecated1&lt;/x:Name&gt;
  &lt;/x:DataSource&gt;
  &lt;x:BoundSeries&gt;
   &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
  &lt;/x:BoundSeries&gt;
  &lt;x:Category&gt;
   &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
  &lt;/x:Category&gt;
  &lt;x:Value&gt;
   &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
   &lt;x:Data&gt;Incident Numbers&lt;/x:Data&gt;
  &lt;/x:Value&gt;
  &lt;x:BoundCharts&gt;
   &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
  &lt;/x:BoundCharts&gt;
  &lt;x:FormatValue&gt;
   &lt;x:DataSourceIndex&gt;-3&lt;/x:DataSourceIndex&gt;
   &lt;x:Data&gt;2&lt;/x:Data&gt;
  &lt;/x:FormatValue&gt;
  &lt;x:PivotAggOrientation&gt;None&lt;/x:PivotAggOrientation&gt;
  &lt;x:NoGrouping/&gt;
  &lt;x:NoFiltering/&gt;
  &lt;x:Palette&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000000&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#8080FF&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#802060&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#FFFFA0&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#A0E0E0&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#600080&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#FF8080&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#008080&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#C0C0FF&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#000080&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#FF00FF&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#80FFFF&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#0080FF&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#FF8080&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#C0FF80&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#FFC0FF&lt;/x:Entry&gt;
   &lt;x:Entry&gt;#FF80FF&lt;/x:Entry&gt;
  &lt;/x:Palette&gt;
  &lt;x:DefaultFont&gt;Arial&lt;/x:DefaultFont&gt;
  &lt;x:Chart&gt;
   &lt;x:PlotArea&gt;
    &lt;x:Graph&gt;
     &lt;x:SubType&gt;Clustered&lt;/x:SubType&gt;
     &lt;x:Type&gt;Column&lt;/x:Type&gt;
     &lt;x:Series&gt;
      &lt;x:FormatMap&gt;
      &lt;/x:FormatMap&gt;
      &lt;x:Name&gt;Abierto&lt;/x:Name&gt;
      &lt;x:Caption&gt;
       &lt;x:DataSourceIndex&gt;-1&lt;/x:DataSourceIndex&gt;
       &lt;x:Data&gt;&amp;quot;Abierto&amp;quot;&lt;/x:Data&gt;
      &lt;/x:Caption&gt;
      &lt;x:Index&gt;0&lt;/x:Index&gt;
      &lt;x:Category&gt;
       &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
      &lt;/x:Category&gt;
      &lt;x:Value&gt;
       &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
       &lt;x:Data&gt;Incident Numbers&lt;/x:Data&gt;
      &lt;/x:Value&gt;
      &lt;x:FormatValue&gt;
       &lt;x:DataSourceIndex&gt;-3&lt;/x:DataSourceIndex&gt;
       &lt;x:Data&gt;2&lt;/x:Data&gt;
      &lt;/x:FormatValue&gt;
      &lt;x:Marker&gt;
      &lt;/x:Marker&gt;
      &lt;x:Explode&gt;0&lt;/x:Explode&gt;
      &lt;x:Thickness&gt;10&lt;/x:Thickness&gt;
      &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
      &lt;x:Identifier&gt;!.Abierto&lt;/x:Identifier&gt;
     &lt;/x:Series&gt;
     &lt;x:Series&gt;
      &lt;x:FormatMap&gt;
      &lt;/x:FormatMap&gt;
      &lt;x:Name&gt;Cerrado&lt;/x:Name&gt;
      &lt;x:Caption&gt;
       &lt;x:DataSourceIndex&gt;-1&lt;/x:DataSourceIndex&gt;
       &lt;x:Data&gt;&amp;quot;Cerrado&amp;quot;&lt;/x:Data&gt;
      &lt;/x:Caption&gt;
      &lt;x:Index&gt;1&lt;/x:Index&gt;
      &lt;x:Category&gt;
       &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
      &lt;/x:Category&gt;
      &lt;x:Value&gt;
       &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
       &lt;x:Data&gt;Incident Numbers&lt;/x:Data&gt;
      &lt;/x:Value&gt;
      &lt;x:FormatValue&gt;
       &lt;x:DataSourceIndex&gt;-3&lt;/x:DataSourceIndex&gt;
       &lt;x:Data&gt;2&lt;/x:Data&gt;
      &lt;/x:FormatValue&gt;
      &lt;x:Marker&gt;
      &lt;/x:Marker&gt;
      &lt;x:Explode&gt;0&lt;/x:Explode&gt;
      &lt;x:Thickness&gt;10&lt;/x:Thickness&gt;
      &lt;x:DataSourceIndex&gt;0&lt;/x:DataSourceIndex&gt;
      &lt;x:Identifier&gt;!.Cerrado&lt;/x:Identifier&gt;
     &lt;/x:Series&gt;
     &lt;x:Dimension&gt;
      &lt;x:ScaleID&gt;47062044&lt;/x:ScaleID&gt;
      &lt;x:Index&gt;Categories&lt;/x:Index&gt;
     &lt;/x:Dimension&gt;
     &lt;x:Dimension&gt;
      &lt;x:ScaleID&gt;47119924&lt;/x:ScaleID&gt;
      &lt;x:Index&gt;Value&lt;/x:Index&gt;
     &lt;/x:Dimension&gt;
     &lt;x:Dimension&gt;
      &lt;x:ScaleID&gt;47120128&lt;/x:ScaleID&gt;
      &lt;x:Index&gt;FormatValue&lt;/x:Index&gt;
     &lt;/x:Dimension&gt;
     &lt;x:Overlap&gt;0&lt;/x:Overlap&gt;
     &lt;x:GapWidth&gt;150&lt;/x:GapWidth&gt;
     &lt;x:FirstSliceAngle&gt;0&lt;/x:FirstSliceAngle&gt;
    &lt;/x:Graph&gt;
    &lt;x:Axis&gt;
     &lt;x:AxisID&gt;68538476&lt;/x:AxisID&gt;
     &lt;x:ScaleID&gt;47062044&lt;/x:ScaleID&gt;
     &lt;x:Type&gt;Category&lt;/x:Type&gt;
     &lt;x:MajorTick&gt;Outside&lt;/x:MajorTick&gt;
     &lt;x:MinorTick&gt;None&lt;/x:MinorTick&gt;
     &lt;x:Placement&gt;Bottom&lt;/x:Placement&gt;
     &lt;x:GroupingEnum&gt;Auto&lt;/x:GroupingEnum&gt;
    &lt;/x:Axis&gt;
    &lt;x:Axis&gt;
     &lt;x:AxisID&gt;68539084&lt;/x:AxisID&gt;
     &lt;x:ScaleID&gt;47119924&lt;/x:ScaleID&gt;
     &lt;x:Type&gt;Value&lt;/x:Type&gt;
     &lt;x:MajorGridlines&gt;
     &lt;/x:MajorGridlines&gt;
     &lt;x:MajorTick&gt;Outside&lt;/x:MajorTick&gt;
     &lt;x:MinorTick&gt;None&lt;/x:MinorTick&gt;
     &lt;x:Placement&gt;Left&lt;/x:Placement&gt;
    &lt;/x:Axis&gt;
   &lt;/x:PlotArea&gt;
   &lt;x:Identifier&gt;&lt;/x:Identifier&gt;
  &lt;/x:Chart&gt;
  &lt;x:Legend&gt;
   &lt;x:Placement&gt;Right&lt;/x:Placement&gt;
  &lt;/x:Legend&gt;
  &lt;x:Scaling&gt;
   &lt;x:ScaleID&gt;47062044&lt;/x:ScaleID&gt;
  &lt;/x:Scaling&gt;
  &lt;x:Scaling&gt;
   &lt;x:ScaleID&gt;47119924&lt;/x:ScaleID&gt;
  &lt;/x:Scaling&gt;
  &lt;x:Scaling&gt;
   &lt;x:ScaleID&gt;47120128&lt;/x:ScaleID&gt;
  &lt;/x:Scaling&gt;
  &lt;x:DisplayToolbar/&gt;
 &lt;/x:ChartSpace&gt;
&lt;/xml&gt;">
      <param name="ScreenUpdating" value="-1">
      <param name="EnableEvents" value="-1">
      <table width='100%' cellpadding='0' cellspacing='0' border='0' height='8'><tr><td bgColor='#336699' height='25' width='10%'>&nbsp;</td><td bgColor='#666666'width='85%'><font face='Tahoma' color='white' size='4'><b>&nbsp; Missing: Microsoft Office Web Components</b></font></td></tr><tr><td bgColor='#cccccc' width='15'>&nbsp;</td><td bgColor='#cccccc' width='500px'><br> <font face='Tahoma' size='2'>This page requires the Microsoft Office Web Components.<p align='center'> <a href='files/owc/setup.exe'>Click here to install Microsoft Office Web Components.</a>.</p></font><p><font face='Tahoma' size='2'>This page also requires Microsoft Internet Explorer 4.01 (SP-1) or higher.</p><p align='center'><a href='http://www.microsoft.com/windows/ie/default.htm'> Click here to install the latest Internet Explorer</a>.</font><br>&nbsp;</td></tr></table></object>
    </p>

      </table>
	  
	  <hr>
		  <p class="footer" align="center">
		  <!-- ########### INCLUDE FOOTER ########### Contact is specific to site-->
		  <!-- #include file="includes/contact.htm" --><br>
		  <!-- #include virtual = "/includes/htm/footer.htm" -->
		  </p>

</body>

</html>