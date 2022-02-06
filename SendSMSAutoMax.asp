<!-- #include file="includes/ParamPVR.asp" -->
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Send SMS Oil Indicators</title>
<SCRIPT LANGUAGE="JavaScript">
		function changeImages() {
			if (document.images ) {
				for (var i=0; i<changeImages.arguments.length; i+=2) {
					document[changeImages.arguments[i]].src = changeImages.arguments[i+1];
					}
				}
		}
		function placeFocus() {
		if (document.forms.length > 0) {
		var field = document.forms[0];
		for (i = 0; i < field.length; i++) {
		if ((field.elements[i].type == "text") || (field.elements[i].type == "textarea") || (field.elements[i].type.toString().charAt(0) == "s")) {
		document.forms[0].elements[0].focus();
		break;
       }
	   }
   	   }
	   }		
</SCRIPT>
<link rel="stylesheet" type="text/css">
</head>

<body onload="placeFocus();">	

<table width="100%" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">
<tr>
<td align="left">
&nbsp;</td>
</tr>
</table>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="860">
    <tr>
    <td valign="top"></td>
	
	
    <td bgcolor="#FFFFFF" valign="top" align="left"><font size="4" color="#000080">&nbsp;&nbsp;</font></td>
    <td valign="top" width="860">


</SCRIPT>
<title>Send SMS Indicators</title>
<meta http-equiv="Content-Language" content="es-ec">



<%
Dim Conex, RS, production, wti, brent, action, tar, b14, b17, url, xxx, day

'xxx es el tamaño de los caracteres que tomamos. En este caso 5 que es ##.## Esta referencia es para el wti y brent
xxx = 5

action = request("action")

if action = "" then
	
		'Busco el indicador de la página web del WTI & Brent
		Set Conex = Server.CreateObject ("ADODB.Connection")
		Conex.Open ConStringIntranet
		Set RS = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT WTI FROM MARQUEE_DB " 
		RS.Open sql, Conex, 1
			wti = RS.Fields(0)
		Conex.Close
		Set RS = Nothing
		Set Conex = Nothing

		Set Conex = Server.CreateObject ("ADODB.Connection")
		Conex.Open ConStringIntranet
		Set RS = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT BRENT FROM MARQUEE_DB " 
		RS.Open sql, Conex, 1
			brent = RS.Fields(0)
		Conex.Close
		Set RS = Nothing
		Set Conex = Nothing	
	     
		'Busco la producción del Día
		Set Conex = Server.CreateObject ("ADODB.Connection")
		Conex.Open ConStringPVR
		Set RS = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT (TO_CHAR((PRORATED_OIL), '99,999')) FROM PVR_AECI.VW_PRODBYBLK WHERE FIELD_ID = '0'" 
		RS.Open sql, Conex, 1
			production = RS.Fields(0)
		Conex.Close
		Set RS = Nothing
		Set Conex = Nothing

		Set Conex = Server.CreateObject ("ADODB.Connection")
		Conex.Open ConStringPVR
		Set RS = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT (TO_CHAR((PRORATED_OIL), '99,999')) FROM PVR_AECI.VW_PRODBYBLK WHERE FIELD_ID = '12'" 
		RS.Open sql, Conex, 1
			tar = RS.Fields(0)
		Conex.Close
		Set RS = Nothing
		Set Conex = Nothing

		Set Conex = Server.CreateObject ("ADODB.Connection")
		Conex.Open ConStringPVR
		Set RS = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT (TO_CHAR((PRORATED_OIL), '99,999')) FROM PVR_AECI.VW_PRODBYBLK WHERE FIELD_ID = '374'" 
		RS.Open sql, Conex, 1
			b14 = RS.Fields(0)
		Conex.Close
		Set RS = Nothing
		Set Conex = Nothing

		Set Conex = Server.CreateObject ("ADODB.Connection")
		Conex.Open ConStringPVR
		Set RS = Server.CreateObject("ADODB.RecordSet")
		sql = "SELECT (TO_CHAR((PRORATED_OIL), '99,999')) FROM PVR_AECI.VW_PRODBYBLK WHERE FIELD_ID = '375'" 
		RS.Open sql, Conex, 1
			b17 = RS.Fields(0)
		Conex.Close
		Set RS = Nothing
		Set Conex = Nothing

%>
<br>
<p class="banner_blue">Datos enviados automáticamente:</p>


<form method="POST" action="SendSMSAutoMax.asp?action=enviar" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
  <table style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="20%" height="32%"><span lang="es-ec">WTI :</span></td>
    <td width="150" height="32">
    <!--webbot bot="Validation" s-display-name="WTI" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars="." b-value-required="TRUE" i-minimum-length="1" i-maximum-length="6" --><input name="wti" type="text"  size="20" maxlength="6" tabindex="1" value="<%=wti%>" readonly="true"></td>
  </tr>

  <tr>
    <td width="20%" height="32%"></span><span lang="es-ec">Brent :</span></td>
    <td width="150" height="32">
    <!--webbot bot="Validation" s-display-name="Brent" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars="." b-value-required="TRUE" i-minimum-length="1" i-maximum-length="6" --><input name="brent" type="text"  size="20" maxlength="6" tabindex="1" value="<%=brent%>" readonly="true"> </td>
  </tr>

  <tr>
    <td width="20%" height="32%"><span lang="es-ec">Total Production :</span></td>
    <td width="150" height="32">
    <!--webbot bot="Validation" s-display-name="Production" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars=".," b-value-required="TRUE" i-minimum-length="4" i-maximum-length="10" --><input name="production" type="text"  size="20" maxlength="10" tabindex="1" value="<%=production%>" readonly="true"></td>
  </tr>

  <tr>
    <td width="20%" height="32%"><span lang="es-ec">Production Tarapoa:</span></td>
    <td width="150" height="32">
    <!--webbot bot="Validation" s-display-name="Production" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars=".," b-value-required="TRUE" i-minimum-length="4" i-maximum-length="10" --><input name="tar" type="text"  size="20" maxlength="10" tabindex="1" value="<%=tar%>" readonly="true"></td>
  </tr>

  <tr>
    <td width="20%" height="32%"><span lang="es-ec">Production B14:</span></td>
    <td width="150" height="32">
    <!--webbot bot="Validation" s-display-name="Production" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars=".," b-value-required="TRUE" i-minimum-length="4" i-maximum-length="10" --><input name="b14" type="text"  size="20" maxlength="10" tabindex="1" value="<%=b14%>" readonly="true"></td>
  </tr>

  <tr>
    <td width="20%" height="32%"><span lang="es-ec">Production B17:</span></td>
    <td width="150" height="32">
    <!--webbot bot="Validation" s-display-name="Production" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars=".," b-value-required="TRUE" i-minimum-length="4" i-maximum-length="10" --><input name="b17" type="text"  size="20" maxlength="10" tabindex="1" value="<%=b17%>" readonly="true"></td>
  </tr>  

  </table>
  
  <p>&nbsp;</p>
  
  <p><i><span lang="es-ec"><font size="1">* Por favor valide previamente que la 
  información ha sido actualizada en la Intranet Corporativa.</font></span></i></p>
</form>

<%
end if
%>

<%
if action = "" then

		mensaje = "" & chr(13)
		mensaje = mensaje & "WTI:   " & wti & chr(13)
		mensaje = mensaje & "Brent: " & brent & chr(13) & chr(13) 
		mensaje = mensaje & "B14:  " & b14 & chr(13)
		mensaje = mensaje & "B17:  " & b17 & chr(13)
		mensaje = mensaje & "TAR:  " & tar & chr(13)
		mensaje = mensaje & "Total: " & production & chr(13)

		response.write "<br>" & mensaje

		'ENVIO DE EMAIL
	   strHost = "mailhost.andespetro.com"
	   Set Mail = Server.CreateObject("Persits.MailSender")
	   Mail.Host = strHost
	   Mail.From = "DailyIndicators@andespetro.com" 
	'Max
	   Mail.AddAddress "max.bravo@andespetro.com"
	   Mail.Subject = " .Oil Indicators as of today"
	   Mail.IsHTML = False
	   body = mensaje
	   Mail.Body = body
	   'tomo el dia de la semana para no enviar sabado y domingo
	   day = Weekday(Now())
	   If (day <> 1 AND day <> 7) Then
		   Mail.Send 
	   End If	
	   Set Mail = Nothing

%><br><p class="banner_blue">Envío Exitoso</p>
<%
End If
%>
	  </table>
	  
	  <hr>

</body>

</html>