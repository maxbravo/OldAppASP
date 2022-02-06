<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>LL ASP LOCAL</title>
</head>
<body>	  
<% 
' Creado por Max Bravo 25-Apr-2012
' String de Conexion con usuario de lectura a LL
'ConString = "Provider=MSDAORA.1;Password=l3cturall;User ID=llinkread;Data Source=ORAPRD03;Persist Security Info=True"
ConString = "DSN=CS10;uID=llinkread;Pwd=l3cturall;"
' Capturo el parámetro que viene

Documento = Request.QueryString("docid")

If Documento=null or Documento="" Then 
Documento = Request.QueryString("docref")

End If

Documento = Replace(Documento,"-LLECU","") 
Documento = Replace(Documento,"-ATHENA","") 
Documento = Replace(Documento,"-LLCGY","")
'Se consulta la base de datos
Set Conex = Server.CreateObject ("ADODB.Connection")
Conex.Open ConString
Set RS = Server.CreateObject("ADODB.RecordSet")
sql = "select 'http://livelinkecuador/livelinkecuador/livelink.exe?func=ll&objId=' || ID || '&objAction=viewheader' from livelinkecuador.llattrdata where DEFID=865951 and attrid=2 and rtrim(ltrim(VALSTR)) like '" & Documento & "%'"
RS.Open sql, Conex, 1
link = RS.Fields(0)
Conex.Close
Set RS = Nothing
Set Conex = Nothing
'Se abre el LL con el ID del objeto
response.redirect link
%>
</body>
</html>