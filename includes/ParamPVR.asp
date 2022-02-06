<%
'Server.ScriptTimeout = 900
'Response.CacheControl = "no-cache"
'Response.AddHeader "Pragma", "no-cache"

' Variables
Dim ConString , ConStringTAS ,UserAuthID , UserDomain , UserID 

' Valores
ConString = "DSN=SITS;uID=SITS;Pwd=s3gur1dad;"
ConStringTAS = "DSN=TAS;uID=TAS;Pwd=n3wt1ck3t;"
ConStringDHL = "DSN=DHL;uID=DHL;Pwd=3ntr3ga;"
ConStringPVR = "DSN=PVR_READ;uID=PVR_AECI;Pwd=summer;"
ConStringCS = "DSN=CS;uID=livelinkecuador;Pwd=budget;"
ConStringIntranet = "DSN=INTRANET;uID=INTRANET;Pwd=webuser;"


UserAuthID = request.serverVariables("LOGON_USER")
lintDelimiterPos = inStr(UserAuthID, "\")
	if lintDelimiterPos > 0 then
		UserDomain = UCase(Mid(UserAuthID, 1, lintDelimiterPos-1))
		UserID = UCase(Mid(UserAuthID, lintDelimiterPos+1, 12))
	else
		UserDomain = "ANDESPETRO"
		UserID = uCase(UserAuthID)
	end if
%>