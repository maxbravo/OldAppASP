<%
'Server.ScriptTimeout = 900
'Response.CacheControl = "no-cache"
'Response.AddHeader "Pragma", "no-cache"

' Variables
Dim ConString , UserAuthID , UserDomain , UserID 

' Valores
'ConString = "Provider=MSDAORA.1;Password=s3gur1dad;User ID=SITS;Data Source=ORAPRD11;Persist Security Info=True"
ConString = "DSN=SITS;uID=SITS;Pwd=s3gur1dad;"

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