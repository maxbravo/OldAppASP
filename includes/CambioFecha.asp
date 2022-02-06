<%
Function CambioFecha(paso)
    Dim fecha
    fecha=paso
    If IsDate(fecha) = True Then
       DIM dteDay, dteMonth, dteYear
       dia = Day(fecha)
       mes = Month(fecha)
       ano = Year(fecha)
       mesn = Month(fecha)
	If dia < 10 Then
		dia = "0" & dia
	End If
	If mesn < 10 Then
		mesn = "0" & mesn
	End If
	If mes = "1" Then
		mes = "Jan"
	End If
	If mes = "2" Then
		mes = "Feb"
	End If
	If mes = "3" Then
		mes = "Mar"
	End If
	If mes = "4" Then
		mes = "Apr"
	End If
	If mes = "5" Then
		mes = "May"
	End If
	If mes = "6" Then
		mes = "Jun"
	End If
	If mes = "7" Then
		mes = "Jul"
	End If
	If mes = "8" Then
		mes = "Aug"
	End If
	If mes = "9" Then
		mes = "Sep"
	End If
	If mes = "10" Then
		mes = "Oct"
	End If
	If mes = "11" Then
		mes = "Nov"
	End If
	If mes = "12" Then
		mes = "Dec"
	End If	
       CambioFecha = dia & "-" & mes & "-" & ano
    Else
       CambioFecha = Null
    End If
End Function
%>