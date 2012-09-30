<%
Function persian_jdn(iYear, _
                     iMonth, _
                     iDay) 'As Long
    Const PERSIAN_EPOCH = 1948321 ' The JDN of 1 Farvardin 1
    Dim epbase 
    Dim epyear 
    Dim mdays 
    If iYear >= 0 Then
        epbase = iYear - 474
    Else
        epbase = iYear - 473
    End If
    epyear = 474 + (epbase Mod 2820)
    If iMonth <= 7 Then
        mdays = (CLng(iMonth) - 1) * 31
    Else
        mdays = (CLng(iMonth) - 1) * 30 + 6
    End If
    persian_jdn = CLng(iDay) _
            + mdays _
            + Fix(((epyear * 682) - 110) / 2816) _
            + (epyear - 1) * 365 _
            + Fix(epbase / 2820) * 1029983 _
            + (PERSIAN_EPOCH - 1)
End Function

Sub jdn_persian(jdn, _
                ByRef iYear, _
                ByRef iMonth, _
                ByRef iDay)
    Dim depoch
    Dim cycle
    Dim cyear
    Dim ycycle
    Dim aux1, aux2
    Dim yday
    depoch = jdn - persian_jdn(475, 1, 1)
    cycle = Fix(depoch / 1029983)
    cyear = depoch Mod 1029983
    If cyear = 1029982 Then
        ycycle = 2820
    Else
        aux1 = Fix(cyear / 366)
        aux2 = cyear Mod 366
        ycycle = Int(((2134 * aux1) + (2816 * aux2) + 2815) / 1028522) + aux1 + 1
    End If
    iYear = ycycle + (2820 * cycle) + 474
    If iYear <= 0 Then
        iYear = iYear - 1
    End If
    yday = (jdn - persian_jdn(iYear, 1, 1)) + 1
    If yday <= 186 Then
        iMonth = Ceil(yday / 31)
    Else
        iMonth = Ceil((yday - 6) / 30)
    End If
    iDay = (jdn - persian_jdn(iYear, iMonth, 1)) + 1
End Sub

' We needed an alternative to Int and Fix.
' Int(8.4) = 8, Int(-8.4) = -9
' Fix(8.4) = 8, Fix(-8.4) = -8
' Ceil(8.4) = 9, Ceil(-8.4) = -9
Function Ceil(number)
    Ceil = -Sgn(number) * Int(-Abs(number))
' or
    'Ceil = CInt(number + (Sgn(number) * 0.5))
End Function

Const Gregorian = 1
Const Julian = 2
Function civil_jdn(ByVal iYear, _
                   ByVal iMonth, _
                   ByVal iDay ) 'As Long
	calendarType = Gregorian
    Dim lYear
    Dim lMonth
    Dim lDay

    If calendarType = Gregorian And ((iYear > 1582) Or _
        ((iYear = 1582) And (iMonth > 10)) Or _
        ((iYear = 1582) And (iMonth = 10) And (iDay > 14))) _
    Then
        lYear = CLng(iYear)
        lMonth = CLng(iMonth)
        lDay = CLng(iDay)
        civil_jdn = ((1461 * (lYear + 4800 + ((lMonth - 14) \ 12))) \ 4) _
            + ((367 * (lMonth - 2 - 12 * (((lMonth - 14) \ 12)))) \ 12) _
            - ((3 * (((lYear + 4900 + ((lMonth - 14) \ 12)) \ 100))) \ 4) _
            + lDay - 32075
    Else
        civil_jdn = 0 ' julian_jdn(iYear, iMonth, iDay)
    End If

End Function

Sub jdn_civil(ByVal jdn, _
              ByRef iyear, _
              ByRef imonth, _
              ByRef iday)

    Dim l
    Dim k
    Dim n
    Dim i
    Dim j

    If (jdn > 2299160) Then
        l = jdn + 68569
        n = ((4 * l) \ 146097)
        l = l - ((146097 * n + 3) \ 4)
        i = ((4000 * (l + 1)) \ 1461001)
        l = l - ((1461 * i) \ 4) + 31
        j = ((80 * l) \ 2447)
        iday = l - ((2447 * j) \ 80)
        l = (j \ 11)
        imonth = j + 2 - 12 * l
        iyear = 100 * (n - 49) + i + l
    Else
        Call jdn_julian(jdn, iyear, imonth, iday)
    End If

End Sub





Function shamsiToday()
    Dim depoch
    Dim cycle
    Dim cyear
    Dim ycycle
    Dim aux1, aux2
    Dim yday
	Dim jdn
	tmpDate=date()

	jdn=civil_jdn(clng(year(tmpDate)),clng(Month(tmpDate)),clng(Day(tmpDate)))

    depoch = jdn - 2121446  ' 2121446 = persian_jdn(475, 1, 1)
    cycle = Fix(depoch / 1029983)
    cyear = depoch Mod 1029983
    If cyear = 1029982 Then
        ycycle = 2820
    Else
        aux1 = Fix(cyear / 366)
        aux2 = cyear Mod 366
        ycycle = Int(((2134 * aux1) + (2816 * aux2) + 2815) / 1028522) + aux1 + 1
    End If
    iYear = ycycle + (2820 * cycle) + 474
    If iYear <= 0 Then
        iYear = iYear - 1
    End If
    yday = (jdn - persian_jdn(iYear, 1, 1)) + 1
    If yday <= 186 Then
        iMonth = Ceil(yday / 31)
    Else
        iMonth = Ceil((yday - 6) / 30)
    End If
    iDay = (jdn - persian_jdn(iYear, iMonth, 1)) + 1
	
	if iDay < 10 then iDay = "0" & iDay
	if iMonth < 10 then iMonth = "0" & iMonth

shamsiToday=iYear & "/" & iMonth & "/" & iDay
End Function

Function shamsiDate(inputDate)
    Dim depoch
    Dim cycle
    Dim cyear
    Dim ycycle
    Dim aux1, aux2
    Dim yday
	Dim jdn

	jdn=civil_jdn(clng(year(inputDate)),clng(Month(inputDate)),clng(Day(inputDate)))

    depoch = jdn - 2121446  ' 2121446 = persian_jdn(475, 1, 1)
    cycle = Fix(depoch / 1029983)
    cyear = depoch Mod 1029983
    If cyear = 1029982 Then
        ycycle = 2820
    Else
        aux1 = Fix(cyear / 366)
        aux2 = cyear Mod 366
        ycycle = Int(((2134 * aux1) + (2816 * aux2) + 2815) / 1028522) + aux1 + 1
    End If
    iYear = ycycle + (2820 * cycle) + 474
    If iYear <= 0 Then
        iYear = iYear - 1
    End If
    yday = (jdn - persian_jdn(iYear, 1, 1)) + 1
    If yday <= 186 Then
        iMonth = Ceil(yday / 31)
    Else
        iMonth = Ceil((yday - 6) / 30)
    End If
    iDay = (jdn - persian_jdn(iYear, iMonth, 1)) + 1
	
	if iDay < 10 then iDay = "0" & iDay
	if iMonth < 10 then iMonth = "0" & iMonth

shamsiDate=iYear & "/" & iMonth & "/" & iDay
End Function

Function CheckDateFormat(ByVal inputDate)

	CheckDateFormat=False

	Set RExp = New RegExp
		RExp.Pattern = "^13\d\d/(0[1-6]/(0[1-9]|[12][0-9]|3[01])|(0[789]|1[012])/(0[1-9]|[12][0-9]|30))$"
		RExp.IgnoreCase = False
		RExp.Global = False

	If RExp.Test(inputDate) Then
		CheckDateFormat=True
	end if

	Set RExp = nothing

End Function

Function currentTime10()
	curTime= now
	curHour		= Hour(curTime)
	curMinute	= Minute(curTime)
	curSecond	= Second(curTime)
	if (len(curHour) = 1) then curHour = "0" & curHour
	if (len(curMinute) = 1) then curMinute = "0" & curMinute 
	if (len(curSecond) = 1) then curSecond = "0" & curSecond 
	currentTime10= curHour & ":" & curMinute & ":" & curSecond
End Function 
function weekdaynameFA(name)
	select case name
		case "Saturday":
			weekdaynameFA="ÔäÈå"
		case "Sunday":
			weekdaynameFA="íßÔäÈå"
		case "Monday":
			weekdaynameFA="ÏæÔäÈå"
		case "Tuesday":
			weekdaynameFA="ÓåÔäÈå"
		case "Wednesday":
			weekdaynameFA="åÇÑÔäÈå"
		case "Thursday":
			weekdaynameFA="äÌÔäÈå"
		case "Friday":
			weekdaynameFA="ÌãÚå"
		case else
			 weekdaynameFA=name
	end select
end function
function splitDate(myDate)
	dim out(3)
	out(0)=mid(myDate,1,4)
	out(1)=mid(myDate,6,2)
	out(2)=mid(myDate,9,2)
	splitDate=out
end function
%>
