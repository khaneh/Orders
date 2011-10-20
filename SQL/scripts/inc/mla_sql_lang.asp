<%
	Dim  gTArray(1000)
	Class mla_sql_lang
		Dim gCurrentLng

		Public Property Let lng (pLng)
			gCurrentLng = pLng
			InitTranslation()
		End Property

		Public Function getTerm(pId)
			getTerm = gTArray(pId)
		End Function

		Private Sub InitTranslation()
			Select Case gCurrentLng
				Case "FR" :
%>
					<!-- #INCLUDE FILE="../lng/french.asp" -->
<%
				Case Else :
%>
					<!-- #INCLUDE FILE="../lng/english.asp" -->
<%
			End Select
		End Sub

	End Class
%>