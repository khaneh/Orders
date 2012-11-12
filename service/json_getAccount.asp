<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.Charset = "utf-8"
	Response.CodePage = 65001
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/JSON_2.0.4.asp"-->
<%
select case request("act")
	case "account":
		set j = jsArray()
		if IsNumeric(Request("search")) then
			id = CDbl(Request("search"))
			mySQL = "select id,accountTitle + '(' + cast(id as varchar(10)) + ')' as accountTitle from accounts where id= " & id
		else
			text = Replace(Request("search")," ","")
			mySQL = "select top 10 id,accountTitle from accounts where replace(accountTitle, ' ', '') like N'%" & text & "%' order by accountTitle"
		end if
		if Len(Request("search"))>3 then 
			set rs = Conn.Execute(mySQL)
			while not rs.eof
				set j(null) = jsObject()
				j(null)("id") = rs("id")
				j(null)("text") = rs("accountTitle")
				rs.MoveNext
			wend		
			rs.close
			set rs = Nothing
		end if
	case "chequeRemain":
		set j = jsObject()
		account = CDbl(request("account"))
		mySQL = "with cheque (ref1,ref2) as (select g.ref1,g.ref2 from EffectiveGLRows as g13 inner join EffectiveGLRows as g on g13.glDoc=g.glDoc and g13.sys=g.sys and g13.link=g.link where g13.GL=" & Session("OpenGL") & " and g13.sys<>'' and g13.GLAccount=13003 and g.ref1<>'' and g13.tafsil= " & account & " group by g.ref1,g.ref2,g13.tafsil) select isnull(sum(amount),0) as amount from EffectiveGLRows where id in (SELECT MAX(ID) AS MaxID FROM EffectiveGLRows inner join cheque on cheque.ref1=EffectiveGLRows.Ref1 and cheque.ref2=EffectiveGLRows.ref2 GROUP BY GLAccount, EffectiveGLRows.Tafsil, Amount, EffectiveGLRows.Ref1, EffectiveGLRows.Ref2, GL HAVING  GLAccount in (select id from GLAccounts where GLGroup<>12000 and GL=" & Session("OpenGL") & ") AND (GL = " & Session("OpenGL") & ") AND (COUNT(EffectiveGLRows.Ref1) % 2 = 1))"
		set rs = conn.Execute(mySQL)
		j("amount")=rs("amount")
end select
Response.Write toJSON(j)
%>