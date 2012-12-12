<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<%
Response.Buffer = true
Response.ContentType = "text/xml;charset=windows-1256"
%>
<!--#include virtual="/beta/config.asp" -->
<!--#include virtual="/beta/include_farsiDateHandling.asp"-->
<%
select case request("act")
	case "all":
		set rs = conn.Execute("select id,name,status,isnull(exten,'') as exten, isnull(camp.member,0) as member  from CampaignName left outer join (select campaign_id, count(account_id) as member from Campaigns group by campaign_id) as camp on campaignName.id=camp.campaign_id")
		set rows=server.createobject("MSXML2.DomDocument")
		rows.loadXML("<camp/>")
		while not rs.eof
			set camp = rows.createElement("row")
			'------------
			set tmp = rows.createElement("id")
			tmp.text = rs("id")
			camp.AppendChild tmp
			'------------
			set tmp = rows.createElement("name")
			tmp.text = rs("name")
			camp.AppendChild tmp
			'------------
			set tmp = rows.createElement("status")
			if rs("status")="0" then
				tmp.text = "ÈÓÊå ÔÏå"
			elseif rs("status")="1" then 
				tmp.text = "ÝÚÇá"
			end if
			camp.AppendChild tmp
			'------------
			set tmp = rows.createElement("member")
			tmp.text = rs("member")
			camp.AppendChild tmp
			'------------
			set tmp = rows.createElement("exten")
			tmp.text = rs("exten")
			camp.AppendChild tmp
			
			rows.documentElement.AppendChild camp
			rs.MoveNext
		wend
		response.write(rows.xml)
		rs.close
		set rs = nothing
end select
%>