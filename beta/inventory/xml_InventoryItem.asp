<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%>
<!--#include file="../config.asp" -->
<!--#include File="../JSON_2.0.4.asp"-->
<%
select case request("act")
	case "fromCode":
		Set result = jsObject()
		set store = jsArray()
		set result("store") = jsArray()
		if request("invCode")<>"" then 
			set rs = Conn.Execute("SELECT * From InventoryItems WHERE invCode="& request("invCode") & " ORDER BY id")	
			if rs.eof then 
				result("errMsg") = "äíä ßÇáÇíí æÌæÏ äÏÇÑÏ"
				result("err")=1
			else
				result("err")=0
				itemID = cint(rs("id"))
				result("itemID") = itemID
				result("itemName") = rs("name")
				set rs = Conn.Execute("select * from inventoryItemsInStore where itemID = " & itemID)
				while not rs.eof
					set result("store")(null) = jsObject()
					result("store")(null)("storeID") = rs("storeID")
					result("store")(null)("unit") = rs("unit")
					rs.moveNext
				wend
			end if
			rs.close()
			conn.close()		
		end if
	case "getStore":
		set result = jsArray()
		set rs = Conn.Execute("select * from inventoryItemsInStore where itemID = " & request("itemID"))
		while not rs.eof
			set result(null) = jsObject()
			result(null)("storeID") = rs("storeID")
			result(null)("unit") = rs("unit")
			rs.moveNext
		wend
		rs.close
		conn.close
	case "search":
		set result = jsArray()
		
		mySQL="SELECT * From InventoryItems WHERE (REPLACE([Name], ' ', '') LIKE REPLACE(N'%"& URLDecode(Server.urlencode(request("search"))) & "%', ' ', '')) ORDER BY Name"
		'mySQL="SELECT * From InventoryItems WHERE [Name] LIKE N'%ÓÈ%' ORDER BY Name"
		'response.write mySQL 'request("search")
		Set rs = conn.Execute(mySQL)
		while not rs.eof 
			'response.write rs("name")
			set result(null) = jsObject()
			result(null)("id") = rs("id")
			result(null)("name") = rs("name")
			result(null)("invCode") = rs("invCode")
			rs.moveNext
		wend
		rs.close
		conn.close
end select
Response.Write toJSON(result)

Function URLDecode(sConvert) 
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) -1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecode = sOutput 
End Function 

%>