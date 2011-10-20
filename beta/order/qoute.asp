<%@LANGUAGE="VBSCRIPT" CODEPAGE="1256"%><%
'Order (2)
PageTitle=" Œ„Ì‰ ﬁÌ„ "
SubmenuItem=4
if not Auth(2 , 4) then NotAllowdToViewThisPage()

%>
<!--#include file="top.asp" -->
<!--#include File="../include_farsiDateHandling.asp"-->
<!--#include File="../include_JS_InputMasks.asp"-->
<SCRIPT LANGUAGE="JavaScript">
<!-- 
// By Alix - 1381/8/28 (ver 1.01)

//-------------------------- Variables ---------------

var A					= 90		// DuplexFee Add-in
var ProofSimplex	= 3000
var ProofDuplex		= 4500
					  Qtt	= Array (	0 ,	 				1,				50,			150,			300 )
/*SimplexFee*/  SF	= Array (	ProofSimplex,	200,			144,			112,			100 )
/*DuplexFee*/	  DF	= Array (	ProofDuplex,		SF[1]+A,		SF[2]+A,	 	SF[3]+A,		SF[4]+A)

/*OLD paperFee	  PF	= Array ( 0, 5, 12, 20, 25, 35, 50 ) 
			0:		60-UNCOAT
			5:		90-UNCOAT
			12:	100-UNCOAT
			20:	120-UNCOAT,		115-SILK
			25:	135-GLOSS,		160-UNCOAT,		150-GLOSS,		150-SILK
			35:	200-UNCOAT,		200-SILK,			200-GLOSS,		170-GLOSS
			50:	250-UNCOAT,		250-SILK			250-GLOSS
*/

/*Summer 83 paperFee*/	  PF	= Array ( 0, 25, 50 )

//-------------------------- Variables ---------------


function Calculate()
{
PT		= document.all.PaperType.value
SoD	= document.all.SoD.checked
Tirag	= Math.round(document.all.Tirag.value)
h		= document.all.height.value
Price	= 0
i		= 1
document.all.Tirag.value= Tirag	

if (h < 15) 
	{
	document.all.height.value=15
	document.all.ho.innerHTML="&nbsp;&nbsp;  Õœ«ﬁ· «‰œ«“Â: 30 * 15"
	h=15
	}
else
	document.all.ho.innerText=" "



/*
while ( Tirag >= Qtt[ i -1] ) 
	{
	a1 = Tirag - Qtt[ i - 1 ]
	a2 = Tirag - Qtt[ i ]
	if (a2>0)	
		a3 = a1 - a2 
	else 
		a3 = a1
	if ( SoD == false ) 
		Price += ( SF[ i-1 ] + PF [ PT ] ) * a3
	else
		Price += ( DF[ i-1 ] + PF [ PT ] ) * a3

	i++
	}
*/

Price = ( 200 + PF [ PT ] ) * Tirag
if ( SoD == false ) 
	Price = Price * 1
else
	Price = Price * 2

Price = Math.round(Price / 21 * h)
unitPrice = Math.round(Price / Tirag)
thenPrice = unitPrice * Tirag
document.all.Qtty.value = Tirag 
document.all.UnitPrice.value = unitPrice 
document.all.Price.value = thenPrice 
}

//-->
</SCRIPT>


<BR>

<FORM METHOD=POST ACTION="">
<br>
<TABLE Align=center height=350 width=350 style="font-size:12pt">
<TR bgcolor=eeeeee>
	<TD Align=center style="font-size:15pt"><font color=black><b> Œ„Ì‰ ﬁÌ„  ç«Å œÌÃÌ «·<br></b>
	<font size=1>Version 1.01 - 1381/8/28
	</font></TD>
</TR>
<TR bgcolor=#DDDDDD>
	<TD>
			&nbsp;&nbsp;&nbsp;ﬂ«€–:
			<SELECT NAME="PaperType"  style="font-size:12pt" onchange="Calculate()" >
			<option Value="0">60-UNCOATED</option>
			<option Value="0">90-UNCOATED</option>
			<option Value="0">100-UNCOATED</option>
			<option Value="1">120-UNCOATED</option>
			<option Value="1">115-SILK COATED</option>
			<option Value="1">135-GLOSS COATED</option>
			<option Value="2">160-UNCOATED</option>
			<option Value="2">150-GLOSS COATED</option>
			<option Value="2">150-SILK COATED</option>
			<option Value="2">200-UNCOATED</option>
			<option Value="2">200-SILK COATED</option>
			<option Value="2">200-GLOSS COATED</option>
			<option Value="2">170-GLOSS COATED</option>
			<option Value="2">250-UNCOATED</option>
			<option Value="2">250-SILK COATED</option>
			<option Value="2">250-GLOSS COATED</option>
			</SELECT>
	</TD>
</TR>
<TR bgcolor=eeeeee>
	<TD>
			&nbsp;&nbsp;&nbsp;<INPUT TYPE="checkbox" NAME="SoD" value="1" style="width:20"  onclick="Calculate()" > œÊ —Ê <!--SMALL>(Â— œÊ ÿ—› ﬂ«€– ç«Å ‘œÂ)</SMALL-->
	</TD>
</TR>
<TR bgcolor=dddddd>
	<TD>
			&nbsp;&nbsp;&nbsp;”«Ì“: <INPUT TYPE="text" NAME="width" value=30  style="font-size:12pt" disabled readonly size=2> * <INPUT TYPE="text" NAME="height" value=21  style="font-size:12pt"  size=2  onchange="Calculate()" ><span id="ho"></span>
	</TD>
</TR>
<TR bgcolor=eeeeee>
	<TD>
			&nbsp;&nbsp;&nbsp; ⁄œ«œ ﬂÅÌ: <INPUT TYPE="text" NAME="Tirag" value=1  style="font-size:12pt"  onchange="Calculate()" size=7>
	</TD>
</TR>
<TR bgcolor=dddddd>
	<TD align=center><INPUT TYPE="button" value="„Õ«”»Â" onclick="Calculate()"  style="font-size:12pt"><br><br>
		<TABLE>
		<TR>
			<TD><SMALL> ⁄œ«œ</SMALL></TD>
			<TD><SMALL>ﬁÌ„  Ê«Õœ</SMALL></TD>
			<TD><SMALL>ﬁÌ„  ﬂ·</SMALL></TD>
		</TR>
		<TR>
			<TD><INPUT TYPE="text" NAME="Qtty"  style="font-size:12pt" readonly size=3> * </TD>
			<TD><INPUT TYPE="text" NAME="UnitPrice"  style="font-size:12pt" readonly size=4> = </TD>
			<TD><INPUT TYPE="text" NAME="Price"  style="font-size:12pt" readonly size=5>  Ê„«‰</TD>
		</TR>
		</TABLE>
	</TD>
</TR>
</TABLE>


</FORM>

<!--#include file="tah.asp" -->