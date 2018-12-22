<!-- #INCLUDE FILE="adovbs.inc" -->
<!-- #INCLUDE FILE="Config.Asp" -->
<%

'Server.ScriptTimeout = 9999

CONST dbFile = "books.mdb"

const xSessionTimeout=""         ' timeout period Use system default

const xUseCookies="No"
const CookieKey="Books"

Sub SetSess (field, value)
	If xUseCookies<>"Yes" then
		Session(field)=value
	else
'		debugwrite field & " value=" & value
		Response.cookies(Cookiekey) (field)=value
	end if   
end sub

Function GetSess (field)
	dim value
	if xUseCookies<>"Yes" then
	  value=Session(field)
	  Getsess=value
	else
	  value=Request.cookies(Cookiekey) (field)
	  Getsess=value
	end if
End Function

Sub SetSessionTimeout
	If XuseCookies<>" Yes" then
		If xSessionTimeout<>"" then
			Session.timeout=xsessiontimeout
		end if
		exit sub
	end if
	Response.Cookies(cookiekey).expires = date+1
end sub

Sub WriteCookie (keyname, dataarea)
	Response.cookies(Cookiekey) (keyname)=dataarea
end sub

Sub ReadCookie (keyname, dataarea)
	dataarea =  Request.cookies(Cookiekey) (keyname)
end sub


Sub OpenDB()
	
'debugwrite "dbConn = " & dbConn
'response.End
	Dim dbConn, strConn, strLogin, strPassword
	Set dbConn = Server.CreateObject("ADODB.Connection")
	strConn = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & Server.MapPath(dbFile) & ";Persist Security Info=False"
	strLogin = "Admin"
	strPassword = ""
	dbConn.open strConn, strLogin, strPassword
	SetSess "dbc", dbConn
	If dbConn.errors.count> 0 then
		SetSess "Openerror", "Open Messages<br>" & dbConn.errors(0).description & " <br>" & GetSess("dbc")
	else
		SetSess "Openerror",""
	end if

End Sub

Sub CloseDatabase (connection)
	on error resume next
	connection.close
	set connection=nothing
End sub


Sub DeBugWrite (strText)
	Response.Write strText & "<br>"

End Sub

sub OpenRS(rs, sql)
	if GetSess("dbc") = "" then 
		OpenDB
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseServer
	rs.Open sql, GetSess("dbc"), adOpenStatic, adLockReadOnly, adCmdText
	
end sub

sub OpenStaticRS(rs, sql)
	if GetSess("dbc") = "" then 
		OpenDB
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = adUseServer
	rs.Open sql, GetSess("dbc"), adOpenStatic, adLockReadOnly, adCmdText
	
end sub

sub Header(strPageTitle, strTitle, strMeta)
%>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
	<HTML>
	<HEAD>
	<TITLE> <%=CookieKey%> - <%=strPageTitle%> </TITLE>
	<META NAME="Generator" CONTENT="EditPlus">
	<META NAME="Author" CONTENT="TSgt Paul P. Beaulieu ESC/CIO 478-9265 Hanscom AFB MA">
	<META NAME="Keywords" CONTENT="">
	<META NAME="Description" CONTENT="">
	<link rel="stylesheet" type="text/css" href="ocean.css">

	<%=strMeta%>
	</HEAD>
	<BODY>
	<p align=center>
	  <table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" >

    <tr>
	<%
	if strTitle <> "" then
		response.write "<td width=100% bgcolor=#FDF4DA><b>&nbsp;&nbsp;" & strTitle & "&nbsp;&nbsp;</B></td>"
	else
		response.write "<td>&nbsp;</td>"
	end if
	%>
    </tr>
	</table>
	</p>
<%
End Sub

sub Footer
%>
	<br>
	<p align="Center">
	</p>
	</BODY>
	</HTML>
<%
End Sub

Sub SendMail(strFrom, strSubject, strBody, strTO)

  Dim objMail
  Set objMail = Server.CreateObject("CDONTS.NewMail")
  
     objMail.From = strFrom '"Paul.Beaulieu@hanscom.af.mil"
     objMail.Subject = strSubject ' "Test"
     objMail.To = strTO '"Paul.Beaulieu@hanscom.af.mil"
'	 objMail.Bcc = strBCC
     objMail.Body = strBody ' "This is a Test email message"
	objMail.Send
	set objMail = Nothing

End Sub

sub DLHeader(strPageTitle, strTitle, strMeta)
%>
	<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
	<HTML>
	<HEAD>
	<TITLE> IDE - <%=strPageTitle%> </TITLE>
	<META NAME="Generator" CONTENT="EditPlus">
	<META NAME="Author" CONTENT="TSgt Paul P. Beaulieu ESC/CIO 478-9265 Hanscom AFB MA">
	<META NAME="Keywords" CONTENT="">
	<META NAME="Description" CONTENT="">
	<%=strMeta%>

<script language="javaScript" type="text/javascript" SRC="pz_chromeless_2.1.js"></SCRIPT>

<script>
function showAlert() {
	theURL="Alert.html"
	wname ="CHROMELESSWIN"
	W=400;
	H=200;
	windowCERRARa 		= "Images/close_a.gif"
	windowCERRARd 		= "Images/close_d.gif"
	windowCERRARo 		= "Images/close_o.gif"
	windowNONEgrf 		= "Images/none.gif"
	windowCLOCK 		= "Images/clock.gif"
	windowREALtit		= " ù WARNING!!"
	windowTIT 	    	= "<font face=verdana size=1> ù IDE Disclaimer</font>"
	windowBORDERCOLOR   	= "#000000"
	windowBORDERCOLORsel	= "#999999"
	windowTITBGCOLOR    	= "#999999"
	windowTITBGCOLORsel 	= "#333333"
	openchromeless(theURL, wname, W, H, windowCERRARa, windowCERRARd, windowCERRARo, windowNONEgrf, windowCLOCK, windowTIT, windowREALtit , windowBORDERCOLOR, windowBORDERCOLORsel, windowTITBGCOLOR, windowTITBGCOLORsel)
}
</script>


	</HEAD>
	<BODY link="#0000FF" vlink="#0000FF" style="font-family: Arial; font-size: 10pt" onLoad=showAlert()>
	<p align=center>
	  <table border="0" cellpadding="0" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" >
    <tr>
      <td width="100%" height="73">
      <img border="0" src="Images/cataloghead.jpg" width="350" height="73"></td>
    </tr>
    <tr>
      <td width="100%" height="19" align=center bgcolor="#DBD733"><b><i><%=strTitle%>&nbsp;</i></b></td>
    </tr>
	</table>
	</p>
<%
End Sub


Sub EndIt
	Response.End
End Sub

sub Admin_Footer
%>
	<p align="Center">
		<a href="Admin_Main.asp">Admin Main Menu</a>
	</p>
<%
End Sub


Function ConvertTF(blnTF)

	if blnTF then
		ConvertTF = "Yes"
	else
		ConvertTF = "No"
	End if

End Function

%>