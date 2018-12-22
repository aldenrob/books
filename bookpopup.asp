<!-- #INCLUDE FILE="bookLib.Asp" -->
<html>
<body>
<%
Dim sSql
	sSql = "Select * from booklist where "

if Request("isbn") <> "" then
	sSql = sSql & "ISBNNumber = '" & Request("isbn") & "'"
end if

OpenDB

OpenRS rsBOOK, sSql

rsBOOK.MoveFirst

strPageTitle = "RTS OnLine"
strTitle = "Extended Text on " & rsBOOK("booktitle")
strMeta = ""
Header strPageTitle, strTitle, strMeta
%>
<table cellspacing=0>
<tr>
<td height=23 width=20>&nbsp;</td>

<td colspan=2>
<table border="0" cellspacing="1" cellpadding="2" valign="top" bgcolor="#efd694">
					<tr bgcolor="#ffffff">
						<td align="center" bgcolor="#FDF4DA">&nbsp;&nbsp;<b>
<%

if Request("description") <> "" then
	response.write "Description"
end if
if Request("samplepara") <> "" then
	response.write "Sample Paragraph"
end if
if Request("review") <> "" then
	response.write "Editorial Review"
end if
if Request("userreview") <> "" then
	response.write "User Review"
end if
%>
						</b>&nbsp;&nbsp;</td>
						<td align="center" bgcolor="white">&nbsp;&nbsp;
<a href='javascript: window.close()'>Close</a>&nbsp;&nbsp;
						</td>
				    </tr>
					</table>
</tr>
<tr><td width=20>&nbsp;</td>
<td colspan=2>
<table border="0" cellspacing="1" cellpadding="2" valign="top" bgcolor="#efd694">
					<tr bgcolor="#ffffff">
						<td>
<%
'description, samplepara, review, userreview
if Request("description") = "1" then
	response.write rsBOOK("description")
end if
if Request("samplepara") = 1 then
	response.write rsBOOK("samplepara")
end if
if Request("review") = 1 then
	response.write rsBOOK("review")
end if
if Request("userreview") = 1 then
	response.write rsBOOK("userreview")
end if
rsBOOK.close
set rsBOOK = nothing
%>
						&nbsp;</td>
				    </tr>
					</table>
<br>
<a href='javascript: window.close()'>Close</a>
</td></tr></table>
</body>
</html>