<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><!--#include file="conn.asp"-->
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<meta http-equiv="refresh" content="1; url=pm.asp"> 
<link rel="stylesheet" type="text/css" href=style1.css></HEAD>
<BODY leftmargin="0" topmargin="0">
<!--#include file="header.asp"-->
<%if request.cookies("UserPassed")="yes" then
touser=request("touser")
totitle=request("totitle")
totext=request("totext")
If touser ="" or totitle="" or totext="" Then
	response.redirect "u2u.asp?msg=���d��!"
End If

Set rs1=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM id WHERE user='" & touser & "'"
rs1.Open SQL, conn,1,3

Set rs2=Server.CreateObject("ADODB.Recordset")
rs2.Open "Select * From u2u  ",conn,1,3

    If rs1.Eof Then 
		response.redirect "u2u.asp?msg=�L���b��!"
	Else
	    rs1("havepm")=1
		rs1.Update
	    rs2.AddNew
		rs2("fromuser")=request.cookies("UserID")
        rs2("touser")=touser
		rs2("totitle")=totitle
		rs2("time")=now()		
		rs2("totext")=totext
		rs2.Update
		response.write "<center><br><br><br><font size=-1>���\�H�X!<font>"
		response.write "<br><center><font size=-1><a href=u2u.asp>�^�W�@��</a><br><br><br><br>"
	End If
else
 response.redirect "index.asp?msg=�Х��n��!"
end if%>
<!--#include file="foot.asp"-->
</BODY></HTML>