   <!--#include file="conn.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<meta http-equiv="refresh" content="3; url=pm.asp">  
<link rel="stylesheet" type="text/css" href=style1.css>
</HEAD>

<BODY leftmargin="0" topmargin="0">

<!--#include file="header.asp"-->
<%
Set rs=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM u2u WHERE touser='" & request.cookies("UserID") & "'"
SQL=SQL & "Order By time Desc"
rs.Open SQL,conn,1,3
If not rs.EOF then rs.movefirst
while not rs.EOF
	If request(rs("id"))="ON" then 
		rs.delete
		response.write "<center><font size=-1>已成功刪除!~<br>"
	End if
	
  rs.movenext
wend
response.write "<br><center><font size=-1><a href=pm.asp>回上一頁</a><br>"
%>
<br><br><br><br>
<!--#include file="foot.asp"-->

</BODY>
</HTML>
