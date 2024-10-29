<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN"><!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
%> </TITLE>

<meta http-equiv="refresh" content="1200; url=index.asp">  
<META NAME="Description" CONTENT="">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href=style1.css>
</HEAD>
<BODY leftmargin="0" topmargin="0">
<!--#include file="safe.asp"-->
<!--#include file="header.asp"-->
<!--#include file="admin/broad.asp"-->
<br>
<% response.cookies("tid")="公告事項"%><!--#include file="choice.asp"--><br><br>
<% ID=Request("ID")
   Set rsbr = Server.CreateObject("ADODB.Recordset")
   SQL = "Select * From broad "
   rsbr.Open SQL, conn,1,3%><center><br><br><br>
<%

If Not rsbr.EOF Then
while not rsbr.eof
%>
<TABLE bordercolor="#808080" border=1 cellspacing="0" cellpadding="0" style="border-collapse: collapse" width="530">
<TR><TD width="528" colspan="2" height="13" bordercolor="#E1E0BB" background="images/Button.gif" ><font size="2">發佈者 : <font color=red><%=rsbr("UserID")%></font></font></TD></TR>
<TR><TD width="528" colspan="2" height="13" ><font size="2">標題 <font color=blue>:<b><%=rsbr("Subject")%></b></font>&nbsp;<font color=red>(<%=rsbr("date")%>)</font></font></TD></TR>
<TR><TD width="528" colspan="2" height="14" bordercolor="#F0EFDD" bgcolor="#F0EFDD" background="images/postbit_lightbg.gif"><font size="2">內容:</font></TD></TR>
<tr><TD width="26" height="16" bgcolor="#F0EFDD" bordercolor="#F0EFDD" background="images/postbit_lightbg.gif"></TD>
	<TD width="500" height="16" bgcolor="#F0EFDD" background="images/postbit_lightbg.gif"><font size=-1 color=#0099FF><%=rsbr("text")%></font></TD></tr>
</TABLE><br>
<%
rsbr.movenext
wend
End If

%><br><br><br><br><br>
  <!--#include file="foot.asp"-->
</BODY></HTML>