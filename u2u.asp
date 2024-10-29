<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link rel="stylesheet" type="text/css" href=style1.css></HEAD>
<BODY leftmargin="0" topmargin="0">
<!--#include file="header.asp"-->
<p align="right" style="margin-top: 0; margin-bottom: 0"><a href="u2u.asp"><img border="0" src="images/newpm.gif"></a><a href="pm.asp"><img border="0" src="images/untitled.bmp"></a><a href="friendpm.asp"><img border="0" src="images/friendpm.gif"></a>&nbsp;&nbsp;&nbsp;</p>
<FORM METHOD=POST ACTION="u2uaction.asp">
<center><font color=red><%=request("msg")%></font>
收件者：<INPUT TYPE="text" NAME="touser" size="25" value=<%=request("touser")%>><br>
&nbsp;&nbsp;&nbsp;&nbsp;標題：<INPUT TYPE="text" NAME="totitle" size="40" value=<%=request("title")%>><br>
內容：<TEXTAREA NAME="totext" ROWS="9" COLS="57"></TEXTAREA><br>
<INPUT TYPE="submit" value="寄出"><INPUT TYPE="reset" value="重寫">
</FORM></font>
<!--#include file="foot.asp"-->
</BODY></HTML>