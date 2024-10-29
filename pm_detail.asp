<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")%> </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link rel="stylesheet" type="text/css" href=style1.css></HEAD>
<body leftmargin="0" topmargin="0">
<!--#include file="header.asp"-->
<%pid=Request("id")
Set ws=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM u2u WHERE id=" & pid & ""
ws.Open SQL,conn,1,3
'ws.Open "Select * From u2u Where id="& pid,conn,1,3
ws("readed")=0
ws.Update%>
<p align="right" style="margin-top: 0; margin-bottom: 0"><a href="u2u.asp"><img border="0" src="images/newpm.gif"></a><a href="pm.asp"><img border="0" src="images/untitled.bmp"></a><a href="friendpm.asp"><img border="0" src="images/friendpm.gif"></a>&nbsp;&nbsp;&nbsp;</p>
<center>
<TABLE border=1 width="75%" height="100" STYLE="WORD-BREAK: break-all; border-collapse:collapse" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0">
<TR>
	<TD width="40" height="23" background="images/Button.gif">
    <p align="center"><font size="2"><b>主題</b></font></TD>
	<TD width="600" height="16" background="images/Button.gif"><font size=3 color=#FF0099><b><%=ws("totitle")%></b></font></TD>
</TR>
<TR>
	<TD width="40" height="100" bgcolor="#F0EEDF" background="images/postbit_lightbg.gif">
    <p align="center"><font size="2"><b>內容</b></font></TD>
	<TD width="600" height="100" bgcolor="#F0EEDF" background="images/postbit_lightbg.gif"><font size=2 color=#666600><%=ws("totext")%></font></TD>
</TR>
</TABLE>
<%if ws("totitle") <> "[系統]文章轉移通知！請勿回應" then%>
<FORM METHOD=POST ACTION="u2uaction.asp">
<center><font color=red><%=request("msg")%></font>
<%IF UserID=ws("fromuser") then%>
收件者：<INPUT TYPE="text" NAME="touser" size="25" value=<%=ws("touser")%>>
&nbsp;&nbsp;&nbsp;&nbsp;
標題：<INPUT TYPE="text" NAME="totitle" size="40" value="<%=ws("totitle")%>~"><br>
<%else%>
收件者：<INPUT TYPE="text" NAME="touser" size="25" value=<%=ws("fromuser")%>>
&nbsp;&nbsp;&nbsp;&nbsp;
標題：<INPUT TYPE="text" NAME="totitle" size="40" value=" Re：<%=ws("totitle")%>"><br>
<%end if%>
<TEXTAREA NAME="totext" ROWS="9" COLS="82"><%=ws("totext")%><hr></TEXTAREA><br>
<INPUT TYPE="submit" value="寄出"><INPUT TYPE="reset" value="重寫">
</font>
</FORM>
<%end if%>
<%ws.close
SQL="SELECT * FROM id WHERE user='" & request.cookies("UserID") & "'"
ws.Open SQL,conn,1,3
ws("havepm")=0 '0就是簡訊已經讀了,沒有
ws.Update
ws.close%><BR>
<!--#include file="foot.asp"-->
</BODY></HTML>