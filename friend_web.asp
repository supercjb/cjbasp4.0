<%sql="select * from friend_web"
set rs_web=conn.execute(sql)%><br>
<TABLE  border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse;WORD-BREAK: break-all" bordercolor="#C0C0C0" width="98%" height="22">
<TR><TD width="524" bordercolor="#FFFFFF" bgcolor="#EFF4FE" background="images/Button.gif" height="7">
	<font size=2><b>¡·&nbsp;Áp ·ù ºô ¯¸</b></font></TD></TR>
<%if not rs_web.eof then%><TR>
	<TD bgcolor="#EFF4FE" background="images/postbit_lightbg.gif"><%while not rs_web.eof%>
<%if rs_web("logo")="0" then%>
&nbsp;&nbsp;<a href="<%=rs_web("url")%>" target=_blank><font size=2><%=rs_web("name")%></font></a>
<%else%>
&nbsp;&nbsp;<a href="<%=rs_web("url")%>" target=_blank ><img src=<%=rs_web("logo")%> border=0 width="88" height="31" alt="<%=rs_web("name")%>"></a>
<%end if%>
<%rs_web.movenext
wend%></TD></TR>
<%end if%>
</TABLE>