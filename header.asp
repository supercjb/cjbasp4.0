<!--#include file="conn.asp"-->

<center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#000000" width="99%" height="3" topmargin="0" leftmargin="0" bgcolor="#E3E3E3" background="images/Button.gif">
<tr><td width="83%" height="3" bordercolor="#E3E3E3" background="images/windowbar_lightbg.gif"><font size=-1>
<% Response.Write "<font color=red SIZE=-1><b>"&UserID&"</font>"&"</b>-�w��z�^��!&nbsp;&nbsp;"%> 
<%if UserID="�X��" then %><img border="0" src="images/reg.gif"><a href="registrule.asp">�|�����U</a>
<img border="0" src="images/login.gif"><a href="login.asp">�|���n��</a>
<%end if%>
<% If request.cookies("UserPassed") = "yes" and not rs1.EOF Then %>
| &nbsp;<a href=edituser.asp?UserID=<%=request.cookies("UserID")%>><font color="#000000">�ӤH���</font></a> |
&nbsp;<a href="user.asp"><font color="#000000">�׾·|��</font></a> |
&nbsp;<a href="pm.asp"><font color="#000000">²�T�޲z</font></a> |
&nbsp;<a href="my_favorite.asp"><font color="#000000">���`���D</font></a> |
&nbsp;<a href="search.asp"><font color="#000000">�j�M�峹</font></a> |
&nbsp;<a href="index_counter.asp"><font color="#000000">�y�q</font></a> |
&nbsp;<a href=loginout.asp><font color="#000000">�n�X�׾�</font></a><%end if%>
<%
If request.cookies("admin_pass")=bbsid Then%>
| &nbsp;<a href="admin/admin_home.asp" style="text-decoration: none" target=_blank><font color="#000000">�޲z</font></a>
<%End If%>
</td>
<td width="80" height="3" bordercolor="#FFFFFF" background="images/windowbar_lightbg.gif">
<p align="center">�@</td>
<td width="50" height="3" bordercolor="#FFFFFF" background="images/windowbar_lightbg.gif">
<p align="center"><a href="index.asp" target="_top"><img border="0" src="images/gohome.gif" width="46" height="16"></a></td>
</tr></table>
 <iframe src="http://www.ettoday.com/common/ticker.htm" width="540" height="20" marginwidth="0" marginheight="0" frameborder="0" scrolling="no" STYLE="filter: Chroma(color=000000)"></iframe>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FFFFFF" width="99%" height="3" topmargin="0" leftmargin="0">
<tr><td width="100" align="center">
  <!-- �ˬd�O�_��pm -->
<%Set rspm=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM u2u WHERE touser='" & request.cookies("UserID") & "'"
SQL=SQL & "Order By time Desc"
rspm.Open SQL,conn,1,1
i=0
If not rspm.BOF then 
rspm.MoveFirst
end if
while Not rspm.EOF
  i=i+rspm("readed")
   rspm.movenext
wend
if i<>0 then
Response.Write "<a href=pm.asp><img src=""images/newnew.gif""border=0></a>"
response.write "<BGSOUND SRC=""images/mail.wav"" WIDTH=0 HEIGHT=0 LOOP=1>"
end if



%></td>
<td align="center"><a href="index.asp"><img border="0" src="title.jpg"></a></td>
<td width="100" valign="top" align="center"></td></tr>
</table>
</center>


