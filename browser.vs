<center><table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="98%" height="6" id="AutoNumber1" bgcolor="#EFF4FE" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
<tr><td width="700" height="5" bordercolor="#C0C0C0" background="images/Button.gif"><font color="#333333"size=-1><b> �� �� �H �T ��</b></font></td></tr>
<tr><td width="700" height="4" bordercolor="#C0C0C0" align="left">
<font size="-1" ><img border="0" src="images/icon_online.gif" width="16" height="16">�z������ IP �O</FONT><font size="-1" color=red> :<%=Request.ServerVariables("REMOTE_ADDR")%>�@</font>
<font size=-1>
<%strsplit=split(Request.Servervariables("HTTP_USER_AGENT"),";")
strsplit(1)=replace(strsplit(1),"MSIE","Internet Explorer")
strsplit(2)=replace(strsplit(2),")","")
strsplit(2)=replace(strsplit(2),"NT 5.1","XP")
strsplit(2)=replace(strsplit(2),"NT 5.0","2000")
strsplit(1)=Trim(strsplit(1))
strsplit(2)=Trim(strsplit(2))
response.write "�A�s������"&strsplit(1)
response.write "&nbsp;&nbsp;�A�@�~�t�ά�"&strsplit(2)%></FONT>
<!-- �ˬd�O�_��pm -->
<%Set rspm=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM u2u WHERE touser='" & request.cookies(bbsid)("UserID") & "'"
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
Response.Write "<br><font color=red size=-1><a href=pm.asp>�z���T���|���\Ū</a></font>�A"
else
Response.Write "<br><font color=red size=-1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�z�L�s���T��</font>�A"
end if
if request.cookies(bbsid)("UserID")<>empty then
SQL="SELECT * FROM id where user='"&request.cookies(bbsid)("UserID")&"'"
set rsuser=conn.Execute(SQL)
RESPONSE.WRITE "<font size=-1>�@�o��L<font color=red><b>"&rsuser("postsum")&"</b></font>�g�峹�C</font>"
else
RESPONSE.WRITE "<font size=-1>�@�o��L<font color=red><b>0</b></font>�g�峹�C</font>"
end if%></td></tr>
</table>
</center>