<center><table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="98%" height="6" id="AutoNumber1" bgcolor="#EFF4FE" bordercolorlight="#C0C0C0" bordercolordark="#C0C0C0">
<tr><td width="700" height="5" bordercolor="#C0C0C0" background="images/Button.gif"><font color="#333333"size=-1><b> ◎ 個 人 訊 息</b></font></td></tr>
<tr><td width="700" height="4" bordercolor="#C0C0C0" align="left">
<font size="-1" ><img border="0" src="images/icon_online.gif" width="16" height="16">您的網路 IP 是</FONT><font size="-1" color=red> :<%=Request.ServerVariables("REMOTE_ADDR")%>　</font>
<font size=-1>
<%strsplit=split(Request.Servervariables("HTTP_USER_AGENT"),";")
strsplit(1)=replace(strsplit(1),"MSIE","Internet Explorer")
strsplit(2)=replace(strsplit(2),")","")
strsplit(2)=replace(strsplit(2),"NT 5.1","XP")
strsplit(2)=replace(strsplit(2),"NT 5.0","2000")
strsplit(1)=Trim(strsplit(1))
strsplit(2)=Trim(strsplit(2))
response.write "，瀏覽器為"&strsplit(1)
response.write "&nbsp;&nbsp;，作業系統為"&strsplit(2)%></FONT>
<!-- 檢查是否有pm -->
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
Response.Write "<br><font color=red size=-1><a href=pm.asp>您有訊息尚未閱讀</a></font>，"
else
Response.Write "<br><font color=red size=-1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;您無新的訊息</font>，"
end if
if request.cookies(bbsid)("UserID")<>empty then
SQL="SELECT * FROM id where user='"&request.cookies(bbsid)("UserID")&"'"
set rsuser=conn.Execute(SQL)
RESPONSE.WRITE "<font size=-1>共發表過<font color=red><b>"&rsuser("postsum")&"</b></font>篇文章。</font>"
else
RESPONSE.WRITE "<font size=-1>共發表過<font color=red><b>0</b></font>篇文章。</font>"
end if%></td></tr>
</table>
</center>