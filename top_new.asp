<%sql="select * from config_setting"
set rstitle=conn.execute(sql)
sql="select Top 5 * from Titles order by TitleID desc"
set rs_top=conn.execute(sql)
sql="select Top 6 * from id order by ID desc"
set rs_id=conn.execute(sql)
sql="select * from id where user='"&request.cookies("UserID")&"'"
set rsus=conn.execute(sql)%><center>
<TABLE border="1" cellpadding="0" cellspacing="0" style="word-break:break-all; border-collapse:collapse" bordercolor="#C4E0FF" width="98%" height="65" id="AutoNumber1" bordercolorlight="#F5F5F5" bordercolordark="#C0C0C0" STYLE="WORD-BREAK: break-all">
<TR ><TD background="images/Button.gif" width="12%" align="center"><font size=2><b>�� �ӤH�T��</b></TD>
	<TD background="images/Button.gif" width="27%"></TD>
	<TD background="images/Button.gif" width="26%"></TD>
	<TD background="images/Button.gif" width="22%"><font size=2><b>&nbsp;&nbsp;�̷s�o���D�D</b></font></TD>
	<TD background="images/Button.gif" width="13%" align="center"><font size=2><b>�s�i����</b></TD></TR>
<TR><TD bgcolor=#EFF4FE height="100" background="images/postbit_lightbg.gif"><%if rsus.eof then%><font size=-1><font color=red>�z�|�����U�A�л��ֵ��U�[�J�|��~~</font>
<%else%>     
  <%if rsus("pic")=0 then%>     
    &nbsp;<img src="<%=rsus("face_url")%>" width=100 height=100 border=1>&nbsp;     
  <%else%>     
     &nbsp;<img src="face/<%=rsus("pic")%>.gif" width=100 height=100 border=1>&nbsp;     
  <%end if%>     
<%end if%></font></TD>     
<TD bgcolor=#EFF4FE height="100" background="images/postbit_lightbg.gif"><%if not rsus.eof then%>     
<%response.write "<font size=-1>�z�n <b><font color=red>"&rsus("user")&"~~</font></b>"%>     
<%sql="select * from member_sort where level="&rsus("level")     
set rs_mem=conn.execute(sql)     
 response.write "<br><font size=-1>�z�����šG<b><font color=red>"&rs_mem("levelname")&"</font></font></b>"  
 response.write "<br>�{�b�ɨ�G"&DateAdd("h",rstitle("servertimezone")+rsus("timezone"),now)&""   
'''''''''''''''''''''''''''''''''''''''''''''''pm     
Set rspm=Server.CreateObject("ADODB.Recordset")     
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
response.write "<font size=-1><br>���U����G"&rsus("registtime")     
if i<>0 then     
Response.Write "<br><a href=pm.asp>�T���G�z���T���|���\Ū</a>"     
else     
Response.Write "<br>�T���G�z�L�s���T��"     
end if     
if request.cookies("UserID")<>empty then     
SQL="SELECT * FROM id where user='"&request.cookies("UserID")&"'"     
set rsuser=conn.Execute(SQL)     
RESPONSE.WRITE "<br>�峹�G�z�@�o��L <font color=red><b>"&rsuser("postsum")&"</b></font> �g�峹�C"     
else     
RESPONSE.WRITE "<br>�峹�G�z�@�o��L <font color=red><b>0</b></font> �g�峹�C"     
end if
else
response.write "<center><font size=-1>�{�b�ɨ�G"
response.write "<center><font size=-1>"&DateAdd("h",rstitle("servertimezone"),now)&""
response.write "<br>�L��ܸ�T"     
response.write "<br><a href=registrule.asp><b>'���U�����|��'</b></a>"     
end if%></TD>     
<TD bgcolor="#FFFFFF" height="100"><center>     
<%strsplit=split(Request.Servervariables("HTTP_USER_AGENT"),";")     
strsplit(1)=replace(strsplit(1),"MSIE","Internet Explorer")     
strsplit(2)=replace(strsplit(2),")","")     
strsplit(2)=replace(strsplit(2),"NT 5.1","XP")     
strsplit(2)=replace(strsplit(2),"NT 5.0","2000")     
strsplit(1)=Trim(strsplit(1))     
strsplit(2)=Trim(strsplit(2))     
'response.write "<font size=-1>&nbsp;&nbsp;�s�����G"&strsplit(1)&"</font><br>"     
'response.write "<font size=-1>&nbsp;&nbsp;�@�~�t�ΡG"&strsplit(2)&"</font><br>"     
response.write "<font size=-1>�z������IP�G"&Request.ServerVariables("REMOTE_ADDR")&"</font>&nbsp;&nbsp;<br>"%></center>  
<%if not rsus.eof then%>     
	<img src=images/post/02.gif border=0 width=13 height=13>     
	<%response.write "<a href=search_action2.asp?sort=aa&name="&UserID&"&sortx=1><font size=-1>�z�o���D�D(post_time)</font></a>"%><br>     
	<img src=images/post/02.gif border=0 width=13 height=13>     
	<%response.write "<a href=search_action2.asp?sort=aa&name="&UserID&"&order=1&sortx=2><font size=-1>�z�o���D�D(reply_time)</font></a>"%><br>     
	<img src=images/post/02.gif border=0 width=13 height=13>     
	<%response.write "<a href=search_action2.asp?sort=aa&replycount=0&order=1&sortx=3><font size=-1>�˵��|���^�Ъ��Ҧ��D�D</font></a>"%><br>     
	<img src=images/post/02.gif border=0 width=13 height=13>     
	<%response.write "<a href=search_action2.asp?sort=aa&timediff=7&order=1&sortx=4><font size=-1>�˵��@�P�������Ҧ��D�D</font></a>"%><br>     
<%end if%></TD>     
	<td bgcolor="#EFF4FE" height="100" background="images/postbit_lightbg.gif">       
	<%while not rs_top.eof%>&nbsp;<img src=images/post/<%=RS_top("R1")%>.gif border=0 width=13 height=13>     
	 <%response.write "<a href=detail.asp?TitleID="&rs_top("TitleID")&"&tid="&rs_top("tid")&"&postname="&rs_top("Name")&"><font size=-1>"&Left(rs_top("subject"),9)&"."&"</font></a><br>"     
	 rs_top.movenext   
	wend%></td>  
	<TD bgcolor="#FFFFFF" align="center" height="100">  
	<%while not rs_id.eof     
	 response.write "<a href=edituser.asp?UserID="&rs_id("user")&"><font size=-1>"&rs_id("user")&"</font></a><br>"     
	 rs_id.movenext     
	wend%></TD></TR>     
</TABLE>