<HTML><!--#include file="conn.asp"-->
<HEAD>
<TITLE> admin </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""><link rel="stylesheet" type="text/css" href=../style1.css>
</HEAD>
<%
Set rs = Server.CreateObject("ADODB.Recordset") 
Select Case request("act") '���Ϫ��e�X����
Case "html"
 If request("htmlencode")="yes" Then'��ܱҰ�
   SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("htmlencode")=True
   rs.Update
   rs.close
   Response.Redirect "admin.asp"
 Else
   SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("htmlencode")=False
   rs.Update
   rs.close	
   Response.Redirect "admin.asp"
 End If
Case "UBBcode"
  If request("UBBcode")="yes" Then'��ܱҰ�
   SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("UBBcode")=True
   rs.Update
   rs.close
   Response.Redirect "admin.asp"
  Else
   SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("UBBcode")=False
   rs.Update
   rs.close	
   Response.Redirect "admin.asp"
  End If
Case "editbbsname"
If request("upload")="yes" Then
  SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("bbsname")=request("bbsname")
   rs("servertimezone")=request("servertimezone")
   rs("word_len")=request("word_len")
   rs("reply_lim")=request("reply_lim")
   rs("reply_counts")=request("reply_counts")
   rs("upload")=true
   rs.Update
   rs.close	
   Response.Redirect "admin.asp"
else
    SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("bbsname")=request("bbsname")
   rs("servertimezone")=request("servertimezone")
   rs("word_len")=request("word_len")
   rs("reply_lim")=request("reply_lim")
   rs("reply_counts")=request("reply_counts")
   rs("upload")=false
   rs.Update
   rs.close	
    Response.Redirect "admin.asp"
end if
Case "en_reg"
If request("en_reg")="yes" Then
  SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("en_reg")=true
   rs.Update
   rs.close	
   Response.Redirect "admin.asp"
else
    SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("en_reg")=false
   rs.Update
   rs.close	
   Response.Redirect "admin.asp"
end if
Case "en_save"
If request("en_save")="yes" Then
  SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("en_save")=true
   rs.Update
   rs.close	
   Response.Redirect "admin.asp"
else
    SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("en_save")=false
   rs.Update
   rs.close	
   Response.Redirect "admin.asp"
end if
Case "bbs_close"
If request("bbs_en")="yes" Then
  SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("close_why")=request("why")
   rs("bbs_close")=true
   rs.Update
   rs.close	
    Response.Redirect "admin.asp"
else
    SQL = "Select * From config_setting where ID=1"
   rs.Open SQL, Conn,1,3
   rs("close_why")=request("why")
   rs("bbs_close")=false
   rs.Update
   rs.close	
    Response.Redirect "admin.asp"
end if
Case "edittid"
   SQL = "Select * From tid where tid=" & request("ID")
   rs.Open SQL, Conn,1,3
   rs("num")=request("num")
   rs("tidname")=request("forumname")
   tidadmin="|"&replace(request("forumadmin"),"||","|")&"|"
   rs("tidadmin")=replace(tidadmin,"||","|")
   rs("text")=request("text")
   rs("have_pass")=request("have_pass")
   rs("password")=request("forum_pass")
   rs.Update
   rs.close	
    Response.Redirect "admin.asp"
Case "addforum"
    SQL = "Select * From tid "
   rs.Open SQL, Conn,1,3
   rs.Addnew
   rs("tidname")=request("addname")
   tidadmin="|"&replace(request("addadmin"),"||","|")&"|"
   rs("tidadmin")=replace(tidadmin,"||","|")
   rs("text")=request("addtext")
   rs("sort")=request("sort")
   rs.update
   rs.close
    Response.Redirect "admin.asp"
Case "deleteforum"
    SQL = "Select * From tid where tid= " & request("id")
    rs.Open SQL, Conn,1,3
   rs.delete
   rs.close
    sql="select * from Titles where tid="  & request("id")
     rs.Open SQL, Conn,1,3'�R���D�D
	 while not rs.eof
	 rs.delete
	 rs.movenext
	 wend
	rs.close
     sql="select * from Details where tid="  & request("id")
     rs.Open SQL, Conn,1,3'�R���^��
	 while not rs.eof
	 rs.delete
	 rs.movenext
	 wend
	rs.close
    response.write "&nbsp;&nbsp;�Q�ת��R������!"
Case Else
   Response.write ""
End Select%>

<BODY leftmargin="0" topmargin="0"><center>
<%   sql="select * from config_setting where ID=1"
     Set rsidid=conn.execute(sql)
       bbsid=Cstr(rsidid("bbsid"))
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else
	SQL = "Select * From config_setting where ID=1"
	rs.Open SQL, Conn,1,1%>
<br>
<TABLE border="1" bordercolor="#C0C0C0" style="border-collapse: collapse" cellpadding="0" cellspacing="0" width="98%">
<tr><TD background="images/title.gif" width="100%" colspan="2"><b><font size="3" align="center">�ֳt�޲z</font></b></td></tr>
<tr><TD bordercolor="#F4FAFF" bgcolor="#F4FAFF" width="100%" colspan="2"><font size="2">&nbsp;&nbsp; </font></td></tr>	
<form method="GET">
<tr>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<font size="3">�׾¦W�١G</font><INPUT TYPE="text" NAME=bbsname Value="<%=rs("bbsname")%>" size="20"></td>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<%If rs("upload")=True Then%>
	<font size="2">�W�Ǫ���\��G�}��</font><INPUT TYPE="radio" NAME="upload" value="yes" checked>
	<font size="2">���}��</font><INPUT TYPE="radio" NAME="upload" value="no">
	<%Else%>
	<font size="2">�W�Ǫ���\��G�}��</font><INPUT TYPE="radio" NAME="upload" value="yes" >
	<font size="2">���}��</font><INPUT TYPE="radio" NAME="upload" value="no" checked>
	<%End If%></td></tr> 
<tr>	
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<font size="2">�峹�̰��r�ƭ���G</font><INPUT TYPE="text" NAME=word_len Value="<%=rs("word_len")%>" size="8"></td>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<font size="2">�^�г̤֦r�ƭ���G</font><INPUT TYPE="text" NAME=reply_lim Value="<%=rs("reply_lim")%>" size="8"></td></tr>
<tr>	
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<font size="2">������ܦ^�мƥءG</font><INPUT TYPE="text" NAME=reply_counts Value="<%=rs("reply_counts")%>" size="8"></td>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<font size="2">���A�����ɰ�վ�G</font><INPUT TYPE="text" NAME=servertimezone Value="<%=rs("servertimezone")%>" size="8"></td></tr>
<tr>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<INPUT TYPE="hidden" NAME=act Value="editbbsname"><INPUT TYPE="submit" VALUE="�H�W�T�w"></td>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF">
	<font color=red size="2"><%response.write "Server Realtime�G"&now&""%></font>�@</td></tr>
<tr><TD bordercolor="#F4FAFF" bgcolor="#F4FAFF" width="100%" colspan="2"><font size="2">&nbsp;&nbsp; </font></td></tr>		
</form>

<tr>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF"><form method="GET"><font size="2">
	<%If rs("en_reg")=True Then%>
	�j�����U�G�}��<INPUT TYPE="radio" NAME="en_reg" value="yes" checked>���}��<INPUT TYPE="radio" NAME="en_reg" value="no">
	<%Else%>
	�j�����U�G�}��<INPUT TYPE="radio" NAME="en_reg" value="yes" >���}��<INPUT TYPE="radio" NAME="en_reg" value="no" checked>
	<%End If%>
	<input type=hidden name=act value="en_reg"><INPUT TYPE="submit" VALUE="�T�w"></font></form></td>	
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF"><form method="GET"><font size="2">
	<%If rs("htmlencode")=True Then%>
	HTML�G�}��<INPUT TYPE="radio" NAME="htmlencode" value="yes" checked>���}��<INPUT TYPE="radio" NAME="htmlencode" value="no">
	<%Else%>
	HTML�G�}��<INPUT TYPE="radio" NAME="htmlencode" value="yes" >���}��<INPUT TYPE="radio" NAME="htmlencode" value="no" checked>
	<%End If%>
	<input type=hidden name=act value="html"><INPUT TYPE="submit" VALUE="�T�w"></font></form></td></tr>
	
<tr>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF"><form method="GET"><font size="2">
	<%If rs("en_save")=True Then%>
	�峹�O�s�����G�}��<INPUT TYPE="radio" NAME="en_save" value="yes" checked>���}��<INPUT TYPE="radio" NAME="en_save" value="no">
	<%Else%>
	�峹�O�s�����G�}��<INPUT TYPE="radio" NAME="en_save" value="yes" >���}��<INPUT TYPE="radio" NAME="en_save" value="no" checked>
	<%End If%>
	<input type=hidden name=act value="en_save"><INPUT TYPE="submit" VALUE="�T�w"></font></form></td>
	<TD width="50%" bordercolor="#F4FAFF" bgcolor="#F4FAFF"><form method="GET"><font size="2">
	<%If rs("UBBcode")=True Then%>
	UBBcode�G�}��<INPUT TYPE="radio" NAME="UBBcode" value="yes" checked>���}��<INPUT TYPE="radio" NAME="UBBcode" value="no">
	<%Else%>
	UBBcode�G�}��<INPUT TYPE="radio" NAME="UBBcode" value="yes" >���}��<INPUT TYPE="radio" NAME="UBBcode" value="no" checked>
	<%End If%>
	<input type=hidden name=act value="UBBcode"><INPUT TYPE="submit" VALUE="�T�w"></font></form></td></tr>
	
<tr>
	<TD width="100%" bordercolor="#F4FAFF" bgcolor="#F4FAFF" colspan="2"><form method="GET"><font size="2">
	<%If rs("bbs_close")=True Then'�þ�����%>
	�׾¡G����<INPUT TYPE="radio" NAME="bbs_en" value="yes" checked>������<INPUT TYPE="radio" NAME="bbs_en" value="no">
	<%Else%>
	�׾¡G����<INPUT TYPE="radio" NAME="bbs_en" value="yes" >������<INPUT TYPE="radio" NAME="bbs_en" value="no" checked>
	<%End If%>
	&nbsp;&nbsp;&nbsp;&nbsp;������]:<INPUT TYPE="text" NAME="why" size=40 value="<%=rs("close_why")%>">
	<input type=hidden name=act value="bbs_close"><INPUT TYPE="submit" VALUE="�T�w"></font></form></td></tr>


<tr><TD width="100%" bordercolor="#F4FAFF" bgcolor="#F4FAFF" colspan="2">
<form method="GET">
<center>
<table border="1" cellspacing="0" style="border-collapse: collapse; font-size: 10 pt" bordercolor="#C0C0C0" id="AutoNumber1" bgcolor="#E1E0BB" width="100%">
	<tr>
		<td width="100%" colspan="5" align="center" style="font-size: 11 pt; color: #FFFFFF" bgcolor="#E1E0BB" background="images/title.gif">
        <font size="3" color="#000000"><b>�s�W�Q�ת�</b></font></td>
	</tr>
	<tr><input type=hidden name=act value="addforum">
        <td style="color: #FFFFFF" bgcolor="#F5F5F5" align="center" width="151">
        <font color="#000000">���W</font></td>
        <td style="color: #FFFFFF" bgcolor="#F5F5F5" align="center" width="160">
        <font color="#000000">���D</font></td>
        <td style="color: #FFFFFF" bgcolor="#F5F5F5" align="center" width="143">
        <font color="#000000">�y�z</font></td>
        <td style="color: #FFFFFF" bgcolor="#F5F5F5" align="center" width="280">
        <font color="#000000">����</font></td>
        <td style="color: #FFFFFF" bgcolor="#F5F5F5" align="center" width="38">
        <font color="#000000">�ʧ@</font></td>
      </tr>
      <tr>
        <td bgcolor="#F5F5F5" width="151">
        <input type="text" name="addname" size="20"></td>
        <td bgcolor="#F5F5F5" width="160">
        <input type="text" name="addadmin" size="20"></td>
        <td bgcolor="#F5F5F5" width="143">
        <input type="text" name="addtext" size="20"></td>
        <td bgcolor="#F5F5F5" width="280">
        <select name="sort"><%sql="select * from sort"
        set rs_sort=conn.execute(sql)
        while not rs_sort.eof 
        %><option><%=rs_sort("sort_name")%></option>�@<%
        rs_sort.movenext
        wend%></select></td>
        <td bgcolor="#F5F5F5" width="38"><input type="submit" value="�s�W" name="b1"></td>
      </tr>
</table></center>
</form></td></tr>

<tr><TD bordercolor="#F4FAFF" bgcolor="#F4FAFF" width="100%" colspan="2"><font size="2">&nbsp;&nbsp; </font></td></tr>	
<center> 
<tr><TD width="100%" bordercolor="#F4FAFF" bgcolor="#F4FAFF" colspan="2">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="100%" height="1" id="AutoNumber1">
	<tr>
          <td width="34" height="5" bgcolor="#F0EEDF" background="images/title.gif">
          <font size="2">����</font></td>
          <td width="116" height="5" bgcolor="#F0EEDF" background="images/title.gif">
          <font size="2">���W</font></td>
          <td width="112" height="5" bgcolor="#F0EEDF" background="images/title.gif">
          <font size="2">���D</font></td>
          <td width="180" height="5" bgcolor="#F0EEDF" background="images/title.gif">
          <font size="2">����</font></td>
          <td width="41" height="5" bgcolor="#F0EEDF" background="images/title.gif" align="center">
          <font size="2">����</font></td>
          <td background="images/title.gif" width="70" align="center">
          <font size="2">�ݩ�</font></td>
          <td background="images/title.gif" width="40" align="center">
          <font size="2">�ʧ@</font></td>
	</tr>
<%sql="select * from sort order by sort_order"
	set rs_sort1=conn.execute(sql)
if not rs_sort1.EOf then
while not rs_sort1.EOf
    Set forumrs = Server.CreateObject("ADODB.Recordset")
    SQL = "Select * From tid where sort='"&rs_sort1("sort_name")&"' order by num"
    forumrs.Open SQL, Conn,1,1%>
	<tr >
      <td width="100%" colspan="7" align="center" style="font-size: 11 pt; color: #FFFFFF" bgcolor="#E1E0BB" background="images/title.gif"><%=rs_sort1("sort_name")%>�@</td>
	</tr>
<%While Not forumrs.EOF%>
<form method="GET">
	<tr>  
        <td width="34" align="center">
        <font size="2"><%=forumrs("num")%></font>�@</td>
        <td width="116" height="50" bgcolor="#F5F5F5">
        <font size="2"><a href=edit_forum.asp?id=<%=forumrs("tid")%>><%=forumrs("tidname")%></a></font></td>
        <td width="112" height="46" bgcolor="#F5F5F5">
        <font size="2"><%=forumrs("tidadmin")%></font></td>
        <td width="180" height="1" bgcolor="#F5F5F5">
        <font size="2"><%=forumrs("text")%></font></td>
        <td width="41" height="1" bgcolor="#F5F5F5" align="center">
        <font size="2"><%=forumrs("have_pass")%></font></td>
        <td width="70">
        <font size="2"><%=forumrs("sort")%></font>�@</td>
        <td width="40"> 
		<INPUT TYPE="hidden" name="act" VALUE="deleteforum">
		<INPUT TYPE="hidden" name="id" VALUE=<%=forumrs("tid")%>>
		<INPUT TYPE="submit" VALUE="�R��"></td>
	</tr>
</form>
<%forumrs.MoveNext
  Wend
	 
  rs_sort1.movenext
  wend%>

<%End If%>
</table></td></tr></center>


<tr><TD width="100%" bordercolor="#F4FAFF" bgcolor="#F4FAFF" colspan="2">
<center>
<br><FONT SIZE=-1 color="#FF0000">
����:���ǬO�������ƦC����,�Ʀr�V�j�V�U��,�h���D�Х�||�Ϲj,�Ҧp|supercjb|hack|
<p></p>
�Y�ݱK�X�n�J���Ͻ׾�,�Цb�������W�ȳ]��1,�å��W�K�X
<p></p>
�Y���ݭn�K�X�n�J,�бN�������W�ȳ]��0,���n���W�K�X
<p></p>
����R���|��Ӫ����D�D�H�Φ^�Ф]�q�q�R��,�ЯS�O�d�N!!
</font>
</center></td></tr>
</table> 
<!--#include file="../foot.asp"-->
<%End If%></BODY></HTML>