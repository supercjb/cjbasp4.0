<!--#include file="conn.asp"-->
<% function sum_anything(sql,rs)
	Set rs = conn.Execute(sql)
		i=0
	while not rs.eof
		i=i+1
		rs.movenext
	wend
	Response.Write "<font color=red><b>"& i &"</b></font>"
  end function%>

<HTML>
<HEAD>
<TITLE><%=rsidid("bbsname")%> </TITLE>



<meta http-equiv="refresh" content="1200; url=index.asp">  
<META NAME="Description" CONTENT="">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href=style1.css>
</HEAD>
<BODY leftmargin="0" topmargin="0">

  <!-- �n�� -->

<!--#include file="header.asp"-->

<!--#include file="admin/broad.asp"-->
<%if rstitle("bbs_close")=true then'���Ͻ׾�����
response.write "<br><center><b>��p!�׾¦]���H�U��]�Ȯ�����!�Фŵo��峹�έק���,�H�K�y����!"
response.write "<br><font color=red>"&rstitle("close_why")&"</font></b></center>"
elseif rstitle("bbs_close")=false then'���ϨS����
%>

  <CENTER>
    <FONT COLOR=RED size=5><b><%=Request("msg")%></b></FONT>
  </CENTER>
<p align="right" style="margin-top: 0; margin-bottom: 0">
         <font size="2">�w��s�i���� 
</font><%
SQL= "Select * From config_setting where ID=1 "
Set registrs=conn.Execute(SQL)
%>
<font color=red size=-1>"<%=registrs("newregist")%>"&nbsp;</font><font size=-1>�[<a name="top">�J</a></font> 
<td width="9">

   

<font size=-1> <br>
  
    �@<% 
StrA = "select * from id order by ID"
StrB = "select * from Titles order by TitleID"
StrC = "select * from Details order by DetailID"

Response.Write "���׾¦@�p��" 
sum_anything StrA,rss
Response.Write "����U�|��,�æ�"
sum_anything StrB,rss1
Response.Write "�g�D�D�H��"
sum_anything StrC,rss2
Response.Write "�g�^�Ф峹�C"
    %>
   </font></p><center>
   <!--#include file="top_new.asp"-->
<table border="1" cellpadding="0" cellspacing="0" style="word-break:break-all; border-collapse:collapse" bordercolor="#C4E0FF" width="98%" height="68" id="AutoNumber1" bordercolorlight="#F5F5F5" bordercolordark="#C0C0C0" STYLE="WORD-BREAK: break-all" background="images/postbit_lightbg.gif">   
<%
	UserID=request.cookies("UserID")
	SQL = "Select * From id Where user='"&request.cookies("UserID")&"'"
    set rs11=conn.execute(SQL)
	If request.cookies("UserPassed")<>"yes" or rs11.eof Then 
%><tr><td width="98%" background="images/images/postbit_lightbg.gif" colspan="6" width="5"
bordercolor="#D1D1D1" height="11"><p align="left"><FORM METHOD=POST ACTION="check.asp">
<% 
   response.cookies("TEST_COOKIE").Expires=date+365
response.cookies("TEST_COOKIE") = "testing" '�n���e����!�ˬdCOOKIE%>
<font size="2"><b><img border="0" src="images/login.gif">�|���n��</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
�Τ�b���G<INPUT TYPE="text" NAME="UserID" size="10">&nbsp;&nbsp;&nbsp;&nbsp;�b���K�X�G<INPUT TYPE="password" Name="Password" size="9">&nbsp;
<select size="1" name="cookie_save_day">
<option selected value="365">�û�</option>
<option value="30">�@�Ӥ�</option>
<option value="7">�@�P</option>
<option value="1">�@��</option>
<option value="0">���ΰO��</option>
</select><INPUT TYPE="submit" VALUE="�n�J"><INPUT TYPE="reset" VALUE="����">&nbsp;&nbsp;
(�Y�O�L�k�n�J�A�Э��CIE���p�v�]�w�A��<font color="#FF0000">�`�N�j�p�g</font>)</font></FORM></p></td></tr><% End If %>
    <tr> 
		<td width="5%" background="images/Button.gif" height="13" bordercolor="#E3E3E3">
        <p align="center"><b><font size="2">���A</font></b></td>
        <td width="40%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"> 
            <b> 
            <font size="2" color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;  </font> 
            <font size="2">�����W��&nbsp;&nbsp;</font><font size="2" color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; </font><font size="2">&nbsp; </font></b></td>
        
        <td width="12%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><b><font size="2">���D</font></b></td>
        <td width="6%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><font size="2"><b>�D�D��</b></font></td>
        <td width="6%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><font size="2"><b>�^�м�</b></font></td>
        <td width="18%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><font size="2"><b>�̫��s</b></font></td>
      </tr>

<%
sql="select * from sort order by sort_order"
set rs_sort=conn.execute(sql)
if not rs_sort.EOf then
while not rs_sort.EOf
%> 
<tr><td width="98%"  colspan="6" width="5" bgcolor="#FAFAFA"
bordercolor="#D1D1D1" height="11"><font size=-1><b>��<%=rs_sort("sort_name")%>��</b></font></td></tr>
<%
SQL= "Select * From tid where sort='"&rs_sort("sort_name")&"' Order By num "
Set rs=conn.Execute(SQL)
If rs.Eof Then 
	
Else
 While Not rs.EOF 
     ' ��ܸ��
%>
      <tr> 
        <td width="3%" align="center" bgcolor="#F5F5F5" bordercolor="#D1D1D1" height="42" background="images/images/postbit_lightbg.gif">
      <%IF Now - rs("lastnewtime") < 1 then 
 	response.write "<img src=""images/icon_folder_new_3.gif"">"
else
    response.write "<img src=""images/icon_folder_3.gif"">"
end if%> �@</td>
        <td width="45%" height="42" bgcolor="#FFFFFF" bordercolor="#D1D1D1" align="left" background="images/postbit_lightbg.gif"> 
        <p align="left"><font color="#000080" size="2">�@</font>
        <font color="#FF9900" size="2"><a href=forum.asp?tid=<%=rs("tid")%>>
        <font color="#CC6666"><b><%=rs("tidname")%></b></font></a></font> 
          <%if rs("have_pass")=1 then %><img src="images/key.gif" alt="�ݭn�K�X"><%end if%> <font size=-1><br><%=rs("text")%></font> </td>
        
<td width="8%" height="42" bgcolor="#FFFFFF" align="center" bordercolor="#D1D1D1"> 
<font color="#000080" size="2">�@
<% 
if rs("tidadmin")<>"" then
tidadmin=split(rs("tidadmin"),"|")
for i = 1 to ubound(tidadmin)-1
response.write tidadmin(i)&"<br>"
next
end if
%></font></td>
        <td width="49" height="42" bgcolor="#F5F5F5" align="center" bordercolor="#D1D1D1" background="images/postbit_lightbg.gif"> 
          <font size="2" color=red><b><%=rs("counts")%></b></font></td>
        <td width="7%" height="42" bgcolor="#FFFFFF" align="center" bordercolor="#D1D1D1"> 
          <font size="2" color=red><b><%=rs("replycounts")%></b></font></td>
        <td width="14%" height="42" bgcolor="#EFF4FE" bordercolor="#D1D1D1" background="images/postbit_lightbg.gif"> 
        <p align="left"><font color="black" size="2">
		<%
'**************************************************************************
'�P�_�̷s��ܪ��v��
sql="select * from tid where tid="&rs("tid")
set rs_forum=conn.execute(sql)
forum_cookies=bbsid+Cstr(rs("tid"))+request.cookies("UserID")
if request.cookies(forum_cookies)="yes" or  rs_forum("have_pass")=0 then	
'***************************************************************************
		%>
		<font color="#003399">�峹�G</font><%if rs("lasttitle")<>0 then
		  sql="select * from Titles where TitleID="&rs("lasttitle")
                   set rslast=conn.execute(sql)
                   if not rslast.EOF then
                   response.write "<a href=detail.asp?TitleID="&rs("lasttitle")&"&tid="&rs("tid")&"#top>"&Left(rslast("Subject"),5)&"..."&"</a>"
                   end if
		       elseif rs("lasttitle")=0 then
			     response.write Left(rs("lasttitle2"),5)&"..."
			end if
				
        %><br><font color=red>�̫�G</font><%=rs("lastposter")%>&nbsp;<img src="images/lastpost.gif">
        <br></font><font color="#7C7C7C" size=-1><%=rs("lastnewtime")%></font>
		<%
'**************************************************************************
else
  response.write "��p,�z�L�v��"
end if
'***************************************************************************
		%>
		
		</td>
      </tr>
      <%
	  rs.MoveNext
	  Wend
	  END IF
	  	  rs_sort.movenext
	  wend

	  else
	    response.write "<p align=center><b><font color=red>���L�D�׾¤���!!</font></b></p>"
	  end if
	  %>
    </table>

 
     
    </select>
  </p>
 

<!--#include file="UserList.asp"-->
 <!--#include file="friend_web.asp"-->
<div align="center">
  <center>
  <br> <br> <br>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="98%" height="22">
    <tr>
      <td width="524" height="22" bordercolor="#FFFFFF" bgcolor="#EFF4FE" background="images/postbit_lightbg.gif"><font size=-1><b><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
%></b></font></td>
      <td width="5%" height="22" bgcolor="#EFF4FE" background="images/postbit_lightbg.gif">
<a href="#top"><img border="0" src="images/up.gif" width="20" height="20"></a></td>
    </tr>
  </table>
  </center>
</div>
<%end if'�׾�������endif%>
   
  <!--#include file="foot.asp"-->



</BODY>
</HTML>