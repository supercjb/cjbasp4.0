<html><!--#include file="conn.asp"-->
<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�s�W����1</title>
</head>
<body>
<%sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else

sql="select * from tid where tid="&request("id")
Set rs= Server.CreateObject("ADODB.Recordset")
rs.Open sql, conn,1,3
if request("act")="edit" then
rs("num")=request("num")
rs("tidname")=request("forumname")
if not rs("sort")=request("sort") then
rs("sort")=request("sort")
end if
tidadmin="|"&replace(request("forumadmin"),"||","|")&"|"
rs("tidadmin")=replace(tidadmin,"||","|")
rs("text")=request("text")
rs("have_pass")=request("have_pass")
rs("password")=request("forum_pass")
rs.update
Response.Redirect "admin.asp"
end if%>
<br>
<%sql="select * from tid where tid="&request("id")
set forumrs=conn.execute(sql)%>
<form method="GET">
<center>
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="374" height="71" id="AutoNumber2">
      <tr>
        <td width="374" height="20" background="images/title.gif">
          <b><font size="2"><img border="0" src="../images/xp.gif">�ק�</font></b></td>
      </tr>
      <tr>
        <td width="374" height="22">
          <font size="2">���ǡG<input name="num" value="<%=forumrs("num")%>" size="1" ></font></td>
      </tr>
      <tr>
        <td width="374" height="12" bordercolor="#FFFFFF"><font size="2">���W�G<input name="forumname" value="<%=forumrs("tidname")%>" size="20"></font></td>
      </tr>
      <tr>
        <td width="374" height="5">
          <font size="2"> </font>

          <font size="2">�ݩ�G<%=forumrs("sort")%>
          <select name="sort"><%sql="select * from sort"
        set rs_sort=conn.execute(sql)
        while not rs_sort.eof 
        %>
		<%if rs("sort")=rs_sort("sort_name") then%>
		<option selected>
		<%else%>
          <option>
		  <%end if%>
		<%=rs_sort("sort_name")%></option>�@<%
        rs_sort.movenext
        wend%></select>
          
          </font></td>
      </tr>
      <tr>
        <td width="374" height="10"><font size="2">���D:<input name="forumadmin" value="<%=forumrs("tidadmin")%>" size="20"></font></td>
      </tr>
      <tr>
        <td width="374" height="11">
          <font size="2">�����G<input name="text" value="<%=forumrs("text")%>" size="44"></font></td>
      </tr>
      <tr>
        <td width="374" height="16">
          <font size="2">���ҡG<input type="text" name="have_pass" size="3" value="<%=forumrs("have_pass")%>">&nbsp;&nbsp;&nbsp; �K�X
          �G<input type="password" name="forum_pass" size="6" value="<%=forumrs("password")%>"></font></td>
      </tr>
      <tr>
        <td width="374" height="21">�@<INPUT TYPE="submit"></td>
      </tr>
    </table>
    <input type="hidden" name="act" value="edit">
    <input type="hidden" name="id" value=<%=forumrs("tid")%>>
    

</center>
</form>       
<center> 
<br><FONT SIZE=-1 color="#FF0000">
����:���ǬO�������ƦC����,�Ʀr�V�j�V�U��,�h���D�Х�||�Ϲj,�Ҧp|supercjb|hack|
<p></p>
�Y�ݱK�X�n�J���Ͻ׾�,�Цb�������W�ȳ]��1,�å��W�K�X
<p></p>
�Y���ݭn�K�X�n�J,�бN�������W�ȳ]��0,���n���W�K�X
</font>
</center>
<%end if%>
</body></html>