<!--#include file="conn.asp"-->
<!--#include file="safe.asp"-->
<HTML><HEAD>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
bbsid=Cstr(rstitle("bbsid"))%> </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<link rel="stylesheet" type="text/css" href=style1.css></HEAD>
<BODY leftmargin="0" topmargin="0">
<%SQL = "Select * From id Where user='"&request.cookies("UserID")&"'"
    set rs11=conn.execute(SQL)
If request("sort")="1" then '向下排
	nextsort="2"
	
	if request("ordery")=1 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By postsum Desc",conn,1,1

	elseif request("ordery")=2 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By ID Desc",conn,1,1

	elseif request("ordery")=3 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By month(birthday) Desc,day(birthday) Desc",conn,1,1

	elseif request("ordery")=4 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By comefrom Desc",conn,1,1

	elseif request("ordery")=5 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By user Desc",conn,1,1

'	elseif request("ordery")=6 then
'	Set rs=Server.CreateObject("ADODB.Recordset")
'	rs.Open "Select * From id Where postsum=0 Order By ID Desc",conn,1,1

	end if
	
Else '向上排	
	nextsort="1"

	if request("ordery")="" then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By ID Desc",conn,1,1

	elseif request("ordery")=1 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By postsum",conn,1,1

	elseif request("ordery")=2 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By ID",conn,1,1

	elseif request("ordery")=3 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By month(birthday),day(birthday)",conn,1,1

	elseif request("ordery")=4 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By comefrom",conn,1,1

	elseif request("ordery")=5 then
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From id Order By user",conn,1,1

'	elseif request("ordery")=6 then
'	Set rs=Server.CreateObject("ADODB.Recordset")
'	rs.Open "Select * From id Where postsum=0 Order By ID",conn,1,1

	end if

End If%>
<!--#include file="header.asp"--><!--#include file="admin/broad.asp"--><br>
<% response.cookies("tid")="會員列表"%><!--#include file="choice.asp"--><br><br>
<%'處理修改以及刪除的程序
Select Case request("act")
 Case "edit"
    response.redirect "edituser.asp?UserID=" & Request("ID")
 Case "delete"	
  if request.cookies("admin_pass")=bbsid then
    Set editrs=Server.CreateObject("ADODB.Recordset")
    editrs.Open "Select * From id Where ID=" & Request("ID"),conn,1,3
    editrs.delete
	editrs.close
	response.redirect "user.asp"
  else
	response.redirect "index.asp?msg=無權限!"
  end if
Case Else
   Response.write ""
End Select%>
<center>
<table border="1" cellpadding="0" cellspacing="0" style="word-break:break-all; border-collapse:collapse" bordercolor="#000000" width="92%" height="68" id="AutoNumber1" bordercolorlight="#F5F5F5" bordercolordark="#C0C0C0" STYLE="WORD-BREAK: break-all">
	<tr>
	<td width="5%" height="8" bgcolor="#FFFFFF" bordercolor="#FFFFFF" align="left">
	<p align="center">　</td>
	<td width="5%" height="8" bgcolor="#FFFFFF" bordercolor="#FFFFFF" align="left">
	<font color="#008080" size="2">&nbsp; </font></td>
	<td width="16%" height="8" bgcolor="#F0EEDF" bordercolor="#808080" background="images/Button.gif">
	<p align="center"><font color="#008080" size="2"><a href="user.asp?sort=<%=nextsort%>&ordery=5">
	<font color="#000000">帳號</font></a></font></td>
	<td width="10%" height="8" bgcolor="#F0EEDF" bordercolor="#808080" background="images/Button.gif">
	<p align="center"><font color="#008080" size="2"><a href="user.asp?sort=<%=nextsort%>&ordery=4">
	<font color="#000000">來自何方</font></a></font></td>
	<td width="22%" height="8" bgcolor="#F0EEDF" bordercolor="#808080" background="images/Button.gif">
	<p align="center"><font size="2">個人通訊</font></td>
	<td width="10%" height="8" bgcolor="#F0EEDF" bordercolor="#808080" background="images/Button.gif">
	<p align="center"><font color="#008080" size="2"><a href="user.asp?sort=<%=nextsort%>&ordery=3">
	<font color="#000000">生日</font></a></font></td>
	<td width="19%" height="8" bgcolor="#F0EEDF" bordercolor="#808080" background="images/Button.gif">
	<p align="center"><font color="#008080" size="2"><a href="user.asp?sort=<%=nextsort%>&ordery=2">
	<font color="#000000">註冊日期</font></a></font></td>
	<td width="7%" height="8" bgcolor="#F0EEDF" bordercolor="#808080" background="images/Button.gif"> 
	<p align="center"><font color="#008080" size="2"><a href="user.asp?sort=<%=nextsort%>&ordery=1">
	<font color="#000000">發表數</font></a></font></td>
	<td width="6%" height="8" bgcolor="#F0EEDF" bordercolor="#808080" background="images/Button.gif">
	<p align="center"><font size="2">等級</font></td>
    </tr>
    <%Page = cint(Request("Page"))
	  rs.PageSize =15           ' 設定每頁顯示 15 筆
     If Page <=0 Then
        Page = 1
     ELSEIF Page>rs.PageCount Then Page=rs.PageCount
     END IF

     If Not rs.eof Then          ' 有資料才執行 
        rs.AbsolutePage = Page   ' 將資料錄移至 PAGE 頁
     End if%>
<%If rs.Eof Then 
	RESPONSE.WRITE "      <center><FONT COLOR=maroon>Sorry !! 目前尚無會員...</FONT></center>"
Else
 FOR SH=1 to rs.PageSize %>
	<tr> 
	<td width="5%" bgcolor="#FFFFFF" align="center" bordercolor="#FFFFFF">
	<%If request.cookies("admin_pass")=bbsid Then%>
	<form method="GET">
	<input type=hidden name=act value="edit"><input type=hidden name=ID value="<%=rs("user")%>"><INPUT TYPE="submit" VALUE="修改">
	</form><%end if%></td>
	<td width="5%" bgcolor="#FFFFFF" align="center" bordercolor="#FFFFFF">
	<%If request.cookies("admin_pass")=bbsid Then%>
	<form method="GET">
	<input type=hidden name=id value="<%=rs("ID")%>"><input type=hidden name=act value="delete"><INPUT TYPE="submit" VALUE="刪除">
	</form><%end if%></td>
	<td width="16%" bgcolor="#EFF4FE" bordercolor="#000000" align="center" background="images/postbit_lightbg.gif"><font size=-1><a href="edituser.asp?UserID=<%=rs("user")%>"><%=rs("user")%></a></font>　</td>
	<td width="10%" bgcolor="#EFF4FE" bordercolor="#000000" align="center" background="images/postbit_lightbg.gif"><font size=-1><%=rs("comefrom")%></font>　</td>	
	<td width="22%" bgcolor="#FFFFFF" align="center" bordercolor="#000000" background="images/postbit_lightbg.gif"><font size=-1>
	<a href="u2u.asp?touser=<%=rs("user")%>"><img border="0" src="images/pm.gif" alt="傳送簡訊"></a>&nbsp;	
	<a href="mailto:<%=rs("email")%>"><img src="images/red_folder.gif" border="0" alt="寫信"></a>&nbsp;
<%if rs("msnm")<>Empty then%>
	 <a href="edituser.asp?UserID=<%=rs("user")%>">
    <img src="images/icon_msnm.gif" width="50" height="15" border="0" alt="<%=rs("msnm")%>"></a>&nbsp;
<%end if
if rs("yim")<>Empty then%>
	<a target="_blank" href="http://messenger.yahoo.com/edit/send/?.target=<%=rs("yim")%>&.src=pg">
	<img src="images/icon_yim.gif" width="50" height="15" border="0" alt="<%=rs("yim")%>"></a>&nbsp;
<%end if
if rs("homepage")<>Empty then%>
	 <a target="_blank" href="<%=rs("homepage")%>">
    <img src="images/homepage.gif" border="0" alt="<%=rs("homepage")%>"></a>
<%end if%>	
	</font>　</td>
	<td width="10%" bgcolor="#EFF4FE" bordercolor="#000000" align="center" background="images/postbit_lightbg.gif"><font size=-1><%=rs("birthday")%></font>　</td>
	<td width="19%" bgcolor="#FFFFFF" align="center" bordercolor="#000000" background="images/postbit_lightbg.gif"><font size=-1>
				<%if rs11 Is Nothing Or rs11.EOF Then%>  
				<%=DateAdd("h",rstitle("servertimezone"),rs("registtime"))%>
				<%else%>
				<%=DateAdd("h",rstitle("servertimezone")+rs11("timezone"),rs("registtime"))%>
				<%end if%></font>　</td>
	<td width="7%" bgcolor="#FFFFFF" align="center" bordercolor="#000000" background="images/postbit_lightbg.gif"><font size=-1><%=rs("postsum")+rs("replycount")%></font>　</td>
	<td width="6%" bgcolor="#EFF4FE" bordercolor="#000000" align="center" background="images/postbit_lightbg.gif"><font size=-1>
	<%sql="SELECT * FROM member_sort where level="&rs("level")
	  set rs12=conn.execute(sql)
      response.write rs12("levelname")%></font>　</td>
	</tr>
	<%rs.MoveNext
	IF rs.EOF THEN EXIT FOR
    Next
	End If%>
</table>
</center> 
  <p align="right"><font size=-1>
  <%myself=request.serverVariables("PATH_INFO")%>
  頁次<%=Page%>/<%=rs.PageCount%>
    <a href=<%=myself%>?Page=<%=(Page-1)%>&sort=<%=request("sort")%>&ordery=<%=request("ordery")%>>上一頁</a>
  <a href=<%=myself%>?Page=<%=(Page+1)%>&sort=<%=request("sort")%>&ordery=<%=request("ordery")%>>下一頁</a>
<form action=<%=myself%> method=GET>輸入頁次:<INPUT TYPE=TEXT NAME=Page SIZE=3>
<INPUT TYPE=hidden NAME=sort VALUE="<%=request("sort")%>">
<INPUT TYPE=hidden NAME=ordery VALUE="<%=request("ordery")%>">
</form></p></font>
<!--#include file="foot.asp"-->
</BODY></HTML>