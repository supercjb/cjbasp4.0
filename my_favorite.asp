<html><!--#include file="conn.asp"--><!--#include file="safe.asp"-->
<head>
<meta http-equiv="Content-Language" content="zh-tw">
<link rel="stylesheet" type="text/css" href=style1.css>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
sql2="select * from member_sort where level="&request.cookies("level")
if request.cookies("level")=empty then
response.redirect "loginout.asp"
end if
set rs_level2=conn.execute(sql2)%> </TITLE></head>
<%if not request("id")=empty then
sql="select * from favorite where titleid="&request("id")&" and name='"&request.cookies("UserID")&"'"
Set rsa= Server.CreateObject("ADODB.Recordset")
rsa.Open sql, conn,1,3
end if
if request("mode")="del" then
rsa.delete
rsa.close
response.redirect "my_favorite.asp"
end if%>
<body  leftmargin="0" topmargin="0">
<!--#include file="header.asp"--><!--#include file="admin/broad.asp"-->
<% response.cookies("tid")="我的關注話題"%><br>
<!--#include file="choice.asp"--><br><br>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="470" height="49" id="AutoNumber1">
    <tr>
      <td width="70%" height="15" background="images/Button.gif"><font size="2">&nbsp;
      <img border="0" src="images/eye.gif">&nbsp; <b>我的關注話題</b>(您可以有<b><%=rs_level2("fav_size")%></b>個關注話題)</font></td>
    </tr>
 <%sql="select * from favorite where name='"&request.cookies("UserID")&"'"
set rs=conn.execute(sql)
if not rs.eof then
i=0
while not rs.eof
i=i+1
sql="select * from Titles where TitleID="&rs("titleid")
set rs2=conn.execute(sql)%>
    <tr>
      <td width="470" height="30">
      <p style="margin-left: 5"><img border="0" src="images/eye.gif">&nbsp;<font size="2">
      <%if rs2 Is Nothing Or rs2.EOF Then
      response.write "<font color=red size=2>這個關注主題文章已被刪除，請將之移除吧！</font>"%>
      ..........<a href="my_favorite.asp?mode=del&id=<%=rs("titleID")%>">移除</a></font>      
      <%else%>      
      <a href="Detail.asp?TitleID=<%=rs2("TitleID")%>&tid=<%=rs2("tid")%>&postname=<%=rs2("Name")%>"><%=rs2("Subject")%></a>
      ..........<a href="my_favorite.asp?mode=del&id=<%=rs2("TitleID")%>">移除</a></font>
      <%end if%>
      </td>
    </tr>
<%rs.movenext
wend
response.write "<font color=red size=2>已經存有<b>"&i&"</b>個關注話題</font>"
else
response.write "<tr><td><center><font color=red size=-1>無關注話題</font></center></td></tr>"
end if%>
  </table>
  </center>
</div>
 </font>
 <!--#include file="foot.asp"-->
</body></html>