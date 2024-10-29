<!--#include file="conn.asp"-->
<html>
<head><link rel="stylesheet" type="text/css" href=style1.css></head>
<body topmargin="0" leftmargin="0"><!--#include file="header.asp"-->
<!--#include file="admin/broad.asp"--><br>
<% response.cookies("tid")="通訊錄"%><!--#include file="choice.asp"--><br><br>
<%Set rs33=Server.CreateObject("ADODB.Recordset") 
rs33.open "friendpm",conn,1,3
Select Case request("act")
case "add"
sql="select * from id where user='"&request("name")&"'"
sql1="select * from friendpm where friendname='"&request("name")&"'"&"and user='"&request.cookies("UserID")&"'"
set rs=conn.execute(sql)
set rs1=conn.execute(sql1)
if rs.eof then%>
<Script Language=Vbscript>
MsgBox "無此帳號!",64,"錯誤！！"
location.href = "javascript:history.back()"
</Script>
<%elseif not rs1.eof then%>
<Script Language=Vbscript>
MsgBox "已存在!",64,"錯誤！！"
location.href = "javascript:history.back()"
</Script>
<%else
rs33.addnew
rs33("user")=request.cookies("UserID")
rs33("friendname")=request("name")
rs33.update
rs33.close
end if
case "delete" 
If not rs33.EOF then rs33.movefirst
while not rs33.eof 
if request(rs33("id"))="on" then
rs33.delete

end if
rs33.movenext
wend

end select%>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="384" height="83" id="AutoNumber1">
    <tr>
      <td width="384" height="7" background="images/Button.gif">
      <p align="center"><font size="2">
      <img border="0" src="images/userlist.gif">通訊錄</font></td>
    </tr>
    <tr><form method="GET">
      <td width="384" height="27" bgcolor="#E8F2FF" bordercolor="#C0C0C0" background="images/postbit_lightbg.gif">
      <p style="margin-top: 0; margin-bottom: 0">
      <font size="2">名單</font><input type="text" name="name" size="20">
      <input type="hidden" value="add" name="act"></p>
        <p style="margin-top: 0; margin-bottom: 0"><input type="submit" value="新增" name="B1"><input type="reset" value="重新設定" name="B2"></p> </td>
    </form></tr>
    <tr>
      <td width="384" height="111" bgcolor="#E8F2FF" background="images/postbit_lightbg.gif">
      <form method="GET">
	  <%sql="select * from friendpm where user='"&request.cookies("UserID")&"'"
      set rs12=conn.execute(sql)
      
      if not rs12.eof then
      while not rs12.eof 
	   sql="select * from id where user='"&rs12("friendname")&"'"
	   set rsp=conn.execute(sql)
       response.write "<img src=face/"&rsp("pic")&".gif width=50 height=50><input type=checkbox name="&rs12("id")&" value=""on"">"&"<a href=u2u.asp?touser="&rs12("friendname")&">"&rs12("friendname")&"</a><br>"
	   rs12.movenext
      wend
      end if%>
      <INPUT TYPE="hidden" name="act" value="delete">
	  <INPUT TYPE="submit" value="刪除">
      </form>
      </td>
    </tr>
  </table>
  </center>
</div><br><br><br><!--#include file="foot.asp"-->
</body></html>