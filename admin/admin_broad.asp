<html><!--#include file="conn.asp"-->
<head>
<%Set rs = Server.CreateObject("ADODB.Recordset")
   UserID=request.cookies("UserID")
If request("act")="add" Then
   SQL = "Select * From broad "
   rs.Open SQL, Conn,1,3
   rs.AddNew
   rs("UserID")=UserID
   rs("Subject")=Request("Subject")
   rs("text")=Request("text")
   rs("date")=now()
   rs.Update
   rs.close
ElseIf request("act")="edit" Then
  SQL="Select * From broad where ID=" & request("ID")
     rs.Open SQL, Conn,1,3
	 rs("Subject")=Request("subject")
     rs("UserID")=UserID	 
	 rs("text")=Request("text")
	 rs("date")=Now()
	 rs.Update
	 rs.close
elseif request("act")="del" Then
  SQL="delete * from broad where ID=" & request("ID")
  conn.Execute SQL
End if	
		 SQL="SELECT * From broad"
Set rs=conn.execute(SQL,,1)%>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>最新消息</title>
<link rel="stylesheet" type="text/css" href=../style1.css>
</head>
<body leftmargin="0" topmargin="0">
<%sql="select * from config_setting where ID=1"
     Set rsidid=conn.execute(sql)
       bbsid=Cstr(rsidid("bbsid"))
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else%>
　<p>　</p>
<form method="GET" style="margin-top: 0; margin-bottom: 0">
<div align="center">
    <center>
    <table border="1" cellspacing="0" style="border-collapse: collapse; font-size: 10 pt" bordercolor="#C0C0C0" id="AutoNumber1" bgcolor="#E1E0BB">
      <tr>
        <td colspan="3" align="center" style="font-size: 11 pt; color: #FFFFFF" bgcolor="#E1E0BB" background="../images/bg2.gif">
        <font size="3" color="#000000">新增最新消息</font></td>
      </tr>
      <tr>
        <td style="color: #FFFFFF" bgcolor="#E8F2FF" align="center">
        <font color="#000000">文字訊息</font></td>
        <td style="color: #FFFFFF" bgcolor="#E8F2FF" align="center">
        <font color="#000000">內容</font></td>
        <td style="color: #FFFFFF" bgcolor="#E8F2FF" align="center">
        <font color="#000000">動作</font></td>
      </tr>
      <tr>
        <td bgcolor="#E8F2FF"><input type="text" name="subject" size="25"></td>
        <td bgcolor="#E8F2FF"><input type="text" name="text" size="30"></td>
        <td bgcolor="#E8F2FF"><input type="submit" value="新增" name="b1">
        <input type=hidden name=act value=add>
     </td>
      </tr>
    </table>
    </center>
  </div>
</form>
<br>
<div align="center">
    <center>
    <table border="1" cellspacing="0" style="border-collapse: collapse; font-size: 10 pt" bordercolor="#E8F2FF" id="AutoNumber1" bgcolor="#E8F2FF">
      <tr>
        <td width="100%" colspan="3" align="center" style="font-size: 11 pt; color: #FFFFFF" bgcolor="#E8F2FF" bordercolor="#E1E0BB" background="../images/bg2.gif">
        <font size="3" color="#000000">修改/刪除最新消息</font></td>
      </tr>
      <tr>
        <td style="color: #FFFFFF" bgcolor="#E8F2FF" align="center" bordercolor="#C0C0C0">
        <font color="#000000">文字訊息</font></td>
        <td style="color: #FFFFFF" bgcolor="#E8F2FF" align="center" bordercolor="#C0C0C0">
        <font color="#000000">內容</font></td>
        <td style="color: #FFFFFF" bgcolor="#E8F2FF" align="center" bordercolor="#E1E0BB">
        <font color="#000000">動作</font></td>
      </tr>
      <% Do While (not rs.eof)%>
      <form method="GET" style="margin-top: 0; margin-bottom: 0">
      <tr>
        <td bgcolor="#E8F2FF" bordercolor="#E1E0BB"><input type="text" name="subject" size="25" value="<%=rs("Subject")%>"></td>
        <td bgcolor="#E8F2FF" bordercolor="#E1E0BB"><input type="text" name="text" size="30" value="<%=rs("text")%>"></td>
        <td bgcolor="#E8F2FF" bordercolor="#E1E0BB"><input type="submit" value="修改" name="b1">
        <input type=hidden name=act value=edit>
        <input type=hidden name=id value=<%=rs("ID")%>><input type="button" value="刪除" name="b2" onclick="window.location='admin_broad.asp?act=del&id=<%=rs("ID")%>'"></td>
      </tr>
      </form>
      <% rs.movenext : loop : rs.close : set rs=nothing%>
    </table>
    </center>
  </div>
</form>
  <%End If%>
<!--#include file="../foot.asp"-->
</body>
</html>