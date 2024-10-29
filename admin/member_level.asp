<html><!--#include file="conn.asp"-->
<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>新增網頁1</title><link rel="stylesheet" type="text/css" href=../style1.css>
</head>
<%sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else
if request("act")="add_level" then
   Set rs2= Server.CreateObject("ADODB.Recordset") 
   SQL="select * from member_sort "
   rs2.Open SQL, conn,1,3
   rs2.addnew
   rs2("level")=request("level")
   rs2("levelname")=request("levelname")
   rs2.update
elseif request("act")="delete" then
   Set rs2= Server.CreateObject("ADODB.Recordset") 
   SQL="select * from member_sort where id="&request("id")
   rs2.Open SQL, conn,1,3
   rs2.delete   
end if%>
<body>
<div align="left">
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="391" height="91" id="AutoNumber1">
    <tr>
      <td width="391" height="19" background="images/title.gif"><b>
      <font size="3">新增等級</font></b></td>
    </tr>
    <tr>
      <td width="391" height="68">
      <form method="GET" >
        <font size="2">頭銜：</font><input type="text" name="levelname" size="20"><font size="2">&nbsp; 等級代號：</font><input type="text" name="level" size="7">
        <br><br><input type="hidden" name="act" value="add_level" size="20">
        <input type="submit" value="確定" name="B1"><input type="reset" value="重新設定" name="B2">
      </form>
      </td>
    </tr>
  </table>
</div>

<br>
<div align="left">
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="198" height="29" id="AutoNumber2">
    <tr>
      <td width="30" height="14" background="images/title.gif"><b>
      <font size="2">代號</font></b></td>
      <td width="90" height="14" background="images/title.gif"><b>
      <font size="2">頭銜</font></b></td>
      <td width="40" height="14" colspan="2" background="images/title.gif"><b>
      <font size="2">動作</font></b></td>
    </tr>
    <%
    sql="select * from member_sort order by id"
    set rs=conn.execute(sql)
    while not rs.eof
    %>
    <tr>
      <td width="30" height="11"> <font size="2"><%=rs("level")%></font></td>
      <td width="90" height="11"> <font size="2"><%=rs("levelname")%></font></td>
      <td width="40" height="11">
      <form method="post" action="member_level_edit.asp">
        <input type="submit" value="修改" name="B1">
        <input type="hidden" value=<%=rs("id")%>  name="id">
      </form>
        </td>
      <td width="40" height="11">
 <form method="GET">
        <input type="submit" value="刪除" name="B1">
        <input type="hidden" value="delete" name="act">
         <input type="hidden" value=<%=rs("id")%> name="id">
      </form>      </td>
    </tr>
    <%rs.movenext
    wend%>
  </table>
</div>
  <!--#include file="../foot.asp"-->
<%end if%>
</body>
</html>