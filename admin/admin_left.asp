<html><!--#include file="conn.asp"-->
<head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>new document</title><link rel="stylesheet" type="text/css" href=../style1.css>
</head>
<body>
<%sql="select * from config_setting where ID=1"
Set rsidid=conn.execute(sql)
bbsid=Cstr(rsidid("bbsid"))
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else%>
<div align="left">
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="128" height="378" id="AutoNumber1">
    <tr>
      <td width="128" height="17" background="images/title.gif"><b>
      <font size="2">CJB-ASP�޲z���</font></b></td>
    </tr>
    <tr>
      <td width="128" height="27" background="images/title.gif"><b>
      <font size="2">�t�κ޲z</font></b></td>
    </tr>
    <tr>
      <td width="60" height="30">
      <p style="margin-top: 0; margin-bottom: 0"><font size="2">
      <a target="main" href="admin_broad.asp">���i�޲z</a></font></p>
      <p style="margin-top: 3; margin-bottom: 0"><font size="2">
      <a target="main" href="admin.asp">�ֳt�޲z</a></font><p style="margin-top: 3; margin-bottom: 0">
      <font size="2"><a target="main" href="admin_pm.asp">pm�q��</a></font><p style="margin-top: 3; margin-bottom: 0">
      <font size="2"><a target="main" href="admin_web.asp">�ͱ��s��</a></font></td>
    </tr>
    <tr>
      <td width="128" height="27" background="images/title.gif"><b>
      <font size="2">�׾º޲z</font></b></td>
    </tr>
    <tr>
      <td width="128" height="28"><font size="2">
      <a target="main" href="admin_sort.asp">�׾¥D�����޲z</a></font></td>
    </tr>
    <tr>
      <td width="128" height="22" background="images/title.gif"><b>
      <font size="2">�|���޲z</font></b></td>
    </tr>
    <tr>
      <td width="128" height="28">
      <p style="margin-top: 2; margin-bottom: 0"><font size="2">
      <a target="main" href="member_level.asp">�|�����ų]�w</a></font></p>
      <p style="margin-top: 3; margin-bottom: 0"><font size="2">
      <a target="main" href="admin_user.asp">�|���޲z</a></font></td>
    </tr>
    <tr>
      <td width="128" height="123">�@</td>
    </tr>
  </table>
</div>
<%end if%>
</body>
</html>