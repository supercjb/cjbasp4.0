<!--#include file="conn.asp"-->
<html><head>
<meta http-equiv="Content-Language" content="zh-tw">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><%=rsidid("bbsname")%></title>
<link rel="stylesheet" type="text/css" href=style1.css></head>
<body leftmargin="0" topmargin="0">
<!--#include file="header.asp"--><!--#include file="admin/broad.asp"-->
<%response.cookies("tid")="�|���n��"%><!--#include file="choice.asp"--><br>
<div align="center">
  <center><b><font color=red><%=request("msg")%></font></b>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="316" height="68" id="AutoNumber1">
    <tr>
      <td width="316" height="14" background="images/Button.gif">&nbsp;<font size="2"><b><img border="0" src="images/login.gif" width="12" height="13">�|���n�J</b></font></td>
    </tr>
    <tr>
      <td width="316" height="50" background="images/postbit_lightbg.gif"><FORM METHOD=POST ACTION="check.asp">
      <% 
	  ExpireDate=DateAdd("d",3650,Date)
   response.cookies("TEST_COOKIE").Expires=FormatDateTime(ExpireDate)
	  response.cookies("TEST_COOKIE") = "testing" '�n���e����!�ˬdCOOKIE%>
<font size="2">&nbsp;�Τ�b���G<INPUT TYPE="text" NAME="UserID" size="12">&nbsp;&nbsp;&nbsp;&nbsp;</font><p>
<font size="2">&nbsp; �b���K�X�G<INPUT TYPE="password" Name="Password" size="12">&nbsp;</font></p>
		<p>
<font size="2">&nbsp; �b���O���G<select size="1" name="cookie_save_day">
<option selected value="365">�û�</option>
<option value="30">�@�Ӥ�</option>
<option value="7">�@�P</option>
<option value="1">�@��</option>
<option value="0">���ΰO��</option>
</select></font></p>
<p>
<INPUT TYPE="hidden" Name="act" size="12" value="head_login">
<font size="2"><INPUT TYPE="submit" VALUE="�n�J"><INPUT TYPE="reset" VALUE="����">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</font>
<%if request("mod")="reg" then%>
<input type="hidden" name="url" value="index.asp">
<%else%>
<input type="hidden" name="url" value="<%=request.servervariables("HTTP_REFERER")%>">
<%end if%></p>
</FORM></td>
    </tr>
  </table>
  </center>
</div>
  <!--#include file="foot.asp"-->
</body></html>