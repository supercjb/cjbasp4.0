<!--#include file="conn.asp"--><!--#include file="fun.asp"-->
<HTML><HEAD>
<meta http-equiv="Content-Language" content="zh-tw">
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")%></TITLE>
<meta http-equiv="refresh" content="1200; url=index.asp">  
<META NAME="Description" CONTENT="">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href=style1.css></HEAD>
<BODY leftmargin="0" topmargin="0">
<!-- 登錄 --><!--#include file="header.asp"--><!--#include file="admin/broad.asp"-->
<%countobj="guest"%><!--#include file="count.asp"-->
<center><br><p align="left">
<% response.cookies("tid")="流量統計"%><!--#include file="choice.asp"--></p><br><br>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="50%" id="AutoNumber1">
  <tr>
	<td width="100%" align="center" background="images/Button.gif" colspan="2">
	<b><font size="2"><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
%>訪客流量統計</font></b></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">總訪客數：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=totalguest%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">今日訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=dayguest%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">昨日訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=dayguestu%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">本月訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=monthguest%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">上月訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=monthguestu%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">本季訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=seasonguest%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">上季訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=seasonguestu%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">本年訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=yearguest%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">去年訪客：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=yearguestu%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">您的ＩＰ：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=ipguest%>　</font></td></tr>
  <tr>
	<td width="50%" align="center" bgcolor="#F7F5EE" bordercolor="#000000" background="images/postbit_lightbg.gif">
	<font size="2" color="#800000">您來訪次數：</font></td>
      <td width="50%" align="center" bgcolor="#EFF4FE" bordercolor="#000000" background="images/postbit_lightbg.gif">
		<font size="2"><%=youguest%>　</font></td></tr>
</table>
<p>  </p><p>  </p><p>  </p><p>  </p><p>  </p>
</center><!--#include file="foot.asp"-->
</BODY></HTML>