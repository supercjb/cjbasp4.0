<!--#include file="conn.asp"-->
<html><head>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")%> </TITLE>
<link rel="stylesheet" type="text/css" href=style1.css></head>
<body leftmargin="0" topmargin="0"><!--#include file="header.asp"--><!--#include file="admin/broad.asp"--><br>
<% response.cookies("tid")="�j�M�峹"%>
<!--#include file="choice.asp"--><br><br><center>
<table width="50%" style="border-collapse: collapse" bordercolor="#C0C0C0" cellpadding="0" cellspacing="0" border="1">
<tr><td background="images/Button.gif"><font size="2">&nbsp;</font><font size="2" color="#000080">�׾·j�M</font></tr></td>
<FORM METHOD=POST ACTION="search_action.asp"><tr><td>
<font size="2">�Q�װ�&nbsp;&nbsp;&nbsp; �G</font><select name="sort">
<%Set rsjump= Server.CreateObject("ADODB.Recordset") 
rsjump.Open "tid order by tid",conn
%><option Value="aa">�Ҧ����Q�װ�</option><%
while not rsjump.eof%><option VALUE=<%=rsjump("tid")%>><%=rsjump("tidname")%></option>
<%rsjump.movenext
wend
rsjump.close%>
</select></tr></td>
	<tr>
		<td>
<font size="2">�j�M�d��G</font><select name="range" size="1">
<option selected value="0">���D</option>
<option value="1">�@��</option>
<option value="2">���e</option>
</select></tr>
	</tr>
<tr><td><font size="2">�j�M���e�G</font><INPUT TYPE="text" NAME="tit" size="35"></tr></td>
<tr><td>�@</tr></td>
<tr>
	<td><font size="2">�j�M���G�ƧǨ�</font><select size="1" name="order">
	<option value="0" selected>�̫��s</option>
	<option value="1">�o����</option>
		<option value="2">�I�\��</option>
			<option value="3">�^�м�</option>
	</select>
	
	</tr>
</tr>
<tr><td><font size="2">����r�j�M�G</font><INPUT TYPE="checkbox" NAME="key" value="ON"></tr></td>
<tr><td>
  <p align="center"><INPUT TYPE="submit" value="�}�l�j�M"></tr></td>
</FORM></table>
  <p align="right">
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;}"><option selected>����׾¦�...</option><%
Set rsjump= Server.CreateObject("ADODB.Recordset") 
rsjump.Open "tid order by tid",conn
while not rsjump.eof
%><option value=forum.asp?tid=<%=rsjump("tid")%>&Page=1><%=rsjump("tidname")%></option><%
rsjump.movenext
wend
rsjump.close%></select></p>
<!--#include file="foot.asp"-->
</body></html>