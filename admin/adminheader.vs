<!--#include file="adminconn.vs"-->
<% If request.cookies(bbsid)("UserPassed") = "yes" Or Session("Passed")=True Then %>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#CCCCFF" width="90%" height="3" background="file:///C:/Inetpub/wwwroot/bbs/images/Button.gif">
    <tr>
      <td width="568" height="3">�@
      <font size=-1><% Response.Write "<font color=red SIZE=-1><b>"&request.cookies(bbsid)("UserID")&"</font>"&"</b>-�w��z�^��!&nbsp;&nbsp;&nbsp;"%>|
<a href=../edituser.asp?UserID=<%=request.cookies(bbsid)("UserID")%> style="text-decoration: none">
      <font color="#000000">�ӤH���|</font></a> <a href="../user.asp" style="text-decoration: none"><font color="#000000">�׾·|��</font></a> 
      | <a href="../pm.asp"style="text-decoration: none"><font color="#000000">²�T�޲z </font></a>
      |
<a href=../loginout.asp style="text-decoration: none"> <font color="#000000">�n�X�׾� </font></a>
      |<%If Session("Passed")=True Then%><a href="../admin.asp" style="text-decoration: none"><font color="#000000">
�޲z�׾�</font></a><%End If%> 












</td>
    </tr>
  </table>
  </center>
</div>






 <%End If%>

<br>

<p align="center" style="margin-top: 0; margin-bottom: 0"><a href="../index.asp"><img border="0" src="../title.jpg"></a> 
</p>
  <!-- �ˬd�O�_��pm -->
<%
If request.cookies(bbsid)("UserID")<>Empty Then
Set rspm=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM id WHERE user='" & request.cookies(bbsid)("UserID") & "'"
rspm.Open SQL, conn,1,3

If rspm("havepm")=1 and not rspm.EOF Then
	Response.Write "<a href=../pm.asp><img src=""../images/newnew.gif""border=0></a><embed src=""../images/mail.wav""width=0 height=0>"
End If
end if
%>