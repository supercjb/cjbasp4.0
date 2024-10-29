<%
Set rsb=Server.CreateObject("ADODB.Recordset")
rsb.Open "Select * From broad ",conn,1,1

SQL = "Select * From id Where user='"&request.cookies("UserID")&"'"
set rs11=conn.execute(SQL)
sql="select * from config_setting"
set rstitle=conn.execute(sql)
%>

<CENTER>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#D1D1D1" width="543" height="2" id="AutoNumber1">
    <tr>
      <td width="552" height="2" bordercolor="#F0EFDD" 
      bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" background="../images/postbit_lightbg.gif">
      <p align="center"><font size="2" color="#008080"><img src=images/icon_announce.gif border=0>
<applet code="CreditRoll.class" width=500 height=30>
<param name=BGCOLOR value="ffffff">
<param name=TEXTCOLOR value="#660099">
<param name=FADEZONE value="50">
<%i=1%>
 <% Do While (not rsb.eof)%>
<param name=TEXT<%=i%>  value="<%=rsb("Subject")%>">
<%i=i+1%>
<param name=URL value="broad.asp">
<param name=REPEAT value="yes">
<param name=SPEED value="90">
<param name=VSPACE value="20">
<param name=FONTSIZE value="12">
<% rsb.movenext : loop : rsb.close : set rsb=nothing%>
</applet>
</font></td>
    </tr>
  </table>
  </CENTER>





