<HTML><!--#include file="conn.asp"-->
<HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""><link rel="stylesheet" type="text/css" href=style1.css>
</HEAD>
<BODY>
<%
If request.cookies("admin_pass")<>bbsid Then
Response.Redirect "index.asp"
else
sql="select * from id where ID="&request("id")
set rs=conn.execute(sql)%>
 <FORM METHOD=POST ACTION="../edituseraction.asp" name="form">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="98%" height="1" id="AutoNumber1">
  <tr>
    <td width="100%" colspan="4" height="19" bgcolor="#F0EEDF" bordercolor="#FFFFFF" background="images/bg2.gif">
    <p align="center"><b><font color="#800000">�ӤH���</font></b><font size="2" color=red>�]�������񶵥ءA���i�d�ա^</font></p>   
    </td>
  </tr>
  <tr>
    <td width="15%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">�b��*</font></td>
    <td width="33%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="UserID" VALUE="<%=rs("user")%>" size="20" disabled>
    <INPUT TYPE="hidden" NAME="username" VALUE="<%=rs("user")%>" size="20"></font></td>
    <td width="17%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">�ɮt�վ�</font></td>
    <td width="35%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="timezone" size="20" VALUE="<%=rs("timezone")%>"></font></td>
  </tr>
  <tr>
    <td width="15%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000" >�K�X*</font></td>
    <td width="33%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="Password" NAME="Password1" size="20" VALUE="<%=rs("password")%>"></font></td>
    <td width="17%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">�ЦA��J�@��*</font></td>
    <td width="35%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="Password" NAME="Password2" size="20" VALUE="<%=rs("password")%>"></font></td>
  </tr>
  <tr>
    <td width="15%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000"><a href="mailto:<%=rs("email")%>">E-MAIL*</a></font></td>
    <td width="33%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="email" size="20" VALUE="<%=rs("email")%>"></font></td>
    <td width="17%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <%if rs("homepage")<>Empty then%>
    <font size="2" color="#800000"><a target="_blank" href="<%=rs("homepage")%>">�ӤH����</a></font></td>
    <%else%>
    <font size="2" color="#800000">�ӤH����</font></td>
    <%end if%>
    <td width="35%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="homepage" size="20" VALUE="<%=rs("homepage")%>"></font></td>
  </tr>
  <tr>
    <td width="15%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">MSN Messenger</font></td>
    <td width="33%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="msnm" size="20" VALUE="<%=rs("msnm")%>"></font></td>
    <td width="17%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
	<%if rs("yim")<>Empty then%>
	<font size="2" color="#800000"><a target="_blank" href="http://messenger.yahoo.com/edit/send/?.target=<%=rs("yim")%>&.src=pg">�_�� Messenger</a></font></td>
	<%else%>
    <font size="2" color="#800000">�_�� Messenger</font></td>
	<%end if%>    
    <td width="35%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="yim" size="20" VALUE="<%=rs("yim")%>"></font></td>
  </tr>
  <tr>
    <td width="15%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">�ͤ�*</font></td>
    <td width="33%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="birthday" size="20" VALUE="<%=rs("birthday")%>"></font></td>
    <td width="17%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">�Ӧ�*</font></td>
    <td width="35%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="comefrom" size="20" VALUE="<%=rs("comefrom")%>"></font></td>
  </tr>  
  <tr>
    <td width="15%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000"></font></td>
    <td width="33%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"></font></td>
    <td width="17%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">����</font></td>
    <td width="35%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000"><INPUT TYPE="text" NAME="level" size="20" VALUE="<%=rs("level")%>"></font></td>
  </tr>
  <tr>
    <td width="15%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000" size="2">�ӤH�Y��:</font></td>
    <td width="33%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font color="#800000">  </font></td>
    <td width="17%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">  </font></td>
    <td width="35%" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <font size="2" color="#800000">(�޲z��:1,���D:2,�|��:0)</font></td>
  </tr>
  <tr>
    <td width="100%" colspan="4" height="13" bgcolor="#FFFFFF" bordercolor="#FFFFFF">
    <%if rs("pic") <> 0 then%>
    <img src="../face/<%=rs("pic")%>.gif" name="tus" width="100" height="100">
    <%else%>
    <img src="<%=rs("face_url")%>" name="tus" width="100" height="100">
    <%end if%>
�@<script>function showimage(){document.images.tus.src="../face/"+document.form.pic.options[document.form.pic.selectedIndex].value+".gif";}</script>    
	<select name=pic size=1 onChange="showimage()">
<%for i=1 to 38%>
	<%if rs("pic")=i then%>
	<option value=<%=i%> selected><%=i%></option>
	<%else%>
	<option value=<%=i%>><%=i%></option>
	<%end if%>
<%next%>
	<%if rs("pic")=0 then%>
	<option value=0 selected>�ۭq</option>
	<%else%>
	<option value=0>�ۭq</option>
	<%end if%>
	<select>
      <font size="2">�Y�ۭq�п�J������m:</font><input type="text" name="face_url" size="35" value="<%=rs("face_url")%>"> 
      </select>
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="4" height="19" bgcolor="#F0EEDF" bordercolor="#FFFFFF" background="images/bg2.gif">
    <p align="center"><font color="#800000"><b>�ӤH�T��</b></font></p> 
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="4" height="19" bgcolor="#F0EEDF" bordercolor="#FFFFFF">
    <p align="center"><font size="2">ñ�W��:</font><textarea rows="9" name="memo1" cols="50"><%=rs("memo1")%></textarea></p>   
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="4" height="19" bgcolor="#F0EEDF" bordercolor="#FFFFFF">
    <p align="center"><INPUT TYPE="hidden" name="act" value="admin"><INPUT TYPE="submit" VALUE="�T�w"><INPUT TYPE="reset" VALUE="���s"></p>   
    </td>
  </tr>

  </table>
  </center>
</div>
</FORM>
<%end if%>
</BODY>
</HTML>