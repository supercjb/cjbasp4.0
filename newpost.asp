<html>
<!--#include file="conn.asp"--><!--#include file="fun2.asp"-->
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
 response.cookies("upload")="0"
 response.cookies("filename")="" 
'���B�R���ɮפW�Ǫ�cookies%></title>
<link rel="stylesheet" type="text/css" href=style1.css></head>
<body leftmargin="0" topmargin="0">
<!--#include file="header.asp"--><!--#include file="admin/broad.asp"--><!--#include file="choice.asp"-->
<%if request.cookies("UserID")=Empty then%>
		<script language="javascript">
	           alert("�Х����U�~�i�o��峹��I")
	           location.href="javascript:history.go(-1)"
        </script>
<%end if%>
<center>
  <form method="POST" action="TitleNew.asp" name="InputForm">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="92%" height="1" id="AutoNumber2">
          <tr> 
            <td width="100%" height="10" align="center" bgcolor="#999999" bordercolor="#999999" background="images/Button.gif" colspan="2"> 
              <p align="left"><b><font size="3">�s�W�D�D</font></b></td>
          </tr>
          <tr> 
            <td width="100%" height="26" bgcolor="#E1E1E1" colspan="2" background="images/postbit_lightbg.gif"><font size="2">�b��:</font>
              <input type="text" name="Name" size="28" Value="<%=request.cookies("UserID")%>" disabled>
			  <INPUT TYPE="hidden" name="Name2" size="28" Value="<%=request.cookies("UserID")%>" >
			  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <font size="2">�D�D:</font> <font size="2">
            <select size="1" name="subc">
      			<option value="[���n]">[���n]</option>
				<option value="[����]">[����]</option>
				<option value="[���]">[���]</option>
				<option value="[�q��]">[�q��]</option>
				<option value="[����]">[����]</option>
				<option value="[����]">[����]</option>
                <option value="[�޳N]">[�޳N]</option>
				<option value="[���i]">[���i]</option>
				<option value="[�K��]">[�K��]</option>
				<option value="[����]" selected>[����]</option>
				<option value="[����]">[����]</option>
				<option value="[���n]">[���n]</option>
				<option value="[�Q��]">[�Q��]</option>
				<option value="[�U��]">[�U��]</option>
				<option value="[��K]">[��K]</option>
				<option value="[�߱�]">[�߱�]</option>
				<option value="[�v��]">[�v��]</option>
	            <option value="[�q�v]">[�q�v]</option>
	            <option value="[MTV]">[MTV]</option>
				<option value="[����]">[����]</option>
				<option value="[�F��]">[�F��]</option>
				<option value="[�s�D]">[�s�D]</option>
            </select></font><input type="text" name="Subject" size="28" ><br>
            <p style="margin-top: 0"><font size="2">�߱��ϥ�</font>
            <input type="radio" value="00" name="R1">
            <img border="0" src="images/post/00.gif">
            <input type="radio" name="R1" value="01">
            <img border="0" src="images/post/01.gif">
            <input type="radio" checked name="R1" value="02">
            <img border="0" src="images/post/02.gif">
            <input type="radio" name="R1" value="03">
            <img border="0" src="images/post/03.gif">
            <input type="radio" name="R1" value="04">
            <img border="0" src="images/post/04.gif">
            <input type="radio" name="R1" value="05">
            <img border="0" src="images/post/05.gif">
            <input type="radio" name="R1" value="06">
            <img border="0" src="images/post/06.gif">
            <input type="radio" name="R1" value="07">
            <img border="0" src="images/post/07.gif">
            <input type="radio" name="R1" value="08">
            <img border="0" src="images/post/08.gif">
            <input type="radio" name="R1" value="09">
            <img border="0" src="images/post/09.gif">
            <input type="radio" name="R1" value="10">
            <img border="0" src="images/post/10.gif">
            <input type="radio" name="R1" value="11">
            <img border="0" src="images/post/11.gif">
            <input type="radio" name="R1" value="12">
            <img border="0" src="images/post/12.gif"></p></td>
          </tr>
          <tr>
            <td width="100%" height="28" bgcolor="#F5F7F4" colspan="2" bordercolor="#F5F7F4" background="images/postbit_lightbg.gif">
<!--#include file="inc_ubb_bar.asp"-->
<font size="2">�ֳt����r���C��:</font> <select size="1" name="Color1" onchange="document.InputForm.Words.style.color=document.InputForm.Color1.value">                                                      
<option value="#000000" style="color: #000000" selected>�N�O��</option>                                                      
<option value="#6894F0" style="color: #6894F0">�Ѫ���</option>                                                      
<option value="#000090" style="color: #000090">�`����</option>                                                      
<option value="#B88810" style="color: #B88810">�d���</option>                                                      
<option value="#FF8050" style="color: #FF8050">��l��</option>                                                      
<option value="#880088" style="color: #880088">�t����</option>                                                      
<option value="#FF1498" style="color: #FF1498">�H����</option>                                                      
<option value="#B02420" style="color: #B02420">�@�ئ�</option>                                                      
<option value="#DD0000" style="color: #DD0000">�N�O��</option>                                                      
<option value="#0000FF" style="color: #0000FF">�N�O��</option>                                                       
<option value="#008000" style="color: #008000">�N�O��</option>                                                      
<option value="#006400" style="color: #006400">�`���</option>                                                      
<option value="#FFB312" style="color: #FFB312">�J����</option>                                                      
<option value="#7888A0" style="color: #7888A0">������</option>                                                      
<option value="#709028" style="color: #709028">�ٯ���</option>                                                      
<option value="#FCC22C" style="color: #FCC22C">�~�G��</option>                                                      
</select>

<font size="2" color="#B88810"> (UBB
<%Set rsset = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From config_setting where ID=1"
   rsset.Open SQL, Conn,1,3
If rsset("UBBcode") =False Then %>��<%End If%>�}�ҡA�Ь�</font><font size="2"><a target="_blank" href="help_ubb.htm"><font color="#B88810">UBB���U</font></a></font><font size="2" color="#B88810">)&nbsp;&nbsp; 
            HTML<%If rsset("htmlencode") = False Then%>��<%End If:rsset.close%>�}��</font>            
             </td>
          </tr>         
          <tr>
            <td width="100%" height="1" bgcolor="#F5F7F4" bordercolor="#F5F7F4" colspan="2" background="images/postbit_lightbg.gif">
            <p align="center" style="margin-top: 0; margin-bottom: 0">
            <font size="2" color="#800000">���J���Ϯ�</font></p></td>
		  </tr>
		  <tr>          
            <td width="35%" height="1" bgcolor="#F5F7F4" bordercolor="#F5F7F4" align="right" background="images/postbit_lightbg.gif">
            <img border="0" src="images/mod/aa01.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa01.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa02.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa02.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa03.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa03.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa04.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa04.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa05.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa05.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa06.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa06.gif[/img]" style=cursor:hand>
            <br>
            <img border="0" src="images/mod/aa07.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa07.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa08.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa08.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa09.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa09.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/aa10.gif" onclick=document.InputForm.Words.value+="[img]images/mod/aa10.gif[/img]" style=cursor:hand>
            </td>
            <td width="65%" height="1" bgcolor="#F5F7F4" bordercolor="#F5F7F4" align="left" background="images/postbit_lightbg.gif">
            <img border="0" src="images/mod/1.gif" onclick=document.InputForm.Words.value+="[img]images/mod/1.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/2.gif" onclick=document.InputForm.Words.value+="[img]images/mod/2.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/3.gif" onclick=document.InputForm.Words.value+="[img]images/mod/3.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/4.gif" onclick=document.InputForm.Words.value+="[img]images/mod/4.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/5.gif" onclick=document.InputForm.Words.value+="[img]images/mod/5.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/6.gif" onclick=document.InputForm.Words.value+="[img]images/mod/6.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/7.gif" onclick=document.InputForm.Words.value+="[img]images/mod/7.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/8.gif" onclick=document.InputForm.Words.value+="[img]images/mod/8.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/9.gif" onclick=document.InputForm.Words.value+="[img]images/mod/9.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/10.gif" onclick=document.InputForm.Words.value+="[img]images/mod/10.gif[/img]" style=cursor:hand>
			<br>
            <img border="0" src="images/mod/11.gif" onclick=document.InputForm.Words.value+="[img]images/mod/11.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/12.gif" onclick=document.InputForm.Words.value+="[img]images/mod/12.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/13.gif" onclick=document.InputForm.Words.value+="[img]images/mod/13.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/14.gif" onclick=document.InputForm.Words.value+="[img]images/mod/14.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/15.gif" onclick=document.InputForm.Words.value+="[img]images/mod/15.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/16.gif" onclick=document.InputForm.Words.value+="[img]images/mod/16.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/17.gif" onclick=document.InputForm.Words.value+="[img]images/mod/17.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/18.gif" onclick=document.InputForm.Words.value+="[img]images/mod/18.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/19.gif" onclick=document.InputForm.Words.value+="[img]images/mod/19.gif[/img]" style=cursor:hand>
            <img border="0" src="images/mod/20.gif" onclick=document.InputForm.Words.value+="[img]images/mod/20.gif[/img]" style=cursor:hand>
            </td></tr>
		<tr>                        
            <td width="100%" height="1" bgcolor="#F5F7F4" rowspan="2" bordercolor="#F5F7F4" colspan="2" background="images/postbit_lightbg.gif">
            <p align="left" style="margin-top: 0; margin-bottom: 0"><font size="2">���e:</font></p>
            <p align="left" style="margin-top: 0; margin-bottom: 0"><textarea rows="9" name="Words" cols="98" wrap="virtual"></textarea></p>
            <p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2" color="#800000"> 
            <a href="javascript:checklength(document.InputForm.Words)">�d�ݦr��</a>(�r�ƭ���<%=rstitle("word_len")%>�r)</font></p>
            <p align="left" style="margin-top: 0; margin-bottom: 0">�@
  <%
	if rsidid("upload")=true then%>
<IFRAME name=ad src="upload0.asp" frameBorder=0 width="70%" scrolling=no height=80></IFRAME></p>
  <%end if%>
  <%if rsidid("en_save")=true then%>
	<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2">�峹�O�s����</font>
            <select size="1" name="save_date">
            <option value=1>1��</option>
            <option value=3>3��</option>
            <option value=10>10��</option>
            <option value=30 selected>30��</option>
            <option value=60>60��</option>
            <option value=90>90��</option>
            <option value=182>�b�~</option>
            <option value=365>�@�~</option>            
            <option value="���]�w">���]�w</option>            
            </select></p>
  <%end if%>
<p align="center" style="margin-top: 0; margin-bottom: 0"><font size="2">���^�ХH²�T�q��</font><input type="checkbox" name="reply_pm" value="ON"></p>
<p align="center">
	<input type="submit" value="�T�w�F,�ڭn�o��峹" name="B1"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font> 
	<input type="reset" value="�ڭn���s�o��峹@" name="B2">
	<input type="hidden" name="tid" value="<%=Request("tid")%>"></p>
		</td>
	</tr>
</table>
<p align="center"><a href="forum.asp?tid=<%=Request("tid")%>"><font size="2" color="#800000">��^</font></a></p>
        <!--#include file="foot.asp"-->
</form>
</body></html>