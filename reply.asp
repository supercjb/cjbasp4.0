<!--#include file="conn.asp"--><!--#include file="fun.asp"--><!--#include file="fun2.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD>
<TITLE><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
 response.cookies("upload")="0"
 response.cookies("filename")="" 
'此處刪除檔案上傳的cookies%> </TITLE>
<link rel="stylesheet" type="text/css" href=style1.css>
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></HEAD>
<BODY leftmargin="0" topmargin="0">
<!--#include file="header.asp"--><!--#include file="admin/broad.asp"--><!--#include file="choice.asp"--><BR>
<center>
<form method="POST" action="DetailNew.asp" name="InputForm">
<INPUT TYPE="hidden" NAME=TitleID Value=<%=Request("TitleID")%>>
<INPUT TYPE="hidden" NAME=tid Value=<%=Request("tid")%>>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="92%" height="1" id="AutoNumber2">
	<tr> 
		<td width="100%" height="15" align="center" bgcolor="#EFECCB" bordercolor="#FFFFFF" background="images/windowbar_lightbg.gif" colspan="2"> 
		<p align="left"><font size="3">&nbsp; <b>新增回覆</b></font></td>
	</tr>
	<tr> 
		<td width="100%" height="25" bgcolor="#EBEBEB" colspan="2" background="images/postbit_lightbg.gif"><font size="2">&nbsp;&nbsp;&nbsp; 帳號:</font>
		<input type="text" name="Name1" size="28" Value="<%=request.cookies("UserID")%>" disabled>
		<input type="hidden" name="Name" size="28" Value="<%=request.cookies("UserID")%>"></td>
	</tr>
	<tr>
		<td width="100%" height="32" bgcolor="#EBEBEB" colspan="2" background="images/postbit_lightbg.gif">&nbsp;&nbsp;
		<font size="2">字型顏色:</font>
<select size="1" name="Color1" onchange="document.InputForm.Words.style.color=document.InputForm.Color1.value">                                                      
<option value="#000000" style="color: #000000" selected>就是黑</option>                                                      
<option value="#6894F0" style="color: #6894F0">天空藍</option>                                                      
<option value="#000090" style="color: #000090">深色藍</option>                                                      
<option value="#B88810" style="color: #B88810">卡其色</option>                                                      
<option value="#FF8050" style="color: #FF8050">橘子色</option>                                                      
<option value="#880088" style="color: #880088">暗紫色</option>                                                      
<option value="#FF1498" style="color: #FF1498">淡粉紅</option>                                                      
<option value="#B02420" style="color: #B02420">咖啡色</option>                                                      
<option value="#DD0000" style="color: #DD0000">就是紅</option>                                                      
<option value="#0000FF" style="color: #0000FF">就是藍</option>                                                       
<option value="#008000" style="color: #008000">就是綠</option>                                                      
<option value="#006400" style="color: #006400">深綠色</option>                                                      
<option value="#FFB312" style="color: #FFB312">蛋黃色</option>                                                      
<option value="#7888A0" style="color: #7888A0">馬路灰</option>                                                      
<option value="#709028" style="color: #709028">抹茶綠</option>                                                      
<option value="#FCC22C" style="color: #FCC22C">芒果黃</option>                                                      
</select>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="2" color="#B88810"> 
<!--#include file="inc_ubb_bar.asp"-->

(UBB
<%Set rsset = Server.CreateObject("ADODB.Recordset")
SQL = "Select * From config_setting where ID=1"
   rsset.Open SQL, Conn,1,3
If rsset("UBBcode") =False Then %>不<%End If%>開啟，請看</font><font size="2"><a target="_blank" href="help_ubb.htm"><font color="#B88810">UBB幫助</font></a></font><font size="2" color="#B88810">)&nbsp;&nbsp; 
HTML<%If rsset("htmlencode") = False Then%>不<%End If:rsset.close%>開啟</font>
		</td>
	<tr>		            
		<td width="100%" height="1" bgcolor="#EBEBEB" bordercolor="#F5F7F4" colspan="2" background="images/postbit_lightbg.gif">
            <p align="center" style="margin-top: 0; margin-bottom: 0">
            <font size="2" color="#800000">插入表情圖案</font></p></td>
	</tr>
	<tr>          
		<td width="35%" height="1" bgcolor="#EBEBEB" bordercolor="#F5F7F4" align="right" background="images/postbit_lightbg.gif">
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
		<td width="65%" height="1" bgcolor="#EBEBEB" bordercolor="#F5F7F4" align="left" background="images/postbit_lightbg.gif">
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
		<td width="100%" height="1" align="center" rowspan="2" bgcolor="#EBEBEB" colspan="2" background="images/postbit_lightbg.gif"> 
		<p align="left"><font size="2">內容:</font>
<%'*******************************引用功能*************************
	if request("mode")="titlequote" then%>
	<textarea rows="9" name="Words" cols="98" ><%sql="select * from Titles where TitleID="&request("TitleID")
	set titlequote=conn.execute(sql)%>[quote]引用<%=titlequote("Name")%>的文章：<%=replace(replace(titlequote("Words"),"[quote]",""),"[/quote]","")%>[/quote]</textarea>
	<%elseif request("mode")="replyquote" then
	sql="select * from Details where DetailID="&request("DetailID") 
	set replyquote=conn.execute(sql)%>
	<textarea rows="9" name="Words" cols="98" >[quote]引用<%=replyquote("Name")%>的文章：<%=replace(replace(replyquote("Words"),"[quote]",""),"[/quote]","")%>[/quote]</textarea>
	<%else%>
	<textarea rows="9" name="Words" cols="98" ></textarea>
	<%end if%></p>
	<p align="left" style="margin-top: 0; margin-bottom: 0">　
	<%
	if rsidid("upload")=true then%>
	<IFRAME name=ad src="upload0.asp" frameBorder=0 width="70%" scrolling=no height=80></IFRAME></p>
	<%end if%>
	<a href="javascript:checklength(document.InputForm.Words)">
	<font size="2">查看字數</font></a><font size="2">(字數限制<%=rstitle("word_len")%>字)</font>
	<p align="center">
	<input type="submit" value="確定回覆文章嘍" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
	<input type="reset" value="重新發表" name="B2"></p>
		</td>
	</tr>
</table>
<p align="center" style="margin-top: 0; margin-bottom: 0">　</p>
<p align="center" style="margin-top: 0; margin-bottom: 0"><a href="javascript:history.go(-1)"><font size="2">返回</font></a></p>　	
<table border="1" style="border-collapse: collapse" bordercolor="#808080" cellpadding="0" cellspacing="0" width="90%"><tr>
	<tr>
		<td bgcolor="red" bordercolor="#808080" STYLE="WORD-BREAK: break-all" background="images/postbit_lightbg.gif"><font size="-1">
<%if request.cookies("UserID")=Empty then%>
		<script language="javascript">
	           alert("請先註冊才可回覆文章唷！")
	           location.href="javascript:history.go(-1)"
        </script>
<%end if%>
<%sql="select * from Titles where TitleID="&Request("TitleID")
    set rs_ti=conn.execute(sql)
if rs_ti("reply_lock")=1 Or request.cookies("UserID")=Empty then%>
		<script language="javascript">
	           alert("此主題禁止回覆。")
	           location.href="javascript:history.go(-1)"
		</script>
<%else
     response.write "<b><font color=red>標題："&rs_ti("Subject")&"<br>"
     response.write "作者："&rs_ti("Name")&"</font></b><br>"
    If rstitle("UBBcode")=True Then
	  If rstitle("htmlencode")=True Then
	    subjectcolor UBBCode(rs_ti("Words")),rs_ti
      Else
	    subjectcolor  UBBCode(HTMLEncode(rs_ti("Words"))),rs_ti 
	  End If       
    Else 	 
	  If  rstitle("htmlencode")=True Then 
	    subjectcolor rs_ti("Words"),rs_ti
	  Else
	    subjectcolor  HTMLEncode(rs_ti("Words")),rs_ti
	  End If
	End If
end if%></font>　</td></tr></table></center>  
</form><!--#include file="foot.asp"-->
</BODY></HTML>