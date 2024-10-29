<!--#include file="conn.asp"-->
<% function sum_anything(sql,rs)
	Set rs = conn.Execute(sql)
		i=0
	while not rs.eof
		i=i+1
		rs.movenext
	wend
	Response.Write "<font color=red><b>"& i &"</b></font>"
  end function%>

<HTML>
<HEAD>
<TITLE><%=rsidid("bbsname")%> </TITLE>



<meta http-equiv="refresh" content="1200; url=index.asp">  
<META NAME="Description" CONTENT="">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" type="text/css" href=style1.css>
</HEAD>
<BODY leftmargin="0" topmargin="0">

  <!-- 登錄 -->

<!--#include file="header.asp"-->

<!--#include file="admin/broad.asp"-->
<%if rstitle("bbs_close")=true then'假使論壇關閉
response.write "<br><center><b>抱歉!論壇因為以下原因暫時關閉!請勿發表文章或修改資料,以免造成遺失!"
response.write "<br><font color=red>"&rstitle("close_why")&"</font></b></center>"
elseif rstitle("bbs_close")=false then'假使沒關閉
%>

  <CENTER>
    <FONT COLOR=RED size=5><b><%=Request("msg")%></b></FONT>
  </CENTER>
<p align="right" style="margin-top: 0; margin-bottom: 0">
         <font size="2">歡迎新進成員 
</font><%
SQL= "Select * From config_setting where ID=1 "
Set registrs=conn.Execute(SQL)
%>
<font color=red size=-1>"<%=registrs("newregist")%>"&nbsp;</font><font size=-1>加<a name="top">入</a></font> 
<td width="9">

   

<font size=-1> <br>
  
    　<% 
StrA = "select * from id order by ID"
StrB = "select * from Titles order by TitleID"
StrC = "select * from Details order by DetailID"

Response.Write "本論壇共計有" 
sum_anything StrA,rss
Response.Write "位註冊會員,並有"
sum_anything StrB,rss1
Response.Write "篇主題以及"
sum_anything StrC,rss2
Response.Write "篇回覆文章。"
    %>
   </font></p><center>
   <!--#include file="top_new.asp"-->
<table border="1" cellpadding="0" cellspacing="0" style="word-break:break-all; border-collapse:collapse" bordercolor="#C4E0FF" width="98%" height="68" id="AutoNumber1" bordercolorlight="#F5F5F5" bordercolordark="#C0C0C0" STYLE="WORD-BREAK: break-all" background="images/postbit_lightbg.gif">   
<%
	UserID=request.cookies("UserID")
	SQL = "Select * From id Where user='"&request.cookies("UserID")&"'"
    set rs11=conn.execute(SQL)
	If request.cookies("UserPassed")<>"yes" or rs11.eof Then 
%><tr><td width="98%" background="images/images/postbit_lightbg.gif" colspan="6" width="5"
bordercolor="#D1D1D1" height="11"><p align="left"><FORM METHOD=POST ACTION="check.asp">
<% 
   response.cookies("TEST_COOKIE").Expires=date+365
response.cookies("TEST_COOKIE") = "testing" '登錄前必做!檢查COOKIE%>
<font size="2"><b><img border="0" src="images/login.gif">會員登錄</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
用戶帳號：<INPUT TYPE="text" NAME="UserID" size="10">&nbsp;&nbsp;&nbsp;&nbsp;帳號密碼：<INPUT TYPE="password" Name="Password" size="9">&nbsp;
<select size="1" name="cookie_save_day">
<option selected value="365">永遠</option>
<option value="30">一個月</option>
<option value="7">一周</option>
<option value="1">一天</option>
<option value="0">不用記住</option>
</select><INPUT TYPE="submit" VALUE="登入"><INPUT TYPE="reset" VALUE="重填">&nbsp;&nbsp;
(若是無法登入，請降低IE隱私權設定，請<font color="#FF0000">注意大小寫</font>)</font></FORM></p></td></tr><% End If %>
    <tr> 
		<td width="5%" background="images/Button.gif" height="13" bordercolor="#E3E3E3">
        <p align="center"><b><font size="2">狀態</font></b></td>
        <td width="40%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"> 
            <b> 
            <font size="2" color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;  </font> 
            <font size="2">分類名稱&nbsp;&nbsp;</font><font size="2" color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp; </font><font size="2">&nbsp; </font></b></td>
        
        <td width="12%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><b><font size="2">版主</font></b></td>
        <td width="6%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><font size="2"><b>主題數</b></font></td>
        <td width="6%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><font size="2"><b>回覆數</b></font></td>
        <td width="18%" height="13" bgcolor="#999999" bordercolor="#E3E3E3" background="images/Button.gif"> 
        <p align="center"><font size="2"><b>最後更新</b></font></td>
      </tr>

<%
sql="select * from sort order by sort_order"
set rs_sort=conn.execute(sql)
if not rs_sort.EOf then
while not rs_sort.EOf
%> 
<tr><td width="98%"  colspan="6" width="5" bgcolor="#FAFAFA"
bordercolor="#D1D1D1" height="11"><font size=-1><b>◎<%=rs_sort("sort_name")%>↓</b></font></td></tr>
<%
SQL= "Select * From tid where sort='"&rs_sort("sort_name")&"' Order By num "
Set rs=conn.Execute(SQL)
If rs.Eof Then 
	
Else
 While Not rs.EOF 
     ' 顯示資料
%>
      <tr> 
        <td width="3%" align="center" bgcolor="#F5F5F5" bordercolor="#D1D1D1" height="42" background="images/images/postbit_lightbg.gif">
      <%IF Now - rs("lastnewtime") < 1 then 
 	response.write "<img src=""images/icon_folder_new_3.gif"">"
else
    response.write "<img src=""images/icon_folder_3.gif"">"
end if%> 　</td>
        <td width="45%" height="42" bgcolor="#FFFFFF" bordercolor="#D1D1D1" align="left" background="images/postbit_lightbg.gif"> 
        <p align="left"><font color="#000080" size="2">　</font>
        <font color="#FF9900" size="2"><a href=forum.asp?tid=<%=rs("tid")%>>
        <font color="#CC6666"><b><%=rs("tidname")%></b></font></a></font> 
          <%if rs("have_pass")=1 then %><img src="images/key.gif" alt="需要密碼"><%end if%> <font size=-1><br><%=rs("text")%></font> </td>
        
<td width="8%" height="42" bgcolor="#FFFFFF" align="center" bordercolor="#D1D1D1"> 
<font color="#000080" size="2">　
<% 
if rs("tidadmin")<>"" then
tidadmin=split(rs("tidadmin"),"|")
for i = 1 to ubound(tidadmin)-1
response.write tidadmin(i)&"<br>"
next
end if
%></font></td>
        <td width="49" height="42" bgcolor="#F5F5F5" align="center" bordercolor="#D1D1D1" background="images/postbit_lightbg.gif"> 
          <font size="2" color=red><b><%=rs("counts")%></b></font></td>
        <td width="7%" height="42" bgcolor="#FFFFFF" align="center" bordercolor="#D1D1D1"> 
          <font size="2" color=red><b><%=rs("replycounts")%></b></font></td>
        <td width="14%" height="42" bgcolor="#EFF4FE" bordercolor="#D1D1D1" background="images/postbit_lightbg.gif"> 
        <p align="left"><font color="black" size="2">
		<%
'**************************************************************************
'判斷最新顯示的權限
sql="select * from tid where tid="&rs("tid")
set rs_forum=conn.execute(sql)
forum_cookies=bbsid+Cstr(rs("tid"))+request.cookies("UserID")
if request.cookies(forum_cookies)="yes" or  rs_forum("have_pass")=0 then	
'***************************************************************************
		%>
		<font color="#003399">文章：</font><%if rs("lasttitle")<>0 then
		  sql="select * from Titles where TitleID="&rs("lasttitle")
                   set rslast=conn.execute(sql)
                   if not rslast.EOF then
                   response.write "<a href=detail.asp?TitleID="&rs("lasttitle")&"&tid="&rs("tid")&"#top>"&Left(rslast("Subject"),5)&"..."&"</a>"
                   end if
		       elseif rs("lasttitle")=0 then
			     response.write Left(rs("lasttitle2"),5)&"..."
			end if
				
        %><br><font color=red>最後：</font><%=rs("lastposter")%>&nbsp;<img src="images/lastpost.gif">
        <br></font><font color="#7C7C7C" size=-1><%=rs("lastnewtime")%></font>
		<%
'**************************************************************************
else
  response.write "抱歉,您無權限"
end if
'***************************************************************************
		%>
		
		</td>
      </tr>
      <%
	  rs.MoveNext
	  Wend
	  END IF
	  	  rs_sort.movenext
	  wend

	  else
	    response.write "<p align=center><b><font color=red>●無主論壇分類!!</font></b></p>"
	  end if
	  %>
    </table>

 
     
    </select>
  </p>
 

<!--#include file="UserList.asp"-->
 <!--#include file="friend_web.asp"-->
<div align="center">
  <center>
  <br> <br> <br>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="98%" height="22">
    <tr>
      <td width="524" height="22" bordercolor="#FFFFFF" bgcolor="#EFF4FE" background="images/postbit_lightbg.gif"><font size=-1><b><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
%></b></font></td>
      <td width="5%" height="22" bgcolor="#EFF4FE" background="images/postbit_lightbg.gif">
<a href="#top"><img border="0" src="images/up.gif" width="20" height="20"></a></td>
    </tr>
  </table>
  </center>
</div>
<%end if'論壇關閉的endif%>
   
  <!--#include file="foot.asp"-->



</BODY>
</HTML>