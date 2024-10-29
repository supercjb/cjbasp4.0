<html>
<!--#include file="conn.asp"-->
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><%sql="select * from config_setting"
set rstitle=conn.execute(sql)
response.write rstitle("bbsname")
if request.cookies("level")=empty then
response.redirect "loginout.asp"
end if
%> </title>
<link rel="stylesheet" type="text/css" href=style1.css>
</head>

<body leftmargin="0" topmargin="0">
<!--#include file="header.vs"-->
<!--#include file="admin/broad.vs"--><br>
<% response.cookies("tid")="個人簡訊管理"%>
<!--#include file="choice.vs"-->
<FORM METHOD=POST ACTION="delete_pm.asp">

<%

Set rs=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM u2u WHERE touser='" & request.cookies("UserID") & "'"
SQL=SQL & "Order By time Desc"
rs.Open SQL,conn,1,3
set rs22=conn.execute(SQL)
i=0
while not rs22.eof
i=i+1
rs22.movenext
wend

sql2="select * from member_sort where level="&request.cookies("level")
set rs_level2=conn.execute(sql2)
if request.cookies("level")=empty then
response.redirect "loginout.asp"
end if
sql="select * from u2u where touser='"&request.cookies("UserID")&"'"
set rsb=conn.execute(sql)
i=0
while not rsb.eof
i=i+1
rsb.movenext
wend
if i>(rs_level2("pm_size")+3) then

while i>(rs_level2("pm_size")+3)
rs.movefirst
rs.delete
rs.movenext
response.redirect "pm.asp"
wend
response.redirect "pm.asp"
elseif i>=rs_level2("pm_size") then
response.write "<center><font color=red><b>空間已滿,請盡快刪除!!~</b></font></center>"
end if%>



　<p align="right" style="margin-top: 0; margin-bottom: 0"><a href="u2u.asp"><img border="0" src="images/newpm.gif"></a><a href="pm.asp"><img border="0" src="images/untitled.bmp"></a><a href="friendpm.asp"><img border="0" src="images/friendpm.gif"></a>&nbsp;&nbsp;&nbsp;
</p>
<p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size=-1><font color=red><b><%=request.cookies("UserID")%></font>&nbsp;&nbsp;的簡訊收件夾</b>，您可以存<b><%=rs_level2("pm_size")%></b>封簡訊，已有<%=i%>封簡訊</font>
<BR>
</p><center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#0066CC" width="96%" STYLE="WORD-BREAK: break-all">
  <tr>
    <td width="100%" align="center" colspan="5" background="images/bg2.gif">
    <p align="left"><font size=-1>&nbsp;狀態&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    簡訊主題&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 來自&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    寄到時間&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;執行</td>
  </tr>
 <%
     Page = cint(Request("Page"))
	 myself=request.serverVariables("PATH_INFO")
	
	  rs.PageSize =10      ' 設定每頁顯示 10 筆
     If Page<1 Then 
		Page=1
     ELSEIF Page>rs.PageCount Then 
		Page=rs.PageCount
     END IF
    
%>

  <%
If rs.Eof Then 
    	RESPONSE.WRITE "</table><FONT COLOR=maroon size=-1>Sorry !!尚無訊息...</FONT>"
%>

<%
Else%>
<%
  rs.AbsolutePage = Page ' 將資料錄移至 PAGE 頁
FOR SH=1 to rs.PageSize      ' 顯示資料
     
%>

  <tr>
    <td width="5%"align="center"><font size=-1>
	<%If rs("readed")=1 Then%><img src="images/icon_folder_new_2.gif">
	<%Else%><img src="images/icon_folder_2.gif"><%End If%>
	
	　</td>
    <td width="47%"><font size=-1><!-- 假使沒讀過就顯示粗體1,簡訊讀過是0 -->
	<%If rs("readed")=1 Then%><A HREF="pm_detail.asp?id=<%=rs("id")%>"><b><%=rs("totitle")%></b></a>　</td>
     <%else%>
<A HREF="pm_detail.asp?id=<%=rs("id")%>"><%=rs("totitle")%></a></td>
      <%end if%>
    <td width="20%" align="center"><font size=-1><b><%=rs("fromuser")%></b>　</td>
    <td width="19%" align="center"><font size=-1><%=rs("time")%>　</td>
    <td width="9%">
    <p align="center"><font size="2">刪除<input type="checkbox" name="<%=rs("id")%>" value="ON"><INPUT TYPE="hidden"NAME></font></td>
  </tr>
  <%
  rs.MoveNext
IF rs.EOF THEN EXIT FOR
Next

  %>
  </table>




  
<p style="margin-top: 0; margin-bottom: 0" align="right">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input TYPE="submit" value="確定刪除" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </p>
<%END IF%>
<p style="margin-top: 0; margin-bottom: 0">


<p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<img border="0" src="images/icon_folder_new_2.gif"><font size="2">新的簡訊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<img border="0" src="images/icon_folder_2.gif">舊的簡訊</font></p>
</form>

  <!-- 快速跳頁功能 -->
 <%If rs.PageCount=0 Then%>
<p align="right"><font size=-1>頁次1/1</font></p>
<%Else%>
  <form action=<%=myself%> method=GET>
 <p align="right"><font size=-1>
 頁次<%=Page%>/<%=rs.PageCount%>
 <%If Page<>1 Then%>
<a href="<%=myself%>?Page=1&tid=<%=tid%>">[1]....</a>
<a href="<%=myself%>?Page=<%=(Page-1)%>&tid=<%=tid%>">[上一頁]</a>
<%End If%>
<%If Page<>rs.PageCount Then%>
<a href="<%=myself%>?Page=<%=(Page+1)%>&tid=<%=tid%>">[下一頁]</a>
<a href="<%=myself%>?Page=<%=rs.PageCount%>&tid=<%=tid%>">....[<%=rs.PageCount%>]</a>
<%End If%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
輸入頁次:<INPUT TYPE=TEXT NAME=Page SIZE=2>  
<INPUT TYPE=hidden NAME=tid VALUE="<%=tid%>">
</form>
</p>
 
 <%End If%> </font>
<hr> 
 
 
 
 
 








<FORM METHOD=POST ACTION="delete_pm.asp">

<%

Set rs=Server.CreateObject("ADODB.Recordset")
SQL="SELECT * FROM u2u WHERE fromuser='" & request.cookies("UserID") & "'"
SQL=SQL & "Order By time Desc"
rs.Open SQL,conn,1,3
set rs22=conn.execute(SQL)

%>



<p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size=-1><font color=red><b><%=request.cookies("UserID")%></font>&nbsp;&nbsp;的簡訊發送備忘夾</b></font>
<BR>
</p><center>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#0066CC" width="96%" STYLE="WORD-BREAK: break-all">
  <tr>
    <td width="100%" align="center" colspan="5" background="images/bg2.gif">
    <p align="left"><font size=-1>&nbsp;狀態&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
    簡訊主題&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 傳送對象&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    傳送時間&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;執行 
    </font>
	</td>
  </tr>
 <%
     Page = cint(Request("Page"))
	 myself=request.serverVariables("PATH_INFO")
	
	  rs.PageSize =10      ' 設定每頁顯示 10 筆
     If Page<1 Then 
		Page=1
     ELSEIF Page>rs.PageCount Then 
		Page=rs.PageCount
     END IF
    
%>

  <%
If rs.Eof Then 
    	RESPONSE.WRITE "</table><FONT COLOR=maroon size=-1>Sorry !!尚無訊息...</FONT>"
%>

<%
Else%>
<%
  rs.AbsolutePage = Page ' 將資料錄移至 PAGE 頁
FOR SH=1 to rs.PageSize      ' 顯示資料
     
%>

  <tr>
    <td width="5%"align="center"><font size=-1>
	<%If rs("readed")=1 Then%><img src="images/icon_folder_new_2.gif">
	<%Else%><img src="images/icon_folder_2.gif"><%End If%>
	
	　</td>
    <td width="47%"><font size=-1><!-- 假使沒讀過就顯示粗體1,簡訊讀過是0 -->
	<%If rs("readed")=1 Then%><A HREF="pm_detail2.asp?id=<%=rs("id")%>"><b><%=rs("totitle")%></b></a>　</td>
     <%else%>
<A HREF="pm_detail2.asp?id=<%=rs("id")%>"><%=rs("totitle")%></a></td>
      <%end if%>
    <td width="20%" align="center"><font size=-1><b><%=rs("touser")%></b>　</td>
    <td width="19%" align="center"><font size=-1><%=rs("time")%>　</td>
    <td width="9%">
    <p align="center"><font size="2">刪除<input type="checkbox" name="<%=rs("id")%>" value="ON"><INPUT TYPE="hidden"NAME></font></td>    
  </tr>
  <%
  rs.MoveNext
IF rs.EOF THEN EXIT FOR
Next

  %>
  </table>




  
<p style="margin-top: 0; margin-bottom: 0" align="right">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<input TYPE="submit" value="確定刪除" name="B1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </p>
<%END IF%>
<p style="margin-top: 0; margin-bottom: 0">


<p align="left" style="margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<img border="0" src="images/icon_folder_new_2.gif"><font size="2">對方未收的簡訊</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<img border="0" src="images/icon_folder_2.gif">對方已收的簡訊</font></p>
</form>

  <!-- 快速跳頁功能 -->
 <%If rs.PageCount=0 Then%>
<p align="right"><font size=-1>頁次1/1</font></p>
<%Else%>
  <form action=<%=myself%> method=GET>
 <p align="right"><font size=-1>
 頁次<%=Page%>/<%=rs.PageCount%>
 <%If Page<>1 Then%>
<a href="<%=myself%>?Page=1&tid=<%=tid%>">[1]....</a>
<a href="<%=myself%>?Page=<%=(Page-1)%>&tid=<%=tid%>">[上一頁]</a>
<%End If%>
<%If Page<>rs.PageCount Then%>
<a href="<%=myself%>?Page=<%=(Page+1)%>&tid=<%=tid%>">[下一頁]</a>
<a href="<%=myself%>?Page=<%=rs.PageCount%>&tid=<%=tid%>">....[<%=rs.PageCount%>]</a>
<%End If%>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
輸入頁次:<INPUT TYPE=TEXT NAME=Page SIZE=2>  
<INPUT TYPE=hidden NAME=tid VALUE="<%=tid%>">
</form>
</p>
 
 <%End If%> </font>
<hr>





<!--#include file="foot.vs"-->





</body>

</html>