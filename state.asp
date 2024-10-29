<table border="1" width="98%">   
<tr><td width="69%" align="left"><!--#include file="UserList.asp"-->
<%response.write "<font size=-1 align=left>共有&nbsp;<font color=red><b>"&online_sum&"</font></b>&nbsp;位會員在線上，"
  response.write "共有&nbsp;<font color=red><b>"&online_guest&"</font></b>&nbsp;位訪客在線上。</font>"%>　</td>
<td width="31%" align="center" bgcolor="#EFF4FE"><p align="center" style="margin-top: 0; margin-bottom: 0"><font size=-1>
<%StrA = "select * from id order by ID"
StrB = "select * from Titles order by TitleID"
StrC = "select * from Details order by DetailID"
Response.Write "本論壇共計有&nbsp;" 
sum_anything StrA,rss
Response.Write "&nbsp;位註冊會員<br>並有&nbsp;"
sum_anything StrB,rss1
Response.Write "&nbsp;篇主題以及&nbsp;"
sum_anything StrC,rss2
Response.Write "&nbsp;篇回覆文章"
countobj="guest"%><!--#include file="count.asp"-->
<%Response.Write "<br><b><a href=index_counter.asp>流量統計</a></b> 自 2003/10/09  18:00  起<br>總訪客數： <font color=red><b>"&totalguest&"</b></font> &nbsp;&nbsp;&nbsp;<font color=black size=-1>今日訪客數： </font><font color=red><b>"&dayguest&"</b></font>"%>
</font></p></td></tr></table><!--#include file="top_new.asp"--><br>