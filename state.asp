<table border="1" width="98%">   
<tr><td width="69%" align="left"><!--#include file="UserList.asp"-->
<%response.write "<font size=-1 align=left>�@��&nbsp;<font color=red><b>"&online_sum&"</font></b>&nbsp;��|���b�u�W�A"
  response.write "�@��&nbsp;<font color=red><b>"&online_guest&"</font></b>&nbsp;��X�Ȧb�u�W�C</font>"%>�@</td>
<td width="31%" align="center" bgcolor="#EFF4FE"><p align="center" style="margin-top: 0; margin-bottom: 0"><font size=-1>
<%StrA = "select * from id order by ID"
StrB = "select * from Titles order by TitleID"
StrC = "select * from Details order by DetailID"
Response.Write "���׾¦@�p��&nbsp;" 
sum_anything StrA,rss
Response.Write "&nbsp;����U�|��<br>�æ�&nbsp;"
sum_anything StrB,rss1
Response.Write "&nbsp;�g�D�D�H��&nbsp;"
sum_anything StrC,rss2
Response.Write "&nbsp;�g�^�Ф峹"
countobj="guest"%><!--#include file="count.asp"-->
<%Response.Write "<br><b><a href=index_counter.asp>�y�q�έp</a></b> �� 2003/10/09  18:00  �_<br>�`�X�ȼơG <font color=red><b>"&totalguest&"</b></font> &nbsp;&nbsp;&nbsp;<font color=black size=-1>����X�ȼơG </font><font color=red><b>"&dayguest&"</b></font>"%>
</font></p></td></tr></table><!--#include file="top_new.asp"--><br>