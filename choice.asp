<font size=-1>
<%'��ܦb���Ӫ�
if IsNumeric(request.cookies("tid")) then
Set rcs=Server.CreateObject("ADODB.Recordset")
rcs.Open "Select * From tid where tid =" & request.cookies("tid"),conn,1,1
end if
Set rc2s=Server.CreateObject("ADODB.Recordset")
rc2s.Open "Select * From config_setting where ID=1",conn,1,1
Response.write "<img border=0 src=""images/icon_folder_open.gif"">"&"<a href=index.asp>"&rc2s("bbsname")&"</a> -> "
if IsNumeric(request.cookies("tid")) then
Response.write "<img border=0 src=""images/icon_folder_open_topic.gif"">"&"<a href=forum.asp?tid="&request.cookies("tid")&">"&rcs("tidname")&"</a>"
else
Response.write "<img border=0 src=""images/icon_folder_open_topic.gif"">"&request.cookies("tid")
end if
if not request("TitleID")=empty then
sql="select * from Titles where TitleID="&request("TitleID")
set rs_cho=conn.execute(sql)
response.write " -> "
Response.write "<img border=0 src=""images/icon_folder_open_topic.gif"">"&rs_cho("Subject")
end if
if request("sortx")=1 then
response.write "  -> �z�o���D�D(sort by post_time) "
elseif  request("sortx")=2 then
response.write "  -> �z�o���D�D(sort by reply_time) "
elseif  request("sortx")=3 then
response.write "  -> �˵��|���^�Ъ��Ҧ��D�D "
elseif  request("sortx")=4 then
response.write "  -> �˵� "&request("timediff")&" �Ѥ����Ҧ��D�D "
end if%>
</font>