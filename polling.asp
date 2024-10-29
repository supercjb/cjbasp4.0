<!--#include file="safe.asp"-->
<%
set write_poll=Server.CreateObject("ADODB.Recordset")
sql="select * from poll where titleid="&request("TitleID")
write_poll.open sql,conn,1,3
if request("cho")="T1" then
write_poll("T1C")=write_poll("T1C")+1
elseif  request("cho")="T2" then
write_poll("T2C")=write_poll("T2C")+1
elseif  request("cho")="T3" then
write_poll("T3C")=write_poll("T3C")+1
elseif  request("cho")="T4" then
write_poll("T4C")=write_poll("T4C")+1
elseif  request("cho")="T5" then
write_poll("T5C")=write_poll("T5C")+1
elseif  request("cho")="T6" then
write_poll("T6C")=write_poll("T6C")+1
elseif  request("cho")="T7" then
write_poll("T7C")=write_poll("T7C")+1
elseif  request("cho")="T8" then
write_poll("T8C")=write_poll("T8C")+1
elseif  request("cho")="T9" then
write_poll("T9C")=write_poll("T9C")+1
elseif  request("cho")="T10" then
write_poll("T10C")=write_poll("T10C")+1
else
end if
write_poll.update
set write_poller=Server.CreateObject("ADODB.Recordset")
sql="select * from poller"
write_poller.open sql,conn,1,3
write_poller.addnew
write_poller("titleid")=request("TitleID")
write_poller("name")=request.cookies("UserID")
write_poller.update
response.redirect "detail.asp?TitleID="&request("TitleID")&"&tid="&request("tid")&"&postname="&request("postname")
%>