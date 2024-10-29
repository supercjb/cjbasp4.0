
<%
if request.cookies("bbsid")<>bbsid then
response.redirect "loginout.asp"
elseif request.cookies("UserPassed")=empty then
response.redirect "loginout.asp"
elseif request.cookies("UserPassed")="no" then
response.redirect "loginout.asp"
end if
%>