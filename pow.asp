<%if request.cookies("UserID")=empty then
 response.redirect "registrule.asp"
end if%>