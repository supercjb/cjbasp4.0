<%@LANGUAGE=VBScript codepage="950"%>
<!-- #include file="conn.asp" -->
<%Server.ScriptTimeOut=60 %>
<!-- #include file="uploadx.asp" -->
<%set fs=server.createobject("scripting.filesystemobject")
dim filename
path = Server.MapPath("upload/")
if not fs.folderexists (path) then
fs.createfolder path
end if
set fs = nothing
sql="select * from id where user='"&request.cookies("UserID")&"'"
set rs_level=conn.execute(sql)
sql2="select * from member_sort where level="&rs_level("level")
set member_sort=conn.execute(sql2)

filename = SaveFile("fruit",path,member_sort("upload_size"),0)
If filename<>"" Then 
If filename <> "*TooBig*" Then

 response.cookies("upload")="1"
response.cookies("filename")=cstr(filename)
 Response.Cookies("upload").Expires=date+1
  Response.Cookies("filename").Expires=date+1
Response.redirect "upload0.asp?message=附加"& filename &" 成功"

Else
Response.redirect "upload0.asp?message=檔案超過預設大小(<="&member_sort("upload_size")&"K)"
End IF
End IF%>