<!--#include file="conn.asp"-->
<HTML><HEAD>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""></HEAD>
<BODY>
<%response.cookies("UserID")=empty
UserID = Request("UserID")   
Password1 = Request("Password1") 
Password2 = Request("Password2") 
birthday= Request("birthday") 
comefrom= Request("comefrom") 
email = Request("email") 
homepage= Request("homepage") 
msnm=Request("msnm")
yim=Request("yim")
timezone=Request("timezone")
pic= Request("pic")
face_url= Request("face_url") 
memo1=request("memo1")

Set rs2= Server.CreateObject("ADODB.Recordset") '檢查帳號是否被用
SQL="SELECT * FROM id WHERE user='" & UserID & "'"
rs2.Open SQL, conn,1,1
If Not rs2.EOF Then %>
<Script Language=Vbscript>
MsgBox "錯誤!!～帳號已存在",64,"帳號已存在!"
location.href = "javascript:history.back()"
</Script>

<%'若欄位留白
ElseIf UserID=Empty Or Password1=Empty Or birthday=Empty Or comefrom=Empty Or email=Empty Then%>
<Script Language=Vbscript>
MsgBox "錯誤!!～必填欄位留白",64,"欄位留白!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Password1 <> Password2 Then '檢查密碼%>
<Script Language=Vbscript>
MsgBox "錯誤!!～密碼錯誤",64,"密碼錯誤!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(email,"@")=0 Or Right(email,1)="@" Or Left(email,1)="@" Then '檢查email%>
<Script Language=Vbscript>
MsgBox "錯誤!!～mail格式錯誤",64,"mail錯誤!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(yim,"@")<>0 Then '檢查奇摩帳號%>
<Script Language=Vbscript>
MsgBox "錯誤!!～奇摩即時通帳號格式錯誤，只填帳號即可",64,"帳號錯誤!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(birthday,"/") = 0 Or Mid(birthday,5,1) <> "/" Then '檢查birthday%>
<Script Language=Vbscript>
MsgBox "錯誤!!～生日格式錯誤",64,"birthday錯誤!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(homepage,"http://")=0 and homepage<>empty then%>
<Script Language=Vbscript>
MsgBox "錯誤!!～首頁格式錯誤",64,"http://錯誤!"
location.href = "javascript:history.back()"
</Script>
<%Else '帳號若沒被用
Set rs = Server.CreateObject("ADODB.Recordset")
 
SQL = "Select * From id "
rs.Open SQL, conn,1,3
			
rs.AddNew
rs("registtime")=now()
rs("user")=UserID
rs("email")=email
rs("password")=Password1
rs("birthday")=birthday
rs("comefrom")=comefrom
rs("msnm")=msnm
rs("yim")=yim
rs("timezone")=timezone
rs("face_url")=face_url
if homepage=empty then
rs("homepage")=""
else
rs("homepage")=homepage
end if
rs("pic")=pic
rs("memo1")=memo1
rs.Update
rs.close

SQL = "Select * From config_setting Where ID=1 "
rs.Open SQL, conn,1,3
rs("newregist")=UserID
rs.Update
     
response.redirect "login.asp?mod=reg&msg=恭喜您註冊成功!請登錄論壇~"
End if
'檢查帳號是否被用的END%>
</BODY></HTML>