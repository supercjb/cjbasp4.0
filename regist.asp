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

Set rs2= Server.CreateObject("ADODB.Recordset") '�ˬd�b���O�_�Q��
SQL="SELECT * FROM id WHERE user='" & UserID & "'"
rs2.Open SQL, conn,1,1
If Not rs2.EOF Then %>
<Script Language=Vbscript>
MsgBox "���~!!��b���w�s�b",64,"�b���w�s�b!"
location.href = "javascript:history.back()"
</Script>

<%'�Y���d��
ElseIf UserID=Empty Or Password1=Empty Or birthday=Empty Or comefrom=Empty Or email=Empty Then%>
<Script Language=Vbscript>
MsgBox "���~!!�㥲�����d��",64,"���d��!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Password1 <> Password2 Then '�ˬd�K�X%>
<Script Language=Vbscript>
MsgBox "���~!!��K�X���~",64,"�K�X���~!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(email,"@")=0 Or Right(email,1)="@" Or Left(email,1)="@" Then '�ˬdemail%>
<Script Language=Vbscript>
MsgBox "���~!!��mail�榡���~",64,"mail���~!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(yim,"@")<>0 Then '�ˬd�_���b��%>
<Script Language=Vbscript>
MsgBox "���~!!��_���Y�ɳq�b���榡���~�A�u��b���Y�i",64,"�b�����~!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(birthday,"/") = 0 Or Mid(birthday,5,1) <> "/" Then '�ˬdbirthday%>
<Script Language=Vbscript>
MsgBox "���~!!��ͤ�榡���~",64,"birthday���~!"
location.href = "javascript:history.back()"
</Script>

<%ElseIf Instr(homepage,"http://")=0 and homepage<>empty then%>
<Script Language=Vbscript>
MsgBox "���~!!�㭺���榡���~",64,"http://���~!"
location.href = "javascript:history.back()"
</Script>
<%Else '�b���Y�S�Q��
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
     
response.redirect "login.asp?mod=reg&msg=���߱z���U���\!�еn���׾�~"
End if
'�ˬd�b���O�_�Q�Ϊ�END%>
</BODY></HTML>