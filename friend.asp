 <!--#include file="conn.asp"-->
<HTML><HEAD>
<TITLE> New Document </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT=""></HEAD>
<BODY>
<%if act="add" then
Set rs33=Server.CreateObject("ADODB.Recordset") 
rs33.open "friendpm",conn,1,3
rs33.addnew
rs33("user")=request.cookies("UserID")
rs33("friendname")=request("name")
rs33.update
rs33.close
end if%>
      <form method="get" >
        �W��<input type="text" name="name" size="20"></p>
        <input type="hidden" name="act" value="add" size="20">
        <p><input type="submit" value="�s�W" name="B1"><input type="reset" value="���s�]�w" name="B2"></p>
      </form>
<%    sql="select * from friendpm where user='"&request.cookies("UserID")&"'"
      set rs12=conn.execute(sql)
      
      if not rs12.eof then
      while not rs12.eof 
       response.write rs12("friendname")
	   rs12.movenext
      wend
      end if %>
</BODY></HTML>