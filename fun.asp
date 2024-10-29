<!--#include file="conn.asp"--><%
function subjectcolor(string,rs)
        response.write "<font color="&rs("Color")&">"&string&"</font>"
end function

function ChkBadWords(fString)
   if fString<>"" then
    bwords = split(BadWords, "|")
    for i = 0 to ubound(bwords)
        fString = Replace(fString, bwords(i), string(len(bwords(i)),"*")) 
    next
  end if
    ChkBadWords = fString
end function

function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp; ")
    fString = Replace(fString, CHR(13), "")
'    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
end if
end function

function UBBCode(strContent)
    dim re
    Set re=new RegExp
    re.IgnoreCase =true
    re.Global=True

'    if strIMGInPosts = "1" then
    re.Pattern="(\[IMG\])(.[^\[]*)(\[\/IMG\])"
    strContent=re.Replace(strContent,"<a href=""$2"" target=_blank><IMG SRC=""$2"" border=0 alt=按此在新窗口瀏覽圖片 onload=""javascript:if(this.width>screen.width-160)this.width=screen.width-160""></a>")
'    end if

re.Pattern="(\[hide\])(.*)(\[\/hide\])"

sql="select * from Details where TitleID="&request("TitleID")&" and Name='"&request.cookies("UserID")&"'"
set see=conn.execute(sql)
SQL = "Select * From tid where tid="&request("tid")
Set tid_admin_p = conn.Execute(SQL)
mastera=split(tid_admin_p("tidadmin"),"|")
for i = 1 to ubound(mastera)-1
if (not see.eof) or request("postname")=request.cookies("UserID") or request.cookies("admin_pass")=bbsid or mastera(i)=request.cookies("UserID") then
strContent=re.Replace(strContent,"<center><font style='font-size:9pt;line-height:15pt' color=red><hr color=#CDCFCE size=1 width=400>以下是隱藏的內容：</font><BR>$2<hr color=#CDCFCE size=1 width=400></center>")
else
strContent=re.Replace(strContent,"<center><font style='font-size:9pt;line-height:15pt' color=red><hr color=#CDCFCE size=1 width=400>此內容只有作者和已經回覆此文章的瀏覽者能瀏覽：</font><BR><hr color=#CDCFCE size=1 width=400></center>")
end if
next
    re.Pattern="\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
    strContent=re.Replace(strContent,"<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
    re.Pattern="\[QT=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/QT]"
    strContent=re.Replace(strContent,"<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
    re.Pattern="\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
    strContent=re.Replace(strContent,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed></object>")
    re.Pattern="\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
    strContent=re.Replace(strContent,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")

 '   if strflash= "1" then
    re.Pattern="(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
    strContent= re.Replace(strContent,"<OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=423 height=300><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")
 '   end if

    re.Pattern="(\[ZIP\])(.[^\[]*)(\[\/ZIP\])"
    strContent=re.Replace(strContent,"<br><IMG SRC=image/zip.gif border=0> <a href=""$2"">點擊下載該文件</a>")
    re.Pattern="(\[RAR\])(.[^\[]*)(\[\/RAR\])"
    strContent=re.Replace(strContent,"<br><IMG SRC=image/rar.gif border=0> <a href=""$2"">點擊下載該文件</a>")
    re.Pattern="(\[UPLOAD=(.[^\[]*)\])(.[^\[]*)(\[\/UPLOAD\])"
    strContent= re.Replace(strContent,"<br><IMG SRC=""image/$2.gif"" border=0>此主題相關圖片如下：<br><A HREF=""$3"" TARGET=_blank><IMG SRC=""$3"" border=0 alt=按此在新窗口瀏覽圖片 onload=""javascript:if(this.width>screen.width-160)this.width=screen.width-160""></A>")

    re.Pattern="(\[URL\])(http:\/\/.[^\[]*)(\[\/URL\])"
    strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
    re.Pattern="(\[URL\])(.[^\[]*)(\[\/URL\])"
    strContent= re.Replace(strContent,"<A HREF=""http://$2"" TARGET=_blank>$2</A>")

    re.Pattern="(\[URL=(http:\/\/.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
    strContent= re.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$3</A>")
    re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
    strContent= re.Replace(strContent,"<A HREF=""http://$2"" TARGET=_blank>$3</A>")

    re.Pattern="(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
    strContent= re.Replace(strContent,"<A HREF=""mailto:$2"">$2</A>")
    re.Pattern="(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
    strContent= re.Replace(strContent,"<A HREF=""mailto:$2"" TARGET=_blank>$3</A>")

	re.Pattern = "^(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "^(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "(ftp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](ftp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "^(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "(rtsp://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](rtsp://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "^(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "(mms://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")
	re.Pattern = "[^>=""](mms://[A-Za-z0-9\.\/=\?%\-&_~`@':+!]+)"
	strContent = re.Replace(strContent,"<a target=_blank href=$1>$1</a>")

    if strIcons = "1" then
    re.Pattern="(\[em(.[^\[]*)\])"
    strContent=re.Replace(strContent,"<img src=image/em$2.gif border=0 align=middle>")
    end if

    re.Pattern="(\[HTML\])(.[^\[]*)(\[\/HTML\])"
    strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='"&abgcolor&"'><td><b>以下內容為程序代碼:</b><br>$2</td></table>")
    re.Pattern="(\[code\])(.[^\[]*)(\[\/code\])"
    strContent=re.Replace(strContent,"<table width='100%' border='0' cellspacing='0' cellpadding='6' bgcolor='"&abgcolor&"'><td><b>以下內容為程序代碼:</b><br>$2</td></table>")

    re.Pattern="(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])"
    strContent=re.Replace(strContent,"<font color=$2>$3</font>")
    re.Pattern="(\[face=(.[^\[]*)\])(.[^\[]*)(\[\/face\])"
    strContent=re.Replace(strContent,"<font face=$2>$3</font>")
    re.Pattern="(\[align=(.[^\[]*)\])(.*)(\[\/align\])"
    strContent=re.Replace(strContent,"<div align=$2>$3</div>")
    re.Pattern="(\[hr=*(#*[a-z0-9]*)(.[^\[]*))"
    strContent=re.Replace(strContent,"<hr color=$2 width=$3>") 

    re.Pattern="(\[QUOTE\])(.*)(\[\/QUOTE\])"
    strContent=re.Replace(strContent,"<table cellpadding=0 cellspacing=0 border=1 WIDTH=94% bgcolor=#F5F5F5 align=center><tr><td><table width=100% cellpadding=5 cellspacing=1 border=0><TR><TD BGCOLOR='"&abgcolor&"'><font size=-1>$2</font></table></table><br>")
    
    re.Pattern="(\[fly\])(.*)(\[\/fly\])"
    strContent=re.Replace(strContent,"<marquee width=90% behavior=alternate scrollamount=3>$2</marquee>")
    re.Pattern="(\[move\])(.*)(\[\/move\])"
    strContent=re.Replace(strContent,"<MARQUEE scrollamount=3>$2</marquee>")	
    re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
    strContent=re.Replace(strContent,"<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
    re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
    strContent=re.Replace(strContent,"<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")

    re.Pattern="(\[i\])(.[^\[]*)(\[\/i\])"
    strContent=re.Replace(strContent,"<i>$2</i>")
    re.Pattern="(\[u\])(.[^\[]*)(\[\/u\])"
    strContent=re.Replace(strContent,"<u>$2</u>")
    re.Pattern="(\[b\])(.[^\[]*)(\[\/b\])"
    strContent=re.Replace(strContent,"<b>$2</b>")
    re.Pattern="(\[fly\])(.[^\[]*)(\[\/fly\])"
    strContent=re.Replace(strContent,"<marquee>$2</marquee>")

    re.Pattern="(\[size=1\])(.[^\[]*)(\[\/size\])"
    strContent=re.Replace(strContent,"<font size=1>$2</font>")
    re.Pattern="(\[size=2\])(.[^\[]*)(\[\/size\])"
    strContent=re.Replace(strContent,"<font size=2>$2</font>")
    re.Pattern="(\[size=3\])(.[^\[]*)(\[\/size\])"
    strContent=re.Replace(strContent,"<font size=3>$2</font>")
    re.Pattern="(\[size=4\])(.[^\[]*)(\[\/size\])"
    strContent=re.Replace(strContent,"<font size=4>$2</font>")
    re.Pattern="(\[size=5\])(.[^\[]*)(\[\/size\])"
    strContent=re.Replace(strContent,"<font size=5>$2</font>")

    re.Pattern="(\[center\])(.[^\[]*)(\[\/center\])"
    strContent=re.Replace(strContent,"<center>$2</center>")
' end if
    strContent=ChkBadWords(strContent)

    set re=Nothing
    UBBCode=strContent
 end function

  function sum_anything(sql,rs)
	Set rs = conn.Execute(sql)
		i=0
	while not rs.eof
		i=i+1
		rs.movenext
	wend
	Response.Write "<font color=red><b>"& i &"</b></font>"
  end function%>