
function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 2) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 3) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}
function AddText(NewCode){
	document.InputForm.Words.value+=NewCode;
	document.InputForm.Words.focus();
}

function email() {
	if (helpstat) {
		alert("Email 連結標記\n插入 Email 超連結\n用法1: [email]nobody@domain.com[/email]\n用法2: [email=nobody@domain.com]佚名[/email]");
	} else if (basic) {
		AddTxt="[email][/email]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("連結顯示的文字.\n如果空白，那麼將只顯示連結的 Email 地址",""); 
		if (txt2!=null) {
			txt=prompt("Email 位址.","name@domain.com");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[email]"+txt+"[/email]";
				} else {
					AddTxt="[email="+txt+"]"+txt2;
					AddText(AddTxt);
					AddTxt="[/email]";
				} 
				AddText(AddTxt);	        
			}
		}
	}
}

function flash() {
 	if (helpstat){
		alert("Flash 動畫\n插入 Flash 動畫.\n用法: [flash=寬度,高度]Flash 動畫文件的位址[/flash]");
	} else if (basic) {
		AddTxt="[flash][/flash]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("請輸入Flash 文件的位址：","http://");
		if (txt!=null) {             
			AddTxt="[flash]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
		}        
	}  
}

function hide() {
 	if (helpstat){
		alert("文章隱藏\n插入 隱藏語法.\n用法: [hide]內容[/hide]");
	} else if (basic) {
		AddTxt="[hide][/hide]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("請輸入隱藏文件的內容：","內容");
		if (txt!=null) {             
			AddTxt="[hide]"+txt;
			AddText(AddTxt);
			AddTxt="[/hide]";
			AddText(AddTxt);
		}        
	}  
}



function director() {
 	if (helpstat){
		alert("Shockwave 動畫\n插入 Shockwave 動畫.\n用法: [dir=寬度,高度]Shockwave 動畫文件的位址[/dir]");
	} else if (basic) {
		AddTxt="[dir][/dir]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("請輸入Shockwave 動畫文件的位址：","http://");
		if (txt!=null) {             
			AddTxt="[dir=360,320]"+txt;
			AddText(AddTxt);
			AddTxt="[/dir]";
			AddText(AddTxt);
		}        
	}  
}

function rm() {
 	if (helpstat){
		alert("RealPlayer 視訊檔案\n插入 RealPlayer *.rm 視訊檔案.\n用法: [rm=寬度,高度]RealPlayer 視訊檔的位址[/rm]");
	} else if (basic) {
		AddTxt="[rm][/rm]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("請輸入RealPlayer 視訊檔案的位址：","http://");
		if (txt!=null) {             
			AddTxt="[rm=]"+txt;
			AddText(AddTxt);
			AddTxt="[/rm]";
			AddText(AddTxt);
		}        
	}  
}

function wmv() {
 	if (helpstat){
		alert("Windows Media Player 視訊檔案\n插入 Windows Media Player *.wmv 視訊檔案.\n用法: [mp=寬度,高度]Windows Media Player 視訊檔案的位址[/mp]");
	} else if (basic) {
		AddTxt="[mp][/mp]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("請輸入Windows Media Player 視訊檔案的位址：","http://");
		if (txt!=null) {             
			AddTxt="[mp=360,320]"+txt;
			AddText(AddTxt);
			AddTxt="[/mp]";
			AddText(AddTxt);
		}        
	}  
}

function mov() {
 	if (helpstat){
		alert("QuickTime 視訊檔案\n插入 QuickTime *.mov 視訊檔.\n用法: [qt=寬度,高度]QuickTime 視訊檔案的位址[/qt]");
	} else if (basic) {
		AddTxt="[qt][/qt]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("請輸入QuickTime 視訊檔的位址：","http://");
		if (txt!=null) {             
			AddTxt="[qt=360,320]"+txt;
			AddText(AddTxt);
			AddTxt="[/qt]";
			AddText(AddTxt);
		}        
	}  
}


function setsize(size) {
	if (helpstat) {
		alert("文字大小標記\n設置文字大小.\n範圍 1 - 4.\n 1 為最小 4 為最大.\n用法: [size="+size+"]這是 "+size+" 文字[/size]");
	} else if (basic) {
		AddTxt="[size="+size+"][/size]";
		AddText(AddTxt);
	} else {                       
		txt=prompt("字體大小:"+size+"\n請輸入文字：","文字"); 
		if (txt!=null) {             
			AddTxt="[size="+size+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/size]";
			AddText(AddTxt);
		}        
	}
}

function bold() {
	if (helpstat) {
		alert("粗體標記\n使文字變為粗體.\n用法: [b]這是粗體文字[/b]");
	} else if (basic) {
		AddTxt="[b][/b]";
		AddText(AddTxt);
	} else {  
		txt=prompt("請輸入欲變成粗體的文字：","文字");     
		if (txt!=null) {           
			AddTxt="[b]"+txt;
			AddText(AddTxt);
			AddTxt="[/b]";
			AddText(AddTxt);
		}       
	}
}

function italicize() {
	if (helpstat) {
		alert("斜體標記\n使文字變為斜體.\n用法: [i]這是斜體字[/i]");
	} else if (basic) {
		AddTxt="[i][/i]";
		AddText(AddTxt);
	} else {   
		txt=prompt("請輸入欲變成斜體的文字：","文字");     
		if (txt!=null) {           
			AddTxt="[i]"+txt;
			AddText(AddTxt);
			AddTxt="[/i]";
			AddText(AddTxt);
		}	        
	}
}

function quote() {
	if (helpstat){
		alert("引用標記\n引用一些文字.\n用法: [quote]引用內容[/quote]");
	} else if (basic) {
		AddTxt="[quote][/quote]";
		AddText(AddTxt);
	} else {   
		txt=prompt("被引用的文字","文字");     
		if(txt!=null) {          
			AddTxt="[quote]"+txt;
			AddText(AddTxt);
			AddTxt="[/quote]";
			AddText(AddTxt);
		}	        
	}
}

function setcolor(color) {
	if (helpstat) {
		alert("顏色標記\n設置文字顏色.  任何顏色都可以被使用.\n用法: [color="+color+"]顏色要改變為"+color+"的文字[/color]");
	} else if (basic) {
		AddTxt="[color="+color+"][/color]";
		AddText(AddTxt);
	} else {  
     	txt=prompt("選擇的顏色是: "+color+"\n請輸入文字：","文字");
		if(txt!=null) {
			AddTxt="[color="+color+"]"+txt;
			AddText(AddTxt);        
			AddTxt="[/color]";
			AddText(AddTxt);
		} 
	}
}

function center() {
 	if (helpstat) {
		alert("對齊標記\n使用這個標記, 可以使文字靠左對齊、置中、靠右對齊.\n用法: [align= center | left | right ]要對齊的文字[/align]");
	} else if (basic) {
		AddTxt="[align=left][/align]";
		AddText(AddTxt);
	} else {  
		txt2=prompt("對齊樣式\n輸入 'center' 表示置中, 'left' 表示靠左對齊, 'right' 表示靠右對齊.","center");               
		while ((txt2!="") && (txt2!="center") && (txt2!="left") && (txt2!="right") && (txt2!=null)) {
			txt2=prompt("錯誤!\n類型只能輸入 'center' 、 'left' 或者 'right'.","");               
		}
		txt=prompt("請輸入要對齊的文字：","文字");     
		if (txt!=null) {          
			AddTxt="[align="+txt2+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/align]";
			AddText(AddTxt);
		}	       
	}
}

function hyperlink() {
	if (helpstat) {
		alert("超連結標記\n插入一個超連結標記\n方法: [url]http://domain.com[/url]\n或是: [url=http://domain.com]連結文字[/url]");
	} else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("超連結顯示文字.\n如果不想使用, 可以空白, 將只顯示超連結位址. ",""); 
		if (txt2!=null) {
			txt=prompt("請輸入超連結位置：","http://");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[url]"+txt;
					AddText(AddTxt);
					AddTxt="[/url]";
					AddText(AddTxt);
				} else {
					AddTxt="[url="+txt+"]"+txt2;
					AddText(AddTxt);
					AddTxt="[/url]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function image() {
	if (helpstat){
		alert("圖片標記\n插入圖片\n用法： [img]http://domain.com/test.gif[/img]");
	} else if (basic) {
		AddTxt="[img][/img]";
		AddText(AddTxt);
	} else {  
		txt=prompt("請輸入圖片的 URL位置：","http://");    
		if(txt!=null) {            
			AddTxt="[img]"+txt;
			AddText(AddTxt);
			AddTxt="[/img]";
			AddText(AddTxt);
		}	
	}
}

function code() {
	if (helpstat) {
		alert("程式碼標記\n使用程式碼標記，可以使你的程式碼裡面的 html 等標記不會被破壞.\n方法:\n [code]這裡是程式碼文字[/code]");
	} else if (basic) {
		AddTxt="[code][/code]";
		AddText(AddTxt);
	} else {   
		txt=prompt("輸入代碼","");     
		if (txt!=null) {          
			AddTxt="[code]"+txt;
			AddText(AddTxt);
			AddTxt="[/code]";
			AddText(AddTxt);
		}	       
	}
}

function list() {
	if (helpstat) {
		alert("清單列表標記\n產生一個清單列表.\n\n用法: [list=標記類型( square | circle )] [*]項目一 [*]項目二 [*]項目三 [/list]");
	} else if (basic) {
		AddTxt="[list][*]  [*]  [*]  [/list]";
		AddText(AddTxt);
	} else {  
		txt=prompt("輸入列表類型\n'square'為方塊,'circle'為白圓圈,空白為黑圓圈.","");
		//if (txt==null){ return false; }               
		while ((txt!="") && (txt.toLowerCase()!="square") && (txt.toLowerCase()!="circle") && (txt!=null)) {
			txt=prompt("錯誤!\n類型只能輸入 'square' 、 'circle' 或者空白.","");               
		}
		if (txt!=null) {
			if (txt=="") {
				AddTxt="[list]";
			} else {
				AddTxt="[list="+txt+"]";
			}
			AddText(AddTxt); 
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("列表項目\n空白或取消表示結束列表",""); 
				if (txt!="" && txt!=null) {             
					AddTxt="[*]"+txt;
					AddText(AddTxt);
				}
			} 
			AddTxt="[/list]";
			AddText(AddTxt); 
		}
	}
}

function olist() {
	if (helpstat) {
		alert("編號清單列表標記\n產生一個有編號的清單列表.\n\n用法: [olist=標記類型( A | a | I | i | 1 )] [*]項目一 [*]項目二 [*]項目三 [/olist]");
	} else if (basic) {
		AddTxt="[olist][*]  [*]  [*]  [/olist]";
		AddText(AddTxt);
	} else {  
		txt=prompt("輸入編號列表類型：'A'為大寫字母,'a'為小寫字母,'I'為大寫羅馬字母,\n'i'為小寫羅馬字母,'1'為數字,空白為數字.","");
		//if (txt==null){ return false; }               
		while ((txt!="") && (txt.toLowerCase()!="a")  && (txt.toLowerCase()!="i") && (txt!="1") && (txt!=null)) {
			txt=prompt("錯誤!\n類型只能輸入 'A','a','I','i','1' 或者空白.","");               
		}
		if (txt!=null) {
			if (txt=="") {
				AddTxt="[olist]";
			} else {
				AddTxt="[olist="+txt+"]";
			}
			AddText(AddTxt);
			txt="1";
			while ((txt!="") && (txt!=null)) {
				txt=prompt("編號列表項目\n空白或取消表示結束列表",""); 
				if (txt!="" && txt!=null) {             
					AddTxt="[*]"+txt;
					AddText(AddTxt);
				}
			} 
			AddTxt="[/olist]";
			AddText(AddTxt); 
		}
	}
}

function dlist() {
	if (helpstat) {
		alert("定義清單列表標記\n產生一個定義清單列表.\n\n用法: [dlist=定義名詞]定義名詞解釋[/dlist]");
	} else if (basic) {
		AddTxt="[dlist=][/dlist]";
		AddText(AddTxt);
	} else {  
		txt=prompt("請輸入定義名詞：","");
		if (txt!=null) {
			AddTxt="[dlist="+txt+"]";
			AddText(AddTxt); 
			txt=prompt("請輸入定義名詞的解釋：",""); 
			if (txt!="" && txt!=null) {             
				AddTxt=txt;
				AddText(AddTxt);
			}
			AddTxt="[/dlist]";
			AddText(AddTxt); 
		}
	}
}

function setfont(font) {
 	if (helpstat){
		alert("設定字體標記\n給文字設置字體.\n用法: [face="+font+"]改變文字字體為"+font+"[/face]");
	} else if (basic) {
		AddTxt="[face="+font+"][/face]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("請輸入要設置"+font+"字體的文字：","文字");
		if (txt!=null) {             
			AddTxt="[face="+font+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/face]";
			AddText(AddTxt);
		}        
	}  
}
function underline() {
  	if (helpstat) {
		alert("底線標記\n使文字變成底線文字.\n用法: [u]要加底線的文字[/u]");
	} else if (basic) {
		AddTxt="[u][/u]";
		AddText(AddTxt);
	} else {  
		txt=prompt("底線文字.","文字");     
		if (txt!=null) {           
			AddTxt="[u]"+txt;
			AddText(AddTxt);
			AddTxt="[/u]";
			AddText(AddTxt);
		}	        
	}
}
function fly() {
 	if (helpstat){
		alert("飛翔標記\n使文字飛行.\n用法: [fly]要產生飛行效果的文字[/fly]");
	} else if (basic) {
		AddTxt="[fly][/fly]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("飛翔文字","文字");
		if (txt!=null) {             
			AddTxt="[fly]"+txt;
			AddText(AddTxt);
			AddTxt="[/fly]";
			AddText(AddTxt);
		}        
	}  
}

function move() {
	if (helpstat) {
		alert("移動標記\n使文字產生移動效果.\n用法: [move]要產生移動效果的文字[/move]");
	} else if (basic) {
		AddTxt="[move][/move]";
		AddText(AddTxt);
	} else {  
		txt=prompt("要產生移動效果的文字","文字");     
		if (txt!=null) {           
			AddTxt="[move]"+txt;
			AddText(AddTxt);
			AddTxt="[/move]";
			AddText(AddTxt);
		}       
	}
}

function shadow() {
	if (helpstat) {
               alert("陰影標記\n使文字產生陰影效果.\n用法: [SHADOW=寬度, 顏色, 邊界]要產生陰影效果的文字[/SHADOW]");
	} else if (basic) {
		AddTxt="[SHADOW=255,blue,1][/SHADOW]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("文字的長度、顏色和邊界大小","255,blue,1"); 
		if (txt2!=null) {
			txt=prompt("要產生陰影效果的文字","文字");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[SHADOW=255, blue, 1]"+txt;
					AddText(AddTxt);
					AddTxt="[/SHADOW]";
					AddText(AddTxt);
				} else {
					AddTxt="[SHADOW="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/SHADOW]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function glow() {
	if (helpstat) {
		alert("光暈標記\n使文字產生光暈效果.\n用法: [GLOW=寬度, 顏色, 邊界]要產生光暈效果的文字[/GLOW]");
	} else if (basic) {
		AddTxt="[glow=255,red,2][/glow]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("文字的長度、顏色和邊界大小","255,red,2"); 
		if (txt2!=null) {
			txt=prompt("要產生光暈效果的文字.","文字");      
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[glow=255,red,2]"+txt;
					AddText(AddTxt);
					AddTxt="[/glow]";
					AddText(AddTxt);
				} else {
					AddTxt="[glow="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/glow]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}
function hr(){
	if (helpstat){
		alert("水平線\n插入水平線.\n用法： [hr=顏色,寬度]");
	}else if (basic){
		AddTxt = "[hr=red,70%]";
		AddText(AddTxt);
	}else{
		txt2 = prompt("請輸入水平線的顏色\(可用#XXXXXX或是英文顏色名\)：","red");
		if (txt2!=null){
			txt = prompt("請輸入水平線的寬度\(可用數字或XX%\)：","70%");
			if (txt!=null){
				if (txt2==""){
					AddTxt = "[hr=red,"+txt+"]";
					AddText(AddTxt);
				}else{
					AddTxt = "[hr="+txt2+","+txt+"]";
					AddText(AddTxt);
				}
			}
		}
	}
}
function hx(){
	if (helpstat){
		alert("標題文字\n使文字變成標題文字.\n用法： [h=大小(1~6)]要變成標題的文字[/h]");
	}else if (basic){
		AddTxt = "[h=1][/h]";
		AddText(AddTxt);
	}else{
		txt2 = prompt("請輸入標題文字的大小\(最大 1 最小 6 範圍 1-6\)：","1");
		if (txt2!=null){
			txt = prompt("請輸入欲變成標題文字的文字：","文字");
			if (txt!=null){
				if (txt2==""){
					AddTxt = "[h=1]"+txt;
					AddText(AddTxt);
					AddTxt = "[/h]";
					AddText(AddTxt);
					
				}else{
					AddTxt = "[h="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt = "[/h]";
					AddText(AddTxt);
				}
			}
		}
	}
}
function up(){
	if (helpstat) {
		alert("上標字\n使文字變為上標字.\n用法: [up]這是上標字[/up]");
	} else if (basic) {
		AddTxt="[up][/up]";
		AddText(AddTxt);
	} else {   
		txt=prompt("請輸入欲變成上標字的文字：","文字");     
		if (txt!=null) {           
			AddTxt="[up]"+txt;
			AddText(AddTxt);
			AddTxt="[/up]";
			AddText(AddTxt);
		}	        
	}
}
function dw(){
	if (helpstat) {
		alert("下標字\n使文字變為下標字.\n用法: [dw]這是下標字[/dw]");
	} else if (basic) {
		AddTxt="[dw][/dw]";
		AddText(AddTxt);
	} else {   
		txt=prompt("請輸入欲變成下標字的文字：","文字");     
		if (txt!=null) {           
			AddTxt="[dw]"+txt;
			AddText(AddTxt);
			AddTxt="[/dw]";
			AddText(AddTxt);
		}	        
	}
}
function strike(){
	if (helpstat) {
		alert("刪除字\n使文字變為刪除字.\n用法: [del]這是下標字[/del]");
	} else if (basic) {
		AddTxt="[del][/del]";
		AddText(AddTxt);
	} else {   
		txt=prompt("請輸入欲變成刪除字的文字：","文字");     
		if (txt!=null) {           
			AddTxt="[del]"+txt;
			AddText(AddTxt);
			AddTxt="[/del]";
			AddText(AddTxt);
		}	        
	}
}

function bigsml(bos){
	var Tag = (bos=="big") ? ("放大") : ("縮小");
	if (helpstat) {
		alert(Tag+"字\n使文字變為"+Tag+"字.\n用法: ["+bos+"]這是"+Tag+"字[/"+bos+"]");
	}else if (basic) {
		AddTxt = "["+bos+"][/"+bos+"]";
		AddText(AddTxt);
	}else {
		txt = prompt("請輸入欲變為"+Tag+"字的文字：","文字");
		if (txt!=null){
			AddTxt = "["+bos+"]" + txt;
			AddText(AddTxt);
			AddTxt = "[/"+bos+"]";
			AddText(AddTxt);
		}
	}
}