
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
		alert("Email �s���аO\n���J Email �W�s��\n�Ϊk1: [email]nobody@domain.com[/email]\n�Ϊk2: [email=nobody@domain.com]�H�W[/email]");
	} else if (basic) {
		AddTxt="[email][/email]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("�s����ܪ���r.\n�p�G�ťաA����N�u��ܳs���� Email �a�}",""); 
		if (txt2!=null) {
			txt=prompt("Email ��}.","name@domain.com");      
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
		alert("Flash �ʵe\n���J Flash �ʵe.\n�Ϊk: [flash=�e��,����]Flash �ʵe��󪺦�}[/flash]");
	} else if (basic) {
		AddTxt="[flash][/flash]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("�п�JFlash ��󪺦�}�G","http://");
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
		alert("�峹����\n���J ���ûy�k.\n�Ϊk: [hide]���e[/hide]");
	} else if (basic) {
		AddTxt="[hide][/hide]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("�п�J���ä�󪺤��e�G","���e");
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
		alert("Shockwave �ʵe\n���J Shockwave �ʵe.\n�Ϊk: [dir=�e��,����]Shockwave �ʵe��󪺦�}[/dir]");
	} else if (basic) {
		AddTxt="[dir][/dir]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("�п�JShockwave �ʵe��󪺦�}�G","http://");
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
		alert("RealPlayer ���T�ɮ�\n���J RealPlayer *.rm ���T�ɮ�.\n�Ϊk: [rm=�e��,����]RealPlayer ���T�ɪ���}[/rm]");
	} else if (basic) {
		AddTxt="[rm][/rm]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("�п�JRealPlayer ���T�ɮת���}�G","http://");
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
		alert("Windows Media Player ���T�ɮ�\n���J Windows Media Player *.wmv ���T�ɮ�.\n�Ϊk: [mp=�e��,����]Windows Media Player ���T�ɮת���}[/mp]");
	} else if (basic) {
		AddTxt="[mp][/mp]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("�п�JWindows Media Player ���T�ɮת���}�G","http://");
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
		alert("QuickTime ���T�ɮ�\n���J QuickTime *.mov ���T��.\n�Ϊk: [qt=�e��,����]QuickTime ���T�ɮת���}[/qt]");
	} else if (basic) {
		AddTxt="[qt][/qt]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("�п�JQuickTime ���T�ɪ���}�G","http://");
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
		alert("��r�j�p�аO\n�]�m��r�j�p.\n�d�� 1 - 4.\n 1 ���̤p 4 ���̤j.\n�Ϊk: [size="+size+"]�o�O "+size+" ��r[/size]");
	} else if (basic) {
		AddTxt="[size="+size+"][/size]";
		AddText(AddTxt);
	} else {                       
		txt=prompt("�r��j�p:"+size+"\n�п�J��r�G","��r"); 
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
		alert("����аO\n�Ϥ�r�ܬ�����.\n�Ϊk: [b]�o�O�����r[/b]");
	} else if (basic) {
		AddTxt="[b][/b]";
		AddText(AddTxt);
	} else {  
		txt=prompt("�п�J���ܦ����骺��r�G","��r");     
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
		alert("����аO\n�Ϥ�r�ܬ�����.\n�Ϊk: [i]�o�O����r[/i]");
	} else if (basic) {
		AddTxt="[i][/i]";
		AddText(AddTxt);
	} else {   
		txt=prompt("�п�J���ܦ����骺��r�G","��r");     
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
		alert("�ޥμаO\n�ޥΤ@�Ǥ�r.\n�Ϊk: [quote]�ޥΤ��e[/quote]");
	} else if (basic) {
		AddTxt="[quote][/quote]";
		AddText(AddTxt);
	} else {   
		txt=prompt("�Q�ޥΪ���r","��r");     
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
		alert("�C��аO\n�]�m��r�C��.  �����C�ⳣ�i�H�Q�ϥ�.\n�Ϊk: [color="+color+"]�C��n���ܬ�"+color+"����r[/color]");
	} else if (basic) {
		AddTxt="[color="+color+"][/color]";
		AddText(AddTxt);
	} else {  
     	txt=prompt("��ܪ��C��O: "+color+"\n�п�J��r�G","��r");
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
		alert("����аO\n�ϥγo�ӼаO, �i�H�Ϥ�r�a������B�m���B�a�k���.\n�Ϊk: [align= center | left | right ]�n�������r[/align]");
	} else if (basic) {
		AddTxt="[align=left][/align]";
		AddText(AddTxt);
	} else {  
		txt2=prompt("����˦�\n��J 'center' ��ܸm��, 'left' ��ܾa�����, 'right' ��ܾa�k���.","center");               
		while ((txt2!="") && (txt2!="center") && (txt2!="left") && (txt2!="right") && (txt2!=null)) {
			txt2=prompt("���~!\n�����u���J 'center' �B 'left' �Ϊ� 'right'.","");               
		}
		txt=prompt("�п�J�n�������r�G","��r");     
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
		alert("�W�s���аO\n���J�@�ӶW�s���аO\n��k: [url]http://domain.com[/url]\n�άO: [url=http://domain.com]�s����r[/url]");
	} else if (basic) {
		AddTxt="[url][/url]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("�W�s����ܤ�r.\n�p�G���Q�ϥ�, �i�H�ť�, �N�u��ܶW�s����}. ",""); 
		if (txt2!=null) {
			txt=prompt("�п�J�W�s����m�G","http://");      
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
		alert("�Ϥ��аO\n���J�Ϥ�\n�Ϊk�G [img]http://domain.com/test.gif[/img]");
	} else if (basic) {
		AddTxt="[img][/img]";
		AddText(AddTxt);
	} else {  
		txt=prompt("�п�J�Ϥ��� URL��m�G","http://");    
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
		alert("�{���X�аO\n�ϥε{���X�аO�A�i�H�ϧA���{���X�̭��� html ���аO���|�Q�}�a.\n��k:\n [code]�o�̬O�{���X��r[/code]");
	} else if (basic) {
		AddTxt="[code][/code]";
		AddText(AddTxt);
	} else {   
		txt=prompt("��J�N�X","");     
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
		alert("�M��C��аO\n���ͤ@�ӲM��C��.\n\n�Ϊk: [list=�аO����( square | circle )] [*]���ؤ@ [*]���ؤG [*]���ؤT [/list]");
	} else if (basic) {
		AddTxt="[list][*]  [*]  [*]  [/list]";
		AddText(AddTxt);
	} else {  
		txt=prompt("��J�C������\n'square'�����,'circle'���ն��,�ťլ��¶��.","");
		//if (txt==null){ return false; }               
		while ((txt!="") && (txt.toLowerCase()!="square") && (txt.toLowerCase()!="circle") && (txt!=null)) {
			txt=prompt("���~!\n�����u���J 'square' �B 'circle' �Ϊ̪ť�.","");               
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
				txt=prompt("�C����\n�ťթΨ�����ܵ����C��",""); 
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
		alert("�s���M��C��аO\n���ͤ@�Ӧ��s�����M��C��.\n\n�Ϊk: [olist=�аO����( A | a | I | i | 1 )] [*]���ؤ@ [*]���ؤG [*]���ؤT [/olist]");
	} else if (basic) {
		AddTxt="[olist][*]  [*]  [*]  [/olist]";
		AddText(AddTxt);
	} else {  
		txt=prompt("��J�s���C�������G'A'���j�g�r��,'a'���p�g�r��,'I'���j�gù���r��,\n'i'���p�gù���r��,'1'���Ʀr,�ťլ��Ʀr.","");
		//if (txt==null){ return false; }               
		while ((txt!="") && (txt.toLowerCase()!="a")  && (txt.toLowerCase()!="i") && (txt!="1") && (txt!=null)) {
			txt=prompt("���~!\n�����u���J 'A','a','I','i','1' �Ϊ̪ť�.","");               
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
				txt=prompt("�s���C����\n�ťթΨ�����ܵ����C��",""); 
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
		alert("�w�q�M��C��аO\n���ͤ@�өw�q�M��C��.\n\n�Ϊk: [dlist=�w�q�W��]�w�q�W������[/dlist]");
	} else if (basic) {
		AddTxt="[dlist=][/dlist]";
		AddText(AddTxt);
	} else {  
		txt=prompt("�п�J�w�q�W���G","");
		if (txt!=null) {
			AddTxt="[dlist="+txt+"]";
			AddText(AddTxt); 
			txt=prompt("�п�J�w�q�W���������G",""); 
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
		alert("�]�w�r��аO\n����r�]�m�r��.\n�Ϊk: [face="+font+"]���ܤ�r�r�鬰"+font+"[/face]");
	} else if (basic) {
		AddTxt="[face="+font+"][/face]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("�п�J�n�]�m"+font+"�r�骺��r�G","��r");
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
		alert("���u�аO\n�Ϥ�r�ܦ����u��r.\n�Ϊk: [u]�n�[���u����r[/u]");
	} else if (basic) {
		AddTxt="[u][/u]";
		AddText(AddTxt);
	} else {  
		txt=prompt("���u��r.","��r");     
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
		alert("�����аO\n�Ϥ�r����.\n�Ϊk: [fly]�n���ͭ���ĪG����r[/fly]");
	} else if (basic) {
		AddTxt="[fly][/fly]";
		AddText(AddTxt);
	} else {                  
		txt=prompt("������r","��r");
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
		alert("���ʼаO\n�Ϥ�r���Ͳ��ʮĪG.\n�Ϊk: [move]�n���Ͳ��ʮĪG����r[/move]");
	} else if (basic) {
		AddTxt="[move][/move]";
		AddText(AddTxt);
	} else {  
		txt=prompt("�n���Ͳ��ʮĪG����r","��r");     
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
               alert("���v�аO\n�Ϥ�r���ͳ��v�ĪG.\n�Ϊk: [SHADOW=�e��, �C��, ���]�n���ͳ��v�ĪG����r[/SHADOW]");
	} else if (basic) {
		AddTxt="[SHADOW=255,blue,1][/SHADOW]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��r�����סB�C��M��ɤj�p","255,blue,1"); 
		if (txt2!=null) {
			txt=prompt("�n���ͳ��v�ĪG����r","��r");
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
		alert("���w�аO\n�Ϥ�r���ͥ��w�ĪG.\n�Ϊk: [GLOW=�e��, �C��, ���]�n���ͥ��w�ĪG����r[/GLOW]");
	} else if (basic) {
		AddTxt="[glow=255,red,2][/glow]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("��r�����סB�C��M��ɤj�p","255,red,2"); 
		if (txt2!=null) {
			txt=prompt("�n���ͥ��w�ĪG����r.","��r");      
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
		alert("�����u\n���J�����u.\n�Ϊk�G [hr=�C��,�e��]");
	}else if (basic){
		AddTxt = "[hr=red,70%]";
		AddText(AddTxt);
	}else{
		txt2 = prompt("�п�J�����u���C��\(�i��#XXXXXX�άO�^���C��W\)�G","red");
		if (txt2!=null){
			txt = prompt("�п�J�����u���e��\(�i�μƦr��XX%\)�G","70%");
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
		alert("���D��r\n�Ϥ�r�ܦ����D��r.\n�Ϊk�G [h=�j�p(1~6)]�n�ܦ����D����r[/h]");
	}else if (basic){
		AddTxt = "[h=1][/h]";
		AddText(AddTxt);
	}else{
		txt2 = prompt("�п�J���D��r���j�p\(�̤j 1 �̤p 6 �d�� 1-6\)�G","1");
		if (txt2!=null){
			txt = prompt("�п�J���ܦ����D��r����r�G","��r");
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
		alert("�W�Цr\n�Ϥ�r�ܬ��W�Цr.\n�Ϊk: [up]�o�O�W�Цr[/up]");
	} else if (basic) {
		AddTxt="[up][/up]";
		AddText(AddTxt);
	} else {   
		txt=prompt("�п�J���ܦ��W�Цr����r�G","��r");     
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
		alert("�U�Цr\n�Ϥ�r�ܬ��U�Цr.\n�Ϊk: [dw]�o�O�U�Цr[/dw]");
	} else if (basic) {
		AddTxt="[dw][/dw]";
		AddText(AddTxt);
	} else {   
		txt=prompt("�п�J���ܦ��U�Цr����r�G","��r");     
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
		alert("�R���r\n�Ϥ�r�ܬ��R���r.\n�Ϊk: [del]�o�O�U�Цr[/del]");
	} else if (basic) {
		AddTxt="[del][/del]";
		AddText(AddTxt);
	} else {   
		txt=prompt("�п�J���ܦ��R���r����r�G","��r");     
		if (txt!=null) {           
			AddTxt="[del]"+txt;
			AddText(AddTxt);
			AddTxt="[/del]";
			AddText(AddTxt);
		}	        
	}
}

function bigsml(bos){
	var Tag = (bos=="big") ? ("��j") : ("�Y�p");
	if (helpstat) {
		alert(Tag+"�r\n�Ϥ�r�ܬ�"+Tag+"�r.\n�Ϊk: ["+bos+"]�o�O"+Tag+"�r[/"+bos+"]");
	}else if (basic) {
		AddTxt = "["+bos+"][/"+bos+"]";
		AddText(AddTxt);
	}else {
		txt = prompt("�п�J���ܬ�"+Tag+"�r����r�G","��r");
		if (txt!=null){
			AddTxt = "["+bos+"]" + txt;
			AddText(AddTxt);
			AddTxt = "[/"+bos+"]";
			AddText(AddTxt);
		}
	}
}