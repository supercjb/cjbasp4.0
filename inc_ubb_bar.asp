<script language="JavaScript">
var helpstat = false;
var stprompt = true;
var basic = false;
</script>
<script language="JavaScript" src="inc_ubb_fun.js">
</script>
<font size="2" color="#000080">模式選擇： </font><font color="#000080">                
<select name="Choose" size="1" onChange="thelp(this.options[this.selectedIndex].value)">                 
<option value="3" selected>插入精靈</option>                  
<option value="1">說明標記</option>                 
<option value="2">只插入標記</option>                 
</select></font><font size="2" color="#000080">                
<br>                
<img onclick=bold() align=absmiddle src="images/toolbar/editor_bold.gif" width="22" height="22" alt="粗體" border="0">                
<img onclick=italicize() align=absmiddle src="images/toolbar/editor_italicize.gif" width="23" height="22" alt="斜體" border="0">                     
<img onclick=underline() align=absmiddle src="images/toolbar/editor_underline.gif" width="23" height="22" alt="底線" border="0">                     
<img onclick=center() align=absmiddle src="images/toolbar/editor_center.gif" width="22" height="22" alt="對齊" border="0">                     
<img onclick=hyperlink() align=absmiddle src="images/toolbar/editor_url.gif" width="22" height="22" alt="超連結" border="0">                     
<img onclick=email() align=absmiddle src="images/toolbar/editor_email.gif" width="23" height="22" alt="Email連結" border="0">                     
<img onclick=image() align=absmiddle src="images/toolbar/editor_image.gif" width="23" height="22" alt="圖片" border="0">                     
<img onclick=flash() align=absmiddle src="images/toolbar/editor_flash.gif" width="23" height="22" alt="Flash檔案" border="0">                     
<img onclick=director() align=absmiddle src="images/toolbar/editor_dir.gif" width="23" height="22" alt="Shockwave檔案" border="0">                     
<img onclick=rm() align=absmiddle src="images/toolbar/editor_rm.gif" width="23" height="22" alt="Real Player視訊檔案" border="0">                     
<img onclick=wmv() align=absmiddle src="images/toolbar/editor_mp.gif" width="23" height="22" alt="Media Player視訊檔案" border="0">                    
<img onclick=mov() align=absmiddle src="images/toolbar/editor_qt.gif" width="23" height="22" alt="QuickTime視訊檔案" border="0">                     
<img onclick=code() align=absmiddle src="images/toolbar/editor_code.gif" width="22" height="22" alt="程式碼" border="0">                     
               
<img onclick=hr() src="images/toolbar/editor_hr.gif" border="0" width="23" height="22" align="absmiddle" alt="水平線">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;               
<img onclick=list() align=absmiddle src="images/toolbar/editor_list.gif" width="23" height="22" alt="清單列表" border="0">                     
<img onclick=olist() border="0" src="images/toolbar/editor_olist.gif" align="absmiddle" width="23" height="22" alt="編號清單列表">   
<img onclick=dlist() border="0" src="images/toolbar/editor_dlist.gif" align="absmiddle" width="23" height="22" alt="定義清單列表">      
<img onclick=fly() align=absmiddle height=22 alt=飛行字 src="images/toolbar/editor_fly.gif" width=23 border="0">                    
<img onclick=move() align=absmiddle height=22 alt=移動字 src="images/toolbar/editor_move.gif" width=23 border="0">                     
<img onclick=glow() align=absmiddle height=22 alt=發光字 src="images/toolbar/editor_glow.gif" width=23 border="0">                     
<img onclick=shadow() align=absmiddle height=22 alt=陰影字 src="images/toolbar/editor_shadow.gif" width=23 border="0">
<img border="0" onclick=hide() align=absmiddle height=22 alt=隱藏文字 src="images/toolbar/hide.gif" align="middle">&nbsp;  
<br>  
字體大小： </font><font color="#000080">                     
<select name="size" onChange="setsize(this.options[this.selectedIndex].value)" size="1">                     
<option value="1">1</option>                    
<option value="2" selected>2</option>                    
<option value="3">3</option>                    
<option value="4">4</option>                    
</select></font><font size="2" color="#000080">                    
顏色： </font><font color="#000080">                     
<SELECT onChange="setcolor(this.options[this.selectedIndex].value)" name=color size="1">                      
<option style="background-color:#F0F8FF;color: #F0F8FF" value="#F0F8FF">#F0F8FF</option>                              
<option style="background-color:#DEB887;color: #DEB887" value="#DEB887">#DEB887</option>                    
<option style="background-color:#5F9EA0;color: #5F9EA0" value="#5F9EA0">#5F9EA0</option>                    
<option style="background-color:#7FFF00;color: #7FFF00" value="#7FFF00">#7FFF00</option>                    
<option style="background-color:#D2691E;color: #D2691E" value="#D2691E">#D2691E</option>                    
<option style="background-color:#FF7F50;color: #FF7F50" value="#FF7F50">#FF7F50</option>                    
<option style="background-color:#6495ED;color: #6495ED" value="#6495ED" selected>#6495ED</option>                    
<option style="background-color:#FFF8DC;color: #FFF8DC" value="#FFF8DC">#FFF8DC</option>                    
<option style="background-color:#DC143C;color: #DC143C" value="#DC143C">#DC143C</option>
<option style="background-color:#9400D3;color: #9400D3" value="#9400D3">#9400D3</option>                             
<option style="background-color:#FDF5E6;color: #FDF5E6" value="#FDF5E6">#FDF5E6</option>                    
<option style="background-color:#808000;color: #808000" value="#808000">#808000</option>                        
<option style="background-color:#2E8B57;color: #2E8B57" value="#2E8B57">#2E8B57</option>                    
<option style="background-color:#87CEEB;color: #87CEEB" value="#87CEEB">#87CEEB</option>                    
<option style="background-color:#6A5ACD;color: #6A5ACD" value="#6A5ACD">#6A5ACD</option>  
</SELECT></font><font size="2" color="#000080">
<br>        
字體選擇：</font>     
<select onChange="setfont(this.options[this.selectedIndex].value)" size="1" name="font">     
<option value="新細明體">新細明體</option>   
<option value="細明體">細明體</option>   
<option value="標楷體">標楷體</option>   
<option value="Times New Roman">Times New Roman</option>   
<option value="Arial">Arial</option>   
<option value="Wingdings">Wingdings</option>   
</select>