<!--#include file="conn.asp"-->
<HTML>
<HEAD><link rel="stylesheet" type="text/css" href=style1.css>
<TITLE><%=rsidid("bbsname")%></TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 6.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<meta http-equiv="Content-Type" content="text/html"></HEAD>
<BODY leftmargin="0" topmargin="0">
<!--#include file="header.asp"-->
<% response.cookies("tid")="�s�|�����U"%>
<!--#include file="choice.asp"-->
<center>
<%response.write "<br><font color=red>�{�b�ɨ�G"&DateAdd("h",rsidid("servertimezone"),now)&"�A�Y��ܮɶ��P�z�Ҧb�a���ɮt�A�Щ�ɮt�վ����J�t��</font>"
response.write "<br>�Ҧp��ܲ{�b�ɨ謰 10/10/2003 9:36:36 PM�A�ӱz�Ҧb�a�ثe�ɨ謰 10/11/2003 2:36:36 AM�A�h��J+5"
response.write "<br>�Ҧp��ܲ{�b�ɨ謰 10/10/2003 9:36:36 PM�A�ӱz�Ҧb�a�ثe�ɨ謰 10/10/2003 2:36:36 PM�A�h��J-7"%>
</center>
<FORM METHOD=POST ACTION="regist.asp" name="form">
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="558" height="1" id="AutoNumber1">
    <tr>
      <td width="556" height="19" bgcolor="#F0EEDF" bordercolor="#FFFFFF" background="images/Button.gif" colspan="2">
      <p align="center"><b><font size="3">��g�Τ�ӤH���</font></b><font size="2" color=red>�]�������񶵥ء^</font></td>
    </tr>
    <tr>
      <td width="160" height="13" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#FF0000"><b>��</b></font><font size="2" color="#800000">�b���G</font></td>
      <td width="400" height="13" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="UserID"  size="26" ></font><font size=-1 color=red>(�Шϥέ^��B�Ʀr�B�c�餤��)</font></td>
    </tr>
    <tr>
      <td width="160" height="11" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#FF0000"><b>��</b></font><font size="2" color="#800000" >�K�X�G</font></td>
      <td width="400" height="11" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="Password" NAME="Password1" size="26" ></font></td>
    </tr>
    <tr>
      <td width="160" height="17" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#FF0000"><b>��</b></font><font size="2" color="#800000">�ЦA��J�@���G</font></td>
      <td width="400" height="17" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="Password" NAME="Password2" size="26"></font></td>
    </tr>
    <tr>
      <td width="160" height="9" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#FF0000"><b>��</b></font><font size="2" color="#800000">�ͤ�G</font></td>
      <td width="400" height="1" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="birthday" size="26" ></font><font size=-1 color=red>(�榡�G1984/01/08)</font></td>
    </tr>
    <tr>
      <td width="160" height="9" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#FF0000"><b>��</b></font><font size="2" color="#800000">�ӦۡG</font></td>
      <td width="400" height="1" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="comefrom" size="26" ></font><font size=-1 color=red>(�Ҧp�G��򰨤�)</font></td>
    </tr>    
    <tr>
      <td width="160" height="8" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#FF0000"><b>��</b></font><font size="2" color="#800000">E-MAIL�G</font></td>
      <td width="400" height="8" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="email" size="26" ></font></td>
    </tr>
    <tr>
      <td width="160" height="1" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#800000">�ӤH�����G</font></td>
      <td width="400" height="9" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="homepage" size="26" ></font></td>
    </tr>
    <tr>
      <td width="160" height="1" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#800000">MSN Messenger�G</font></td>
      <td width="400" height="9" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="msnm" size="26" ></font></td>
    </tr>
    <tr>
      <td width="160" height="1" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#800000">Yahoo!�_�� Messenger�G</font></td>
      <td width="400" height="9" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="yim" size="26" ></font><font size=-1 color=red>(�u��b���Y�i�A�Ҧp�Gsundayhi)</font></td>
    </tr>    
    <tr>
      <td width="160" height="1" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="right">
      <font size="2" color="#800000">�ɮt�վ�G</font></p></td>
      <td width="400" height="1" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <font color="#800000"><INPUT TYPE="text" NAME="timezone" size="26" VALUE=0></font><font size=-1 color=red>�ЬݤW���������վ�</font></td>
    </tr>
    <tr>
      <td width="100%" colspan="2" height="51" bgcolor="#EFF4FE" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
      <p align="left"><font size="2">�Y ��:</font>
      <img src="face/1.gif" name="tus" width="100" height="100">
<script>function showimage(){document.images.tus.src="face/"+document.form.pic.options[document.form.pic.selectedIndex].value+".gif";}</script>
<select name=pic size=1 onChange="showimage()">

<option value=1>1</option>
<option value=2>2</option>
<option value=3>3</option>
<option value=4>4</option>
<option value=5>5</option>
<option value=6>6</option>
<option value=7>7</option>
<option value=8>8</option>
<option value=9>9</option>
<option value=10>10</option>
<option value=11>11</option>
<option value=12>12</option>
<option value=13>13</option>
<option value=14>14</option>
<option value=15>15</option>
<option value=16>16</option>
<option value=17>17</option>
<option value=18>18</option>
<option value=19>19</option>
<option value=20>20</option>
<option value=21>21</option>
<option value=22>22</option>
<option value=23>23</option>
<option value=24>24</option>
<option value=25>25</option>
<option value=26>26</option>
<option value=27>27</option>
<option value=28>28</option>
<option value=29>29</option>
<option value=30>30</option>
<option value=31>31</option>
<option value=32>32</option>
<option value=33>33</option>
<option value=34>34</option>
<option value=35>35</option>
<option value=36>36</option>
<option value=37>37</option>
<option value=38>38</option>
<option value=0>�ۭq</option>
	<select>
      <font size="2">�Y�ۭq�п�J������m:</font><input type="text" name="face_url" size="30" > 

      </select></td>
    </tr>
  <tr>
    <td width="100%" colspan="4" height="19" bgcolor="#F0EEDF" bordercolor="#FFFFFF" background="images/Button.gif">
    <p align="center"><font color="#800000"><b>�ӤH�T��</b></font></p> 
    </td>
  </tr>
  <tr>
    <td width="100%" colspan="4" height="19" bgcolor="#F0EEDF" bordercolor="#FFFFFF" background="images/postbit_lightbg.gif">
    <p align="center"><font size="2">ñ�W��:</font><textarea rows="9" name="memo1" cols="50"></textarea></p>   
    </td>
  </tr>    
    <tr>
      <td width="556" height="18" colspan="2">
<p align="center">
<INPUT TYPE="submit" VALUE="�T�w"><INPUT TYPE="reset" VALUE="���s"></td>
    </tr>
  </table>
  </center></p>
</div>
</FORM>  
<!--#include file="foot.asp"-->
</BODY></HTML>