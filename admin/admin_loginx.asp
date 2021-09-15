<HTML><HEAD><TITLE>后台登陆</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="images/admin.css" type=text/css rel=stylesheet>
<script>
var travel=true

var hotkey=<%=session("mypasskey")%>
if (document.layers)
document.captureEvents(Event.KEYPRESS)
function mypass(e){
if(document.layers){
if(e.which==hotkey&&travel)
mypass.style.display=""
}
else if (document.all){
if(event.keyCode==hotkey)
document.all.mypass.style.display=""
}
}
document.onkeypress=mypass
function login(){
if (document.form1.AdminName.value==""){alert("请输入用户名？");document.form1.AdminName.focus();return false}
if (document.form1.password.value==""){alert("请输入密码？");document.form1.password.focus();return false} 
if (document.form1.verifycode.value==""){alert("请输入验证码？");document.form1.verifycode.focus();return false}
return true}
</script>

</HEAD>

<body onkeydown=return(!(event.keyCode==78&&event.ctrlKey)) scroll=no>
<script language="JavaScript" src="keyboard.js" type="text/javascript"></script>
<br><br><br>


<table border="0" cellpadding="0" cellspacing="0" width="642" background="images/adminlogin.jpg" height="425" align="center">
  <tr>
    <td width="100%" height="370">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" height="401">
        <tr>
          <td width="50%" height="401"></td>
          <td width="50%" height="401">
            <table border="0" cellpadding="0" cellspacing="0" width="100%">
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
              <tr>
                <td width="100%">　
                  <p>　</p>
                  <p>　</p>
                  <p>　</p>
                  <p>　</p>
                  <p>　</td>
              </tr>
              <tr>
               <TD height=25 align=center><form name="Login" method="post" action="admin_klogin.asp?action=login" onsubmit="return login()">
用户名：<input type="text" name="AdminName"  style=width:150px autocomplete="off"  class="ycenchi"><input onclick="showkeyboard('Login.AdminName')" type="button" value="键盘" title="用软键盘输入密码，防止黑客软件记录键盘的录入信息" />
</td></tr>
<tr><td height=25 align=center>
密&nbsp; 码：<input type="password" name="Password"  style=width:150px  class="ycenchi"><input onclick="showkeyboard('Login.Password')" type="button" value="键盘" title="用软键盘输入密码，防止黑客软件记录键盘的录入信息" /></td></tr>   
<tr  id=mypass style="display:none"><td height=25 align=center>
二次码：<input type="password" name="mypassword"  style=width:150px  class="ycenchi" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</td></tr>
<tr><td height=25 align=center>
认证码：<input name="verifycode" type="text" size="12" maxlength="9" class="ycenchi">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="../inc/getcode.asp"  id="GetCodePic" align=absmiddle height=16 border=0></td></tr><tr><td height=25 align=center>
<input type="submit" name="Submit" value="登 录" class="adminbutton">&nbsp;<input type="reset" name="reset" value="清 空" class="adminbutton">
</TD>
              </tr>
              <tr>
                <td width="100%"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>



</body></html>














































