<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
'=====================================================================
' 软件名称：恩池网站管理系统
' 当前版本：enchicms Version 3.0.0
' 更新日期：2005-03-25
' 官方网站：运城市恩池软件科技开发有限公司(www.enchi.com.cn) 
' 技术支持：柳云帆
' 邮箱：liuyunfan@163.com
' QQ：21556923
'=====================================================================
' Copyright 2005-2008  All Rights Reserved.
'=====================================================================
If Not ChkAdmin("gundong") Then
	'Server.Transfer("showerr.asp")
	'Response.End
End If

dim Action
dim strDir,strAdminDir
strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
strDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
Action=Trim(request("Action"))
%>
 <table cellspacing=1 align=center cellpadding=3 border=0 class=tableborder>
    <tr> 
     <td height="22" colspan=2 align=center class=tablerow2><a name="Top"></a><strong>图片左右滚动管理</strong></td>
    </tr>
    <tr>
    <td class=tablerow2> 
    调用方法：在需要调用的地方加载如下标签：{$dibuhuan},如果要修改该FLASH的图片大小等参数请在通栏模版基本设置中修改。最多可设置10张图片，建议使用JPG图片
	</td>
    </table>
<br />
<script language = JavaScript>
function SelectPhoto1(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic1.value=ss[0];
  }
}

function SelectPhoto2(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic2.value=ss[0];
  }
}

function SelectPhoto3(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic3.value=ss[0];
  }
}


function SelectPhoto4(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic4.value=ss[0];
  }
}

function SelectPhoto5(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic5.value=ss[0];
  }
}

function SelectPhoto6(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic6.value=ss[0];
  }
}


function SelectPhoto7(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic7.value=ss[0];
  }
}


function SelectPhoto8(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic8.value=ss[0];
  }
}



function SelectPhoto9(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic9.value=ss[0];
  }
}


function SelectPhoto10(){
  var arr=showModalDialog('Admin_selFile.asp?ChannelID=0&UploadDir=UploadPic', '', 'dialogWidth:800px; dialogHeight:600px; help: no; scroll: yes; status: no');
  if(arr!=null){
    var ss=arr.split('|');
    document.myform.DefaultPic10.value=ss[0];
  }
}




</script>

<%                                     
if Action="SaveModify" then
	call DoSaveRecord
else
	call Show()
end if

If FoundErr = True Then
	ReturnError(ErrMsg)
End If
Admin_footer
SaveLogInfo(AdminName)
CloseConn


Sub Show()
	dim rsInfo
	dim pic(9)
	dim picurl(9)
	dim pictxt(9)
	FoundErr=False
	Set rsInfo = enchiasp.Execute("select * From eccms_dibu order by id")
	if rsInfo.bof and rsInfo.eof then
			FoundErr = True
		ErrMsg = ErrMsg + "<li>发生未知错误，请联系系统供应商！</li>"
		Exit Sub
	else
		dim i
		i=0
		do while not rsinfo.eof
			pic (i)=rsinfo("pic")
			picurl(i)=rsinfo("picurl")
			pictxt(i)=rsinfo("pictext")
			i=i+1
			rsinfo.movenext
		loop
	
%>
<form method="POST" name="myform" onSubmit="Submit;" action="Admin_dibu.asp">
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
 <th>文件上传</th>
 <tr>
 <td>
 <input name='picurl' id=ImageUrl type='hidden' size=60>
<iframe name="image" frameborder=0 width=100% height=42 scrolling=no src=Upload.asp?sType=AD></iframe> </td>
 </tr>

 </table>
 <br>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2" class=tableborder>
<tr>
<th colspan="3">图片1</th>
 </tr>
 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片1：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic1" type="text" id="DefaultPic1" value="<%=pic(0)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto1()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片1连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl1" type="text" id="DefaultPicUrl1" value="<%=picurl(0)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片1说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt1" type="text" id="DefaultPictxt1" value="<%=pictxt(0)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>
  <br>
  
   <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片2</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片2：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic2" type="text" id="DefaultPic2" value="<%=pic(1)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto2()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片2连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl2" type="text" id="DefaultPicUrl2" value="<%=picurl(1)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片2说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt2" type="text" id="DefaultPictxt2" value="<%=pictxt(1)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>

  <br>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片3</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片3：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic3" type="text" id="DefaultPic3" value="<%=pic(2)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto3()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片3连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl3" type="text" id="DefaultPicUrl3" value="<%=picurl(2)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片3说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt3" type="text" id="DefaultPictxt3" value="<%=pictxt(2)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>
  <br>

 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片4</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片4：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic4" type="text" id="DefaultPic4" value="<%=pic(3)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto4()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片4连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl4" type="text" id="DefaultPicUrl4" value="<%=picurl(3)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片4说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt4" type="text" id="DefaultPictxt4" value="<%=pictxt(3)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>
  <br>

 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片5</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片5：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic5" type="text" id="DefaultPic5" value="<%=pic(4)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto5()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片5连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl5" type="text" id="DefaultPicUrl5" value="<%=picurl(4)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片5说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt5" type="text" id="DefaultPictxt5" value="<%=pictxt(4)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>

  <br>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片6</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片6：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic6" type="text" id="DefaultPic6" value="<%=pic(5)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto6()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片6连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl6" type="text" id="DefaultPicUrl6" value="<%=picurl(5)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片6说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt6" type="text" id="DefaultPictxt6" value="<%=pictxt(5)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>
  <br>

 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片7</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片7：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic7" type="text" id="DefaultPic7" value="<%=pic(6)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto7()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片7连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl7" type="text" id="DefaultPicUrl7" value="<%=picurl(6)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片7说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt7" type="text" id="DefaultPictxt7" value="<%=pictxt(6)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>

  <br>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片8</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片8：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic8" type="text" id="DefaultPic8" value="<%=pic(7)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto8()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片8连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl8" type="text" id="DefaultPicUrl8" value="<%=picurl(7)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片8说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt8" type="text" id="DefaultPictxt8" value="<%=pictxt(7)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>

  <br>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片9</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片9：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic9" type="text" id="DefaultPic9" value="<%=pic(8)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto9()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片9连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl9" type="text" id="DefaultPicUrl9" value="<%=picurl(8)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片9说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt9" type="text" id="DefaultPictxt9" value="<%=pictxt(8)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>
  <br>

 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
<tr>
<th colspan="3">图片10</th>
 </tr>

 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片10：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic10" type="text" id="DefaultPic10" value="<%=pic(9)%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传图片中选择' onclick='SelectPhoto10()' class=button>
            </td>
          </tr>
   <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片10连接路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPicUrl10" type="text" id="DefaultPicUrl10" value="<%=picurl(9)%>" size="80" maxlength="200">
             
            </td>
          </tr>
     <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>图片10说明：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPictxt10" type="text" id="DefaultPictxt10" value="<%=pictxt(9)%>" size="80" maxlength="200">
             
            </td>
          </tr>    
  </table>

  
  
  
  
  
  
  
  
  
  
   
    <div align="center"> 
    <p> 
		<input name="Action" type="hidden" id="Action" value="SaveModify">
		<input name="Save" type="submit"  id="Save" value="保 存" style="cursor:hand;">
	<input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_dibu.Asp'" style="cursor:hand;">
    </p>
  </div>
</form>

<% 
		rsinfo.close
		set rsinfo=nothing
	end if
end sub
Sub DoSaveRecord()
	dim pic(9)
	dim picurl(9)
	dim pictxt(9)
	dim i
	for i=0 to 9
		pic(i)=Trim(request.form("DefaultPic"&i+1))
		picurl(i)=Trim(request.form("DefaultPicurl"&i+1))
		pictxt(i)=Trim(request.form("DefaultPictxt"&i+1))
	next 
	
		for i=1 to 10
			'if pic(i-1)<>"" then
				enchiasp.Execute("update eccms_dibu set pic='"& pic(i-1) &"',picurl='"& picurl(i-1) &"',pictext='"& pictxt(i-1) &"' where id="& i)
			'end if
		next
	Succeed("<li>恭喜您！修改成功。</li>")
End Sub


%>