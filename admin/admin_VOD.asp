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
If Not ChkAdmin("999") Then
	Server.Transfer("showerr.asp")
	Response.End
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
     <td height="22" colspan=2 align=center class=tablerow2><a name="Top"></a><strong>首页视频管理</strong></td>
    </tr>
    <tr>
    <td class=tablerow2> 
    调用方法：在需要调用的地方加载如下标签：{$vod},目前仅支持MEDIA PLAY格式文件,如需要其他类型格式文件,请与供应商联系.注意路径为HTTP://全名。 
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
	dim vodpath
	FoundErr=False
	Set rsInfo = enchiasp.Execute("select * From eccms_vod order by id")
	if rsInfo.bof and rsInfo.eof then
	
	else
		vodpath=rsinfo("path")
	end if
	rsinfo.close
	set rsinfo=nothing
%>
<form method="POST" name="myform" onSubmit="Submit;" action="Admin_vod.asp">
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class=tableborder>
 <th>视频文件上传</th>
 <tr>
 <td>
 <input name='picurl' id=ImageUrl type='hidden' size=60>
<iframe name="image" frameborder=0 width=100% height=42 scrolling=no src=Upload.asp?sType=AD></iframe> </td>
 </tr>

 </table>
 <br>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2" class=tableborder>
<tr>
<th colspan="3">视频文件</th>
 </tr>
 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>视频文件路径：</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic1" type="text" id="DefaultPic1" value="<%=vodpath%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='从已上传文件中选择' onclick='SelectPhoto1()' class=button>
            </td>
          </tr>
    </table>
   
   
    <div align="center"> 
    <p> 
		<input name="Action" type="hidden" id="Action" value="SaveModify">
		<input name="Save" type="submit"  id="Save" value="保 存" style="cursor:hand;">
	<input name="Cancel" type="button" id="Cancel" value=" 取 消 " onClick="window.location.href='Admin_vod.Asp'" style="cursor:hand;">
    </p>
  </div>
</form>

<% 
end sub
Sub DoSaveRecord()
	dim vodpath
	vodpath=Trim(request.form("DefaultPic1"))
	enchiasp.Execute("update eccms_vod set path='"& vodpath &"'")
	Succeed("<li>恭喜您！修改成功。</li>")
End Sub


%>