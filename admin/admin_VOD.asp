<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
'=====================================================================
' ������ƣ�������վ����ϵͳ
' ��ǰ�汾��enchicms Version 3.0.0
' �������ڣ�2005-03-25
' �ٷ���վ���˳��ж�������Ƽ��������޹�˾(www.enchi.com.cn) 
' ����֧�֣����Ʒ�
' ���䣺liuyunfan@163.com
' QQ��21556923
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
     <td height="22" colspan=2 align=center class=tablerow2><a name="Top"></a><strong>��ҳ��Ƶ����</strong></td>
    </tr>
    <tr>
    <td class=tablerow2> 
    ���÷���������Ҫ���õĵط��������±�ǩ��{$vod},Ŀǰ��֧��MEDIA PLAY��ʽ�ļ�,����Ҫ�������͸�ʽ�ļ�,���빩Ӧ����ϵ.ע��·��ΪHTTP://ȫ���� 
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
 <th>��Ƶ�ļ��ϴ�</th>
 <tr>
 <td>
 <input name='picurl' id=ImageUrl type='hidden' size=60>
<iframe name="image" frameborder=0 width=100% height=42 scrolling=no src=Upload.asp?sType=AD></iframe> </td>
 </tr>

 </table>
 <br>
 <table width="98%" border="0" align="center" cellpadding="0" cellspacing="2" class=tableborder>
<tr>
<th colspan="3">��Ƶ�ļ�</th>
 </tr>
 <tr class="tdbg"> 
            <td width="100" align="right" class=tablerow2><strong>��Ƶ�ļ�·����</strong></td>
            <td colspan="2" class=tablerow2><input name="DefaultPic1" type="text" id="DefaultPic1" value="<%=vodpath%>" size="80" maxlength="200">
              <br /><input type='button' name='selectpic' value='�����ϴ��ļ���ѡ��' onclick='SelectPhoto1()' class=button>
            </td>
          </tr>
    </table>
   
   
    <div align="center"> 
    <p> 
		<input name="Action" type="hidden" id="Action" value="SaveModify">
		<input name="Save" type="submit"  id="Save" value="�� ��" style="cursor:hand;">
	<input name="Cancel" type="button" id="Cancel" value=" ȡ �� " onClick="window.location.href='Admin_vod.Asp'" style="cursor:hand;">
    </p>
  </div>
</form>

<% 
end sub
Sub DoSaveRecord()
	dim vodpath
	vodpath=Trim(request.form("DefaultPic1"))
	enchiasp.Execute("update eccms_vod set path='"& vodpath &"'")
	Succeed("<li>��ϲ�����޸ĳɹ���</li>")
End Sub


%>