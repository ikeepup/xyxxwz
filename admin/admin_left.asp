<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
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

Response.Expires = 0  
Response.ExpiresAbsolute = Now() - 1  
Response.cachecontrol = "no-cache" 

%>
<head>
<title><%= enchiasp.SiteName %> - ����ҳ��</title>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<style type=text/css>
BODY{
	margin:0px;
	FONT-SIZE: 12px;
	FONT-FAMILY: "����", "Verdana", "Arial", "Helvetica", "sans-serif";
	background-color: #7D899D; 
	scrollbar-highlight-color: #98A0AD; 
	scrollbar-arrow-color: #FFFFFF; 
	scrollbar-base-color: #7D899D
}
table { border:0px; }
td { font:normal 12px ����; }
img { vertical-align:bottom; border:0px; }
a { font:normal 12px ����; color:#111111; text-decoration:none; }
a:hover { color:#384780;text-decoration:underline; }
.sec_menu { border-left:2px solid #335EA8; border-right:2px solid #335EA8; border-bottom:2px solid #335EA8; overflow:hidden; background:#EEEEE3; }
.menu_title {background-color: #335EA8;color: white;font-size: 12px;font-weight:bold;height: 25; }
.menu_title span { position:relative; top:2px; left:8px; color:#FFFF00; font-weight:bold; }
.menu_title2 {background-color: #335EA8;color: white;font-size: 12px;font-weight:bold;height: 25; }
.menu_title2 span { position:relative; top:2px; left:8px; color:#FEFEFE; font-weight:bold; }
input,select,Textarea{
font-family:����,Verdana, Arial, Helvetica, sans-serif; font-size: 12px;}
</style>
<script language=JavaScript>
function logout(){
	if (confirm("ϵͳ��ʾ����ȷ��Ҫ�˳����������"))
	top.location = "logout.asp";
	return false;
}
</script>
<script language=JavaScript1.2>
function showsubmenu(sid) {
	var whichEl = eval("submenu" + sid);
	var menuTitle = eval("menuTitle" + sid);
	if (whichEl.style.display == "none"){
		eval("submenu" + sid + ".style.display=\"\";");
		//if (sid != 0 & sid < 999) {
			//menuTitle.background="images/title_bg_hide.gif";
		//}
	}else{
		eval("submenu" + sid + ".style.display=\"none\";");
		//if (sid != 0 & sid < 999) {
			//menuTitle.background="images/title_bg_show.gif";
		//}
	}
}
</script>
</head>
<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left> 
<tr>
  <td valign=top> <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr> 
        <td height=42 valign=bottom> <img src="images/admin_title.gif" width=158 height=38 style="vertical-align:bottom; border:0px;"> 
        </td>
      </tr>
    </table>
    <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr> 
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin_left_1.jpg> 
          <a href="admin_main.asp" target=main><span>������ҳ</span></a> <span>|</span> <a href="#"  onclick="logout();"><span>�� ��</b></span> 
        </td>
      </tr>
      <tr> 
        <td style="display:"> <div class=sec_menu style="width:158"> 
            
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr> 
                <td height=20>�û�����<font color=red><%=session("AdminName")%></font></td>
              </tr>
              <tr> 
                <td height=20>���ݣ�<font color=red><%=Session("AdminStatus")%></font></td>
              </tr></table></td>
    </table>
    <%
    		if  wordcheck()  then

    %>
    <div  style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr>
            <td height=5></td>
              </tr>
            </table>
     </div>
     <%
     	If Chkreg("SiteConfig") Then

     %>
    <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr> 
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin_left.jpg id=menuTitle0 onClick="showsubmenu(0)"> 
          <span>��������</span> </td>
      </tr>
      <tr> 
        <td style="display:" id='submenu0'> <div class=sec_menu style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
            	 <%If ChkAdmin("SiteConfig") Then
           		 %>
              <tr> 
               <td height=20><img src="images/bullet.gif"><a href=admin_config.asp target=main>��������</a></td>
              </tr>
              <%
              end if
              If ChkAdmin("fengmian") Then
           		 %>
              <tr> 
               <td height=20><img src="images/bullet.gif"><a href=admin_fengmian.asp target=main>��������</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("CreateIndex") Then
              %>
                            <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_createindex.asp target=main>������ҳ��HTML��</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("Template") Then
              %>             
              <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_template.asp target=main>ģ����ʽ�ܹ���</a></td>
              </tr>              

    <%
              end if
              If ChkAdmin("TemplateLoad") Then
              %> 
              
  		<tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_loadskin.asp target=main>ģ�嵼��</a> | <a href=admin_loadskin.asp?action=load target=main>ģ�嵼��</a></td>
              </tr>
              <%
              end if
              If ChkAdmin("Channel") Then
              %> 
			  <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_channel.asp?action=add target=main>���Ƶ��</a> | <a href=admin_channel.asp target=main>Ƶ������</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("Announce") Then
              %>
              
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_announce.asp?action=add target=main>��������</a> | <a href=admin_announce.asp target=main>�������</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("rizhi") or ChkAdmin("SendMessage") Then
              %>

              
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_log.asp target=main>��־����</a> | <a href=admin_message.asp target=main>���Ͷ���</a></td>
              </tr>
              
                <%
              end if
              If ChkAdmin("Advertise") Then
              %>


	     
              
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_admanage.asp?action=add target=main>��ӹ��</a> | <a href=admin_admanage.asp target=main>������</a></td>
              </tr>
                 <%
              end if

              %>

            </table>
          </div>
          <div  style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr>
                <td height=5></td>
              </tr>
            </table>
          </div></td>
      </tr>
    </table>
    <%
    end if
    %>
<%
Dim ChannelName,ChannelDir,ModuleName,strModules
Set Rs = enchiasp.Execute("SELECT ChannelID,ChannelName,ChannelDir,modules,ModuleName FROM ECCMS_Channel WHERE StopChannel = 0 And ChannelType < 2 ORDER BY orders ASC")
Do While Not Rs.EOF
	ChannelID = Rs("ChannelID")
	ChannelName = Rs("ChannelName")
	ChannelDir = Replace(Rs("ChannelDir"), "/", "")
	ModuleName = Rs("ModuleName")
	Select Case Rs("modules")
		Case 1:strModules = "Article"
		Case 2:strModules = "Soft"
		Case 3:strModules = "Shop"
		Case 5:strModules = "Flash"
		Case 6:strModules = "yemian"
		Case 7:strModules = "job"
	Case Else
		strModules = "Article"
	End Select
if Rs("modules")<>4 then

%>
 <%
     	If Chkreg(strModules) Then

     %>

    <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr> 
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin_left.jpg id=menuTitle<%=ChannelID%> onClick="showsubmenu(<%=ChannelID%>)"> 
          <span><%=ChannelName%></span> </td>
      </tr>
      <tr> 
        <td style="display:none" id='submenu<%=ChannelID%>'> <div class=sec_menu style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
            <%
            If  Rs("modules")=7 Then
	            If ChkAdmin("Add"& strModules) or ChkAdmin("Admin"& strModules) Then
				%>
				 <tr> 
	                <td height=20><img src="images/bullet.gif"><a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%>&action=add target=main>���<%=ModuleName%></a> | <a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>����</a></td>
	              </tr>	
	            <% 
	            end if
            else
	            If ChkAdmin("Add"& strModules & ChannelID) or ChkAdmin("Admin"& strModules & ChannelID) Then
	           	 %>
	              <tr> 
	                <td height=20><img src="images/bullet.gif"><a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%>&action=add target=main>���<%=ModuleName%></a> | <a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>����</a></td>
	              </tr>
	               <%
	                end if
            end if

  	If  Rs("modules")=7 Then
		if ChkAdmin("adminjobbook") then
			%>
	   		<tr> 
	             <td height=20><img src="images/bullet.gif"><a href=admin_jobbook.asp target=main>ӦƸ����</a> | <a href=admin_jobbook.asp?isdel=1 target=main>����վ</a></td>
	
	        </tr>
			<%
		end if
		if ChkAdmin("Adminclassjob" & ChannelID) Then
			%>
			   
		   		<tr> 
		            <td height=20><img src="images/bullet.gif"><a href=admin_classifyjob.asp?ChannelID=<%=ChannelID%>&action=add target=main>�����Ŀ</a> | <a href=admin_classifyjob.asp?ChannelID=<%=ChannelID%> target=main>��Ŀ����</a></td>
		
		        </tr>
			<%
		end if
	else
		If ChkAdmin("AdminClass" & ChannelID) Then

		%>
	 	<tr> 
	          <td height=20><img src="images/bullet.gif"><a href=admin_classify.asp?ChannelID=<%=ChannelID%>&action=add target=main>��ӷ���</a> | <a href=admin_classify.asp?ChannelID=<%=ChannelID%> target=main>�������</a></td>
	    </tr>


	<%
		End If
	end if
	If Rs("modules") <> 3 and  Rs("modules") <> 6 and  Rs("modules") <> 7 Then
		If ChkAdmin("Special" & ChannelID) or ChkAdmin("Admin"& strModules & ChannelID)  Then
	%>
	              <tr> 
	                <td height=20><img src="images/bullet.gif"><a href=admin_special.asp?ChannelID=<%=ChannelID%> target=main>ר�����</a> | <a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%>&isAccept=0 target=main>��˹���</a></td>
	              </tr>
	<%
		end if
	End If
	If Rs("modules") <> 6 and  Rs("modules") <> 7 Then
		If ChkAdmin("Admin"& strModules & ChannelID)  Then
	
	%>
	
		      <tr> 
	                <td height=20><img src="images/bullet.gif"><a href=admin_<%=strModules%>.asp?action=setting&ChannelID=<%=ChannelID%> target=main>��������</a> | <a href=admin_<%=strModules%>.asp?action=move&ChannelID=<%=ChannelID%> target=main>�����ƶ�</a></td>
	              </tr>
	<%
		End If
	End If
	If Rs("modules") = 2 Then
		If ChkAdmin("DownServer" & ChannelID)  Then


%>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_server.asp?ChannelID=<%=ChannelID%> target=main>���ط���������</a></td>
              </tr>
              <%
         end if
         if ChkAdmin("ErrorSoft" & ChannelID) or ChkAdmin("SoftCollect") then
              %>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_softerr.asp?ChannelID=<%=ChannelID%> target=main>���󱨸�</a> | <a href=Admin_SoftGather.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>�ɼ�</a></td>
              </tr>
<%
		end if
	ElseIf Rs("modules") = 1 Then
		If ChkAdmin("ArticleGather")  Then

	
%>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=Admin_ArticleGather.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>�ɼ�����</a></td>
              </tr>
<%
		end if
	ElseIf Rs("modules") = 5 Then
	
			If ChkAdmin("DownServer" & ChannelID)  Then

%>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_server.asp?ChannelID=<%=ChannelID%> target=main>���ط���������</a></td>
              </tr>
              <%
              end if
              If ChkAdmin("FlashGather")  Then

              %>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=Admin_FlashGather.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>�ɼ�����</a></td>
              </tr>
              
<%
			end if
	End If
	If ChkAdmin("Channel") then
	
%>
              <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_channel.asp?action=edit&ChannelID=<%=ChannelID%> target=main><%=ModuleName%>Ƶ����������</a></td>
              </tr>
<%
	end if
	If  Rs("modules") <> 7 Then
	If ChkAdmin("Create" & strModules & ChannelID) then

%>	
	      <tr> 

                <td height=20><img src="images/bullet.gif"><a href=admin_create<%=strModules%>.asp?ChannelID=<%=ChannelID%> target=main>����<%=ModuleName%>HTML����</a></td>
        
 </tr>
<%
	End If
	End If
	If ChkAdmin("Template") Then

%> 
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_template.asp?action=manage&ChannelID=<%=ChannelID%> target=main><%=ModuleName%>Ƶ��ģ�����</a></td>
              </tr>
<%
	end if
	If Rs("modules") <> 6 and  Rs("modules") <> 7 Then
	If ChkAdmin("Comment" & ChannelID) Then

%>	     
 <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_comment.asp?ChannelID=<%=ChannelID%> target=main>���۹���</a> | <a href=admin_jsfile.asp?ChannelID=<%=ChannelID%> target=main>JS ����</a></td>
 </tr>
<%
	end if
	End If
			If ChkAdmin("AdminUpload" & ChannelID) Then

%>              
	      <tr>
                <td height=20><img src="images/bullet.gif"><a href=Admin_UploadFile.Asp?ChannelID=<%=ChannelID%>&UploadDir=UploadPic target=main>�ϴ�ͼƬ����</a> | <a href=admin_UploadFile.asp?action=clear&ChannelID=<%=ChannelID%>&UploadDir=UploadPic target=main>����</a></td>
              </tr>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=Admin_UploadFile.Asp?ChannelID=<%=ChannelID%>&UploadDir=UploadFile target=main>�ϴ��ļ�����</a> | <a href=admin_UploadFile.asp?action=clear&ChannelID=<%=ChannelID%>&UploadDir=UploadFile target=main>����</a></td>
              </tr>
              <%
              end if
              %>
            </table>
          </div>
          <div  style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr>
                <td height=5></td>
              </tr>
            </table>
          </div></td>
      </tr>
    </table>
<%
	end if
	End If
	Rs.movenext
Loop
Set Rs = Nothing

%>

    <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr>   
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin_left.jpg id=menuTitle1000 onClick="showsubmenu(1000)"> 
          <span>�û�����</span> </td>
      </tr>
      <tr> 
        <td style="display:none" id='submenu1000'> <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
 
       <%

       If ChkAdmin("ChangePassword") Then

       %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_password.asp target=main>����Ա�����޸�</a></td>
          </tr>
          <%
          end if
          If ChkAdmin("999") Then
          %>
           <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_master.asp?action=add target=main>����Ա���</a> | <a href=admin_master.asp target=main>����</a></td>
          </tr>
          <%
          end if
         
          %>
	
    
	 

	          </table>
          </div>
        <div  style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr>
                <td height=5></td>
              </tr>
            </table>
          </div></td>
      </tr>
    </table>

    <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr>   
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin_left.jpg id=menuTitle1001 onClick="showsubmenu(1001)"> 
          <span>��������</span> </td>
      </tr>
      <tr> 
        <td style="display:none" id='submenu1001'> <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
          <%
          If ChkAdmin("Online") Then

          %>
          
		<%
		end if
		 If ChkAdmin("flashtupian") Then

          %>
                
      
           <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_tupian.asp target=main> ����ͼƬ�任����</a></td>
          </tr>

          <%
          end if
	
        If ChkAdmin("FriendLink") or ChkAdmin("GuestBook") Then

        %>
          <tr> 
            <td height=20><img src="images/bullet.gif">
              <a href="admin_link.asp" target="main">��������</a> | <a href="admin_book.asp" target="main">���Թ���</a></td>
          </tr>
          <%
          end if
          %>
	         </table>
          </div>
        <div  style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr>
                <td height=5></td>
              </tr>
            </table>
          </div></td>
      </tr>
    </table>
    <%
    end if
    %>
     <table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr>   
        <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin_left.jpg id=menuTitle1002 onClick="showsubmenu(1002)"> 
          <span>���ݿ⴦��</span> </td>
      </tr>
      <tr> 
        <td style="display:none" id='submenu1002'> <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
        <%
        If ChkAdmin("BackupData") Then
		%>
          <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=BackupData target=main>���ݿⱸ��</a></td>
          </tr>
          <%
         end if
         If ChkAdmin("RestoreData") Then

          %>
          <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=RestoreData target=main>���ݿ�ָ�</a></td>
          </tr>
          <%
          end if
          If  ChkAdmin("CompressData") Then

          %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=CompressData target=main>���ݿ�ѹ��</a></td>
          </tr>
          <%
          end if
          
          %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_probe.asp target=main>��������Ϣ</a></td>
          </tr>
<%
If ChkAdmin("BatchReplace") Then

%>	 
 <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_replace.asp target=main>���ݿ������滻</a></td>
          </tr>
 <%
 end if
 If ChkAdmin("SpaceSize") Then

 %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=Spacesize target=main>ϵͳ�ռ�ռ��</a></td>
          </tr>
 <%
 end if
 %>
        </table>
          </div>
        <div  style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr>
                <td height=5></td>
              </tr>
            </table>
          </div></td>
      </tr>
    </table>
<table cellpadding=0 cellspacing=0 width=158 align=center>
      <tr>

	<td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin_left.jpg>
	  <span>ϵͳ��Ϣ</span> </td>
      </tr>
      <tr>
	<td> <div class=sec_menu style="width:158">
	<br>
	    <table cellpadding=0 cellspacing=0 align=center width=135>
	    <tr>
	    <td height=20><img src="images/bullet.gif">
	     <a href="admin_reg.asp" target="main">������������ע��</a>
	    </td>
	    </tr>
	      <tr>
		<td height=20><img src="images/bullet.gif">

		  <a href="http://www.enchi.com.cn/" target=_blank>��Ȩ���У��������</a>
	      </tr>
	           <tr>
		<td height=20><img src="images/bullet.gif">

	      <a href="http://www.enchi.com.cn/" target=_blank>����֧�֣�liuyunfan</A><br></td>
	      </tr>

	      
	      
	       <tr align=center>
		<td height=22>��<a href="logout.asp" target=_top>ע���˳�</a>��<br></td>
	      </tr>
	    </table>
	  </div></td>
      </tr>
    </table>
    <BR style="OVERFLOW: hidden; LINE-HEIGHT: 5px">
</body>
</html>



















































