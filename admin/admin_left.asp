<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
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

Response.Expires = 0  
Response.ExpiresAbsolute = Now() - 1  
Response.cachecontrol = "no-cache" 

%>
<head>
<title><%= enchiasp.SiteName %> - 管理页面</title>
<META content="MSHTML 5.00.3315.2870" name=GENERATOR>
<META http-equiv=Content-Type content=text/html;charset=gb2312>
<style type=text/css>
BODY{
	margin:0px;
	FONT-SIZE: 12px;
	FONT-FAMILY: "宋体", "Verdana", "Arial", "Helvetica", "sans-serif";
	background-color: #7D899D; 
	scrollbar-highlight-color: #98A0AD; 
	scrollbar-arrow-color: #FFFFFF; 
	scrollbar-base-color: #7D899D
}
table { border:0px; }
td { font:normal 12px 宋体; }
img { vertical-align:bottom; border:0px; }
a { font:normal 12px 宋体; color:#111111; text-decoration:none; }
a:hover { color:#384780;text-decoration:underline; }
.sec_menu { border-left:2px solid #335EA8; border-right:2px solid #335EA8; border-bottom:2px solid #335EA8; overflow:hidden; background:#EEEEE3; }
.menu_title {background-color: #335EA8;color: white;font-size: 12px;font-weight:bold;height: 25; }
.menu_title span { position:relative; top:2px; left:8px; color:#FFFF00; font-weight:bold; }
.menu_title2 {background-color: #335EA8;color: white;font-size: 12px;font-weight:bold;height: 25; }
.menu_title2 span { position:relative; top:2px; left:8px; color:#FEFEFE; font-weight:bold; }
input,select,Textarea{
font-family:宋体,Verdana, Arial, Helvetica, sans-serif; font-size: 12px;}
</style>
<script language=JavaScript>
function logout(){
	if (confirm("系统提示：您确定要退出控制面板吗？"))
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
          <a href="admin_main.asp" target=main><span>管理首页</span></a> <span>|</span> <a href="#"  onclick="logout();"><span>退 出</b></span> 
        </td>
      </tr>
      <tr> 
        <td style="display:"> <div class=sec_menu style="width:158"> 
            
            <table cellpadding=0 cellspacing=0 align=center width=135>
              <tr> 
                <td height=20>用户名：<font color=red><%=session("AdminName")%></font></td>
              </tr>
              <tr> 
                <td height=20>身　份：<font color=red><%=Session("AdminStatus")%></font></td>
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
          <span>常规设置</span> </td>
      </tr>
      <tr> 
        <td style="display:" id='submenu0'> <div class=sec_menu style="width:158"> 
            <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
            	 <%If ChkAdmin("SiteConfig") Then
           		 %>
              <tr> 
               <td height=20><img src="images/bullet.gif"><a href=admin_config.asp target=main>基本设置</a></td>
              </tr>
              <%
              end if
              If ChkAdmin("fengmian") Then
           		 %>
              <tr> 
               <td height=20><img src="images/bullet.gif"><a href=admin_fengmian.asp target=main>封面设置</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("CreateIndex") Then
              %>
                            <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_createindex.asp target=main>生成首页（HTML）</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("Template") Then
              %>             
              <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_template.asp target=main>模板样式总管理</a></td>
              </tr>              

    <%
              end if
              If ChkAdmin("TemplateLoad") Then
              %> 
              
  		<tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_loadskin.asp target=main>模板导出</a> | <a href=admin_loadskin.asp?action=load target=main>模板导入</a></td>
              </tr>
              <%
              end if
              If ChkAdmin("Channel") Then
              %> 
			  <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_channel.asp?action=add target=main>添加频道</a> | <a href=admin_channel.asp target=main>频道管理</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("Announce") Then
              %>
              
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_announce.asp?action=add target=main>发布公告</a> | <a href=admin_announce.asp target=main>公告管理</a></td>
              </tr>
               <%
              end if
              If ChkAdmin("rizhi") or ChkAdmin("SendMessage") Then
              %>

              
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_log.asp target=main>日志管理</a> | <a href=admin_message.asp target=main>发送短信</a></td>
              </tr>
              
                <%
              end if
              If ChkAdmin("Advertise") Then
              %>


	     
              
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_admanage.asp?action=add target=main>添加广告</a> | <a href=admin_admanage.asp target=main>广告管理</a></td>
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
	                <td height=20><img src="images/bullet.gif"><a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%>&action=add target=main>添加<%=ModuleName%></a> | <a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>管理</a></td>
	              </tr>	
	            <% 
	            end if
            else
	            If ChkAdmin("Add"& strModules & ChannelID) or ChkAdmin("Admin"& strModules & ChannelID) Then
	           	 %>
	              <tr> 
	                <td height=20><img src="images/bullet.gif"><a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%>&action=add target=main>添加<%=ModuleName%></a> | <a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>管理</a></td>
	              </tr>
	               <%
	                end if
            end if

  	If  Rs("modules")=7 Then
		if ChkAdmin("adminjobbook") then
			%>
	   		<tr> 
	             <td height=20><img src="images/bullet.gif"><a href=admin_jobbook.asp target=main>应聘管理</a> | <a href=admin_jobbook.asp?isdel=1 target=main>回收站</a></td>
	
	        </tr>
			<%
		end if
		if ChkAdmin("Adminclassjob" & ChannelID) Then
			%>
			   
		   		<tr> 
		            <td height=20><img src="images/bullet.gif"><a href=admin_classifyjob.asp?ChannelID=<%=ChannelID%>&action=add target=main>添加栏目</a> | <a href=admin_classifyjob.asp?ChannelID=<%=ChannelID%> target=main>栏目管理</a></td>
		
		        </tr>
			<%
		end if
	else
		If ChkAdmin("AdminClass" & ChannelID) Then

		%>
	 	<tr> 
	          <td height=20><img src="images/bullet.gif"><a href=admin_classify.asp?ChannelID=<%=ChannelID%>&action=add target=main>添加分类</a> | <a href=admin_classify.asp?ChannelID=<%=ChannelID%> target=main>分类管理</a></td>
	    </tr>


	<%
		End If
	end if
	If Rs("modules") <> 3 and  Rs("modules") <> 6 and  Rs("modules") <> 7 Then
		If ChkAdmin("Special" & ChannelID) or ChkAdmin("Admin"& strModules & ChannelID)  Then
	%>
	              <tr> 
	                <td height=20><img src="images/bullet.gif"><a href=admin_special.asp?ChannelID=<%=ChannelID%> target=main>专题管理</a> | <a href=admin_<%=strModules%>.asp?ChannelID=<%=ChannelID%>&isAccept=0 target=main>审核管理</a></td>
	              </tr>
	<%
		end if
	End If
	If Rs("modules") <> 6 and  Rs("modules") <> 7 Then
		If ChkAdmin("Admin"& strModules & ChannelID)  Then
	
	%>
	
		      <tr> 
	                <td height=20><img src="images/bullet.gif"><a href=admin_<%=strModules%>.asp?action=setting&ChannelID=<%=ChannelID%> target=main>批量设置</a> | <a href=admin_<%=strModules%>.asp?action=move&ChannelID=<%=ChannelID%> target=main>批量移动</a></td>
	              </tr>
	<%
		End If
	End If
	If Rs("modules") = 2 Then
		If ChkAdmin("DownServer" & ChannelID)  Then


%>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_server.asp?ChannelID=<%=ChannelID%> target=main>下载服务器管理</a></td>
              </tr>
              <%
         end if
         if ChkAdmin("ErrorSoft" & ChannelID) or ChkAdmin("SoftCollect") then
              %>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_softerr.asp?ChannelID=<%=ChannelID%> target=main>错误报告</a> | <a href=Admin_SoftGather.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>采集</a></td>
              </tr>
<%
		end if
	ElseIf Rs("modules") = 1 Then
		If ChkAdmin("ArticleGather")  Then

	
%>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=Admin_ArticleGather.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>采集管理</a></td>
              </tr>
<%
		end if
	ElseIf Rs("modules") = 5 Then
	
			If ChkAdmin("DownServer" & ChannelID)  Then

%>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_server.asp?ChannelID=<%=ChannelID%> target=main>下载服务器管理</a></td>
              </tr>
              <%
              end if
              If ChkAdmin("FlashGather")  Then

              %>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=Admin_FlashGather.asp?ChannelID=<%=ChannelID%> target=main><%=ModuleName%>采集管理</a></td>
              </tr>
              
<%
			end if
	End If
	If ChkAdmin("Channel") then
	
%>
              <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_channel.asp?action=edit&ChannelID=<%=ChannelID%> target=main><%=ModuleName%>频道基本设置</a></td>
              </tr>
<%
	end if
	If  Rs("modules") <> 7 Then
	If ChkAdmin("Create" & strModules & ChannelID) then

%>	
	      <tr> 

                <td height=20><img src="images/bullet.gif"><a href=admin_create<%=strModules%>.asp?ChannelID=<%=ChannelID%> target=main>生成<%=ModuleName%>HTML管理</a></td>
        
 </tr>
<%
	End If
	End If
	If ChkAdmin("Template") Then

%> 
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_template.asp?action=manage&ChannelID=<%=ChannelID%> target=main><%=ModuleName%>频道模板管理</a></td>
              </tr>
<%
	end if
	If Rs("modules") <> 6 and  Rs("modules") <> 7 Then
	If ChkAdmin("Comment" & ChannelID) Then

%>	     
 <tr> 
                <td height=20><img src="images/bullet.gif"><a href=admin_comment.asp?ChannelID=<%=ChannelID%> target=main>评论管理</a> | <a href=admin_jsfile.asp?ChannelID=<%=ChannelID%> target=main>JS 管理</a></td>
 </tr>
<%
	end if
	End If
			If ChkAdmin("AdminUpload" & ChannelID) Then

%>              
	      <tr>
                <td height=20><img src="images/bullet.gif"><a href=Admin_UploadFile.Asp?ChannelID=<%=ChannelID%>&UploadDir=UploadPic target=main>上传图片管理</a> | <a href=admin_UploadFile.asp?action=clear&ChannelID=<%=ChannelID%>&UploadDir=UploadPic target=main>清理</a></td>
              </tr>
	      <tr> 
                <td height=20><img src="images/bullet.gif"><a href=Admin_UploadFile.Asp?ChannelID=<%=ChannelID%>&UploadDir=UploadFile target=main>上传文件管理</a> | <a href=admin_UploadFile.asp?action=clear&ChannelID=<%=ChannelID%>&UploadDir=UploadFile target=main>清理</a></td>
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
          <span>用户管理</span> </td>
      </tr>
      <tr> 
        <td style="display:none" id='submenu1000'> <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
 
       <%

       If ChkAdmin("ChangePassword") Then

       %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_password.asp target=main>管理员密码修改</a></td>
          </tr>
          <%
          end if
          If ChkAdmin("999") Then
          %>
           <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_master.asp?action=add target=main>管理员添加</a> | <a href=admin_master.asp target=main>管理</a></td>
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
          <span>其它管理</span> </td>
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
            <td height=20><img src="images/bullet.gif"><a href=admin_tupian.asp target=main> 新闻图片变换管理</a></td>
          </tr>

          <%
          end if
	
        If ChkAdmin("FriendLink") or ChkAdmin("GuestBook") Then

        %>
          <tr> 
            <td height=20><img src="images/bullet.gif">
              <a href="admin_link.asp" target="main">友情连接</a> | <a href="admin_book.asp" target="main">留言管理</a></td>
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
          <span>数据库处理</span> </td>
      </tr>
      <tr> 
        <td style="display:none" id='submenu1002'> <div class=sec_menu style="width:158"> 
        <table cellpadding=0 cellspacing=0 align=center width=135 style="POSITION: relative; TOP: 5px">
        <%
        If ChkAdmin("BackupData") Then
		%>
          <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=BackupData target=main>数据库备份</a></td>
          </tr>
          <%
         end if
         If ChkAdmin("RestoreData") Then

          %>
          <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=RestoreData target=main>数据库恢复</a></td>
          </tr>
          <%
          end if
          If  ChkAdmin("CompressData") Then

          %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=CompressData target=main>数据库压缩</a></td>
          </tr>
          <%
          end if
          
          %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_probe.asp target=main>服务器信息</a></td>
          </tr>
<%
If ChkAdmin("BatchReplace") Then

%>	 
 <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_replace.asp target=main>数据库批量替换</a></td>
          </tr>
 <%
 end if
 If ChkAdmin("SpaceSize") Then

 %>
	  <tr> 
            <td height=20><img src="images/bullet.gif"><a href=admin_database.asp?action=Spacesize target=main>系统空间占用</a></td>
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
	  <span>系统信息</span> </td>
      </tr>
      <tr>
	<td> <div class=sec_menu style="width:158">
	<br>
	    <table cellpadding=0 cellspacing=0 align=center width=135>
	    <tr>
	    <td height=20><img src="images/bullet.gif">
	     <a href="admin_reg.asp" target="main">点这里进行软件注册</a>
	    </td>
	    </tr>
	      <tr>
		<td height=20><img src="images/bullet.gif">

		  <a href="http://www.enchi.com.cn/" target=_blank>版权所有：恩池软件</a>
	      </tr>
	           <tr>
		<td height=20><img src="images/bullet.gif">

	      <a href="http://www.enchi.com.cn/" target=_blank>技术支持：liuyunfan</A><br></td>
	      </tr>

	      
	      
	       <tr align=center>
		<td height=22>【<a href="logout.asp" target=_top>注销退出</a>】<br></td>
	      </tr>
	    </table>
	  </div></td>
      </tr>
    </table>
    <BR style="OVERFLOW: hidden; LINE-HEIGHT: 5px">
</body>
</html>



















































