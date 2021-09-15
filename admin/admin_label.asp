<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
Response.Write "<base target=""_self"">" & vbNewLine
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
%>
<script language="javascript">
<!--
function copy() {
	document.myform.SkinCode.focus();
	document.myform.SkinCode.select();
	textRange = document.myform.SkinCode.createTextRange();
	textRange.execCommand("Copy");
}

function selflabel(){
	copy()
	window.close()
}
// -->
</script>
<%
Dim Action,i
Action = LCase(Request("action"))
Select Case Trim(Action)
Case "list"
	Call Label_ContentList
Case "image"
	Call Label_ImageUse
Case "text"
	Call Label_PicAndText
Case Else
	Call showmain
End Select
If FoundErr = True Then
	ReturnError(ErrMsg)
End If
%>
<BR style="OVERFLOW: hidden; LINE-HEIGHT: 5px">
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<td class=tablerow1>请将以上标签复制到模板相应的位置</td>
</tr>
</table>
<%
CloseConn
Private Sub showmain()
	%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>常用模板标签</th>
</tr>
<tr>
	<td class=tablerow1>{$showuserinfo}&nbsp;&nbsp;&nbsp;&nbsp;显示会员登陆状态(相关模版在模版常规设置中27修改)</td> 
	             
</tr>




<tr>
	<td class=tablerow1>{$InstallDir}&nbsp;&nbsp;&nbsp;&nbsp;系统安装路径  （系统自动生成）</td>       
	             
</tr>
<tr>
	<td class=tablerow2>{$SkinPath}&nbsp;&nbsp;&nbsp;&nbsp;皮肤图片路径</td>
</tr>
<tr>
	<td class=tablerow1>{$ChannelRootDir}&nbsp;&nbsp;&nbsp;&nbsp;频道目录路径</td>
</tr>
<tr>
	<td class=tablerow2>{$Version}&nbsp;&nbsp;&nbsp;&nbsp;系统版本信息</td>
</tr>
<tr>
	<td class=tablerow1>{$WebSiteName}&nbsp;&nbsp;&nbsp;&nbsp;网站名称 （在基本设置修改）</td>                    
</tr>
<tr>
	<td class=tablerow2>{$WebSiteUrl}&nbsp;&nbsp;&nbsp;&nbsp;网站URL （在基本设置修改）</td>                    
</tr>
<tr>
	<td class=tablerow1>{$MasterMail}&nbsp;&nbsp;&nbsp;&nbsp;管理员E-Mail（在基本设置修改）</td>
</tr>
<tr>
	<td class=tablerow2>{$Keyword}&nbsp;&nbsp;&nbsp;&nbsp;网站关键字 （在基本设置修改）</td>                    
</tr>
<tr>
	<td class=tablerow1>{$Copyright}&nbsp;&nbsp;&nbsp;&nbsp;网站版权信息 （在基本设置修改）</td>                    
</tr>
<tr>
	<td class=tablerow2>{$Width}&nbsp;&nbsp;&nbsp;&nbsp;定义主表格宽度 </td>
</tr>
<tr>
	<td class=tablerow1>{$IndexPage}&nbsp;&nbsp;&nbsp;&nbsp;默认首页文件名</td>
</tr>
<tr>
	<td class=tablerow2>{$Style_CSS}&nbsp;&nbsp;&nbsp;&nbsp;CSS样式</td>
</tr>
<tr>
	<td class=tablerow1>{$PageTitle}&nbsp;&nbsp;&nbsp;&nbsp;HTML文件标题</td>
</tr>
<tr>
	<td class=tablerow2>{$TotalStatistics}&nbsp;&nbsp;&nbsp;&nbsp;系统总统计</td>
</tr>
<tr>
	<td class=tablerow1>{$RenewStatistics}&nbsp;&nbsp;&nbsp;&nbsp;更新信息统计</td>
</tr>
<tr>
	<td class=tablerow2>{$ChannelMenu}&nbsp;&nbsp;&nbsp;&nbsp;顶部频道菜单标签</td>
</tr>
<tr>
	<td class=tablerow2>{$ShowHotArticle}&nbsp;&nbsp;&nbsp;&nbsp;文章热门点击</td>
</tr>
<tr>
	<th>脚本及特殊标签调用</th>
</tr>

<tr>
	<td class=tablerow1>&lt;script language=&quot;javascript&quot;       
      type=&quot;text/javascript&quot; src=&quot;{$WebSiteUrl}{$InstallDir}qqonline/qq.asp&quot;&gt;&lt;/script&gt;       
      在线QQ管理，在通栏模版中增加，注意有的地方可能无法调用，仅用于各个栏目下使用，在根目录下可能会出现无法调用的情况。修改QQ的其他管理中设置</td>                
</tr>

<tr>
	<td class=tablerow1>{$vod} 目前仅支持MEDIA      
      PLAY格式文件，修改在模板常规设置26中修改</td>                
</tr>

<tr>
	<td class=tablerow1>{$tupianhuan}    
      图片报道管理调用方法：在需要调用的地方加载标签,如果要修改该FLASH的图片大小等参数请在通栏模版<font COLOR="#800000" face="宋体">常规设置</font>中24修改。</td>                
</tr>

<tr>
	<td class=tablerow1>{$dibuhuan}    
      图片左右滚动管理调用方法：在需要调用的地方加载标签,如果要修改该FLASH的图片大小等参数请在通栏模版基本设置中25修改。最多可设置10张图片，建议使用JPG图片</td>                
</tr>

<tr>
	<td class=tablerow1><font face="宋体">  
    &lt;script src=&quot;{$WebSiteUrl}{$InstallDir}count/count.asp&quot;&gt;&lt;/script><br>        
    加通栏中为页面统计</font></td>                
</tr>

<tr>
	<td class=tablerow1>{$ReadShopPic(3,0,0,3,4,4,23,0,100,100,2)}   
      商城图片特殊显示，显示内容在通栏模版<font COLOR="#800000" face="宋体">常规设置</font>中28修改。最后一个为2</td>                
</tr>

<tr>
	<td class=tablerow1>{$ReadArticlePic(1,0,0,0,1,1,120,0,105,79,2)} 新闻图片特殊显示，显示内容在通栏模版<font COLOR="#800000" face="宋体">常规设置</font>中29修改。最后一个为2</td>                  
</tr>

<tr>
	<td class=tablerow1>合作伙伴如需要特殊显示，请在该模板下修改配置</td>                
</tr>

<tr>
	<td class=tablerow1><font face="宋体"><b>     
    &lt;head&gt;&lt;/head&gt;中间通用代码</b></font>
      <p><font face="宋体">&lt;meta http-equiv=&quot;Content-Type&quot; content=&quot;text/html;   
      charset=gb2312&quot;&gt;<br>           
	&lt;title&gt;{$PageTitle} - {$Keyword} &lt;/title&gt;<br>              
	&lt;meta name=&quot;keywords&quot; content=&quot;{$WebSiteName} {$Keyword}&quot;&gt;<br>              
	&lt;meta name=&quot;description&quot; content=&quot;{$PageTitle} ,{$WebSiteName}&quot;&gt;<br>             
	&lt;LINK href=&quot;{$InstallDir}{$SkinPath}style.css&quot; type=text/css   
      rel=stylesheet&gt;<br>           
	&lt;meta name=&quot;MSSmartTagsPreventParsing&quot; content=&quot;TRUE&quot;&gt;<br>              
	&lt;meta http-equiv=&quot;MSThemeCompatible&quot; content=&quot;Yes&quot;&gt;<br>              
	&lt;meta http-equiv=&quot;html&quot; content=&quot;no-cache&quot;&gt;</font></td>                   
</tr>


<tr><td class=tablerow1>
  <p align="left"><font face="宋体">广告位置JS文件 &lt;script language=javascript   
  src={$InstallDir}adfile/banner.js&gt;&lt;/script&gt;</font></td></tr>


<tr align="left"><td class=tablerow1>
  <p align="left"><font face="宋体">菜单JS文件&lt;script src=&quot;{$InstallDir}inc/menu.js&quot; type=&quot;text/javascript&quot;&gt;&lt;/script&gt;</font></td></tr>  


<tr align="left"><td class=tablerow1><font face="宋体">表格调用样式&lt;table width=&quot;{$Width}&quot; border=&quot;0&quot; align=&quot;center&quot;   
  cellpadding=&quot;0&quot; cellspacing=&quot;0&quot; class=&quot;tableborder&quot;&gt;</font></td>         
</tr>


<tr align="left"><td class=tablerow1><font face="宋体">表格调用样式&lt;td height=&quot;25&quot; align=&quot;right&quot; class=&quot;tablebody&quot;&gt;</font></td></tr>   


<tr align="left"><td class=tablerow1><font face="宋体">显示繁体及菜单&lt;a href=&quot;{$InstallDir}index_gb.asp&quot;   
  class=navmenu&gt;首 页&lt;/a&gt; {$ChannelMenu} ┆ &lt;a name=&quot;StranLink&quot; style=&quot;color:red&quot;&gt;繁體中文&lt;/a&gt;</font>         
      <p><font face="宋体">其中有关菜单显示个数在模板通栏基本设置中修改</font></p>
      <p><font face="宋体">繁体加以下代码</font></p>
  <p><font face="宋体">&lt;script language=&quot;javascript&quot; src=&quot;{$InstallDir}inc/Std_StranJF.Js&quot;&gt;&lt;/script&gt;</font></p>  
</td>
</tr>


<tr align="left"><td class=tablerow1><font face="宋体">模板设计时替换路径：{$InstallDir}{$SkinPath}</font></td></tr>


<tr align="left"><td class=tablerow1><font face="宋体">&lt;a onclick=&quot;this.style.behavior='url(#default#homepage)';this.sethomepage('{$WebSiteUrl}');return               
	false;&quot; href=&quot;{$WebSiteUrl}&quot; title=&quot;将本站设为你的首页&quot;&gt;设为首页&lt;/a&gt;</font></td></tr>   


<tr align="left"><td class=tablerow1><font face="宋体">&lt;a href=&quot;javascript:window.external.AddFavorite(location.href,document.title)&quot;               
	title=&quot;将本站加入到你的收藏夹&quot;&gt;加入收藏&lt;/a&gt;</font></td></tr>


<tr align="left"><td class=tablerow1><font face="宋体">&lt;a href=&quot;mailto:{$MasterMail}&quot;&gt;联系我们E-MAIL&lt;/a&gt;</font></td></tr>  


<tr align="left"><td class=tablerow1><font face="宋体">您当前的位置：&lt;a  
    href=&quot;{$InstallDir}index_gb.asp&quot;&gt;{$WebSiteName}&lt;/a&gt;              
	-&amp;gt; 首页</font></td></tr>  





<tr align="left"><td class=tablerow1><font face="宋体">&lt;a href=&quot;{$InstallDir}user/logout.asp&quot;&gt;退出登录&lt;/a&gt;</font></td></tr> 





<tr align="left"><td class=tablerow1><font face="宋体">&lt;a href=&quot;{$InstallDir}user/&quot;&gt;用户管理&lt;/a&gt;</font></td></tr> 





<tr align="left"><td class=tablerow1><font face="宋体">&lt;marquee  
    scrollAmount=3&gt;{$ReadAnnounceList(0,12,22,1,1,2,0)}&lt;/marquee&gt;站内公告</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">&lt;marquee onmouseover=this.stop()  
    onmouseout=this.start() scrollAmount=1 scrollDelay=3 direction=up width=&quot;98%&quot; height=&quot;130&quot;              
	align=&quot;left&quot;&gt;{$ReadAnnounceList({$ChannelID},12,22,1,1,2,1)}&lt;/marquee&gt;某个频道站内公告</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">共有各项分类信息共{$ReadStatistic(1,{$ChannelID},0,0)}条<br>
      其中商情快讯共 {$ReadStatistic(1,{$ChannelID},21,0)}条<br>        
      其中招聘信息共  {$ReadStatistic(1,{$ChannelID},22,0)}条<br>        
      其中房产信息共  {$ReadStatistic(1,{$ChannelID},23,0)}条        
      </font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">&lt;iframe src=&quot;vote/vote.htm&quot; border=&quot;0&quot; width=&quot;100%&quot;              
	height=&quot;220&quot; frameborder=&quot;0&quot; scrolling=&quot;no&quot;&gt;&lt;/iframe&gt; 投票调查</font></td></tr>  





<tr align="left"><td class=tablerow1><font face="宋体">本月阅览排行{$ReadPopularArticle(1,0,3,22,12,0,_blank,·,showlist2)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">最新更新文章{$ReadArticleList(1,0,0,0,12,24,0,1,1,5,1,showlist)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">最新图文信息{$ReadArticlePic(1,0,0,0,4,4,12,0,120,90,1)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">用户排行{$ReadUserRank(0,0,10,showlist)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">热门阅读{$ReadArticleList(1,0,0,0,10,24,0,1,1,5,1,showlist)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">友情连接{$ReadFriendLink(24,8,3,1)}</font></td></tr>





<tr align="left"><td class=tablerow1>全文搜索
    <p><font face="宋体">&lt;form onsubmit=&quot;window.location=this.field.options[this.field.selectedIndex].value+this.keyword.value;              
	return false;&quot;&gt;<br>             
	&lt;td bgcolor=&quot;#EFEFEF&quot; height=&quot;25&quot; nowrap&gt;<br>             
	&lt;input name=&quot;keyword&quot; size=&quot;30&quot; value='关键字'  
    maxlength='50' onFocus='this.select();'&gt;            
	<br>
	&lt;select name=&quot;field&quot;&gt;<br>             
	&lt;option value=&quot;soft/search.asp?act=topic&amp;keyword=&quot;&gt;软件下载&lt;/option&gt;<br>            
	&lt;option value=&quot;article/search.asp?act=topic&amp;keyword=&quot;&gt;新闻资讯&lt;/option&gt;<br>            
	&lt;option value=&quot;flash/search.asp?act=topic&amp;keyword=&quot;&gt;FLASH搜索&lt;/option&gt;<br>            
	&lt;option value=&quot;article/search.asp?act=isWeb&amp;keyword=&quot;&gt;网页搜索&lt;/option&gt;<br>            
	&lt;/select&gt;<br>
	&lt;input name=&quot;Submit&quot; src=&quot;skin/default/d_search.gif&quot; type=&quot;image&quot;              
	value=&quot;Submit&quot; width=&quot;60&quot; height=&quot;20&quot; align=&quot;absmiddle&quot; border=&quot;0&quot;&gt;&lt;/td&gt;<br>             
	&lt;/form&gt;</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">&lt;!--专题菜单--&gt;&lt;script src=&quot;{$ChannelRootDir}js/specmenu.js&quot;              
	type=&quot;text/javascript&quot;&gt;&lt;/script&gt;</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">&lt;!--搜索表单--&gt;&lt;script src=&quot;{$ChannelRootDir}js/search.js&quot;              
	type=&quot;text/javascript&quot;&gt;&lt;/script&gt;</font></td></tr>





<tr align="left"><td class=tablerow1><font face="宋体">&lt;table width=&quot;100%&quot; border=&quot;0&quot;  
    cellspacing=&quot;0&quot; cellpadding=&quot;0&quot;&gt;<br>            
	&lt;td height=&quot;165&quot; valign=&quot;top&quot;&gt;&lt;div id=rolllink              
	style=overflow:hidden;height:165;width:180&gt;&lt;div id=rolllink1&gt;<br>             
	{$ReadFriendLink(20,1,1,0)}<br>
	&lt;table width=&quot;100%&quot; border=0 cellpadding=1 cellspacing=3 class=FriendLink1&gt;<br>             
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='申请友情连接'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='申请友情连接'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='申请友情连接'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='申请友情连接'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;/table&gt;&lt;/div&gt;&lt;div id=rolllink2&gt;&lt;/div&gt;&lt;/div&gt;<br>             
	&lt;script&gt;<br>
	var rollspeed=30<br>            
	rolllink2.innerHTML=rolllink1.innerHTML //克隆rolllink1为rolllink2<br>             
	function Marquee(){<br>             
	if(rolllink2.offsetTop-rolllink.scrollTop&lt;=0) //当滚动至rolllink1与rolllink2交界时<br>             
    rolllink.scrollTop-=rolllink1.offsetHeight //rolllink跳到最顶端<br>            
	else{<br>
    rolllink.scrollTop++<br>
	}<br>
	}<br>
	var MyMar=setInterval(Marquee,rollspeed) //设置定时器<br>             
    rolllink.onmouseover=function() {clearInterval(MyMar)}//鼠标移上时清除定时器达到滚动停止的目的<br>            
    rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//鼠标移开时重设定时器<br> 
	&lt;/script&gt;&lt;/td&gt;</font></td></tr>





<tr align="left"><td class=tablerow1>分类中有转向连接，合理利用转向连接，对于特殊页面无法使用的，可以另单独做一页面转向，对于有下属的可以使用#连接</td></tr>





<tr align="left"><td class=tablerow1></td></tr>





<th>函数式标签“()”中间是参数,用“,”分开</th>

<tr>
	<td class=tablerow1>{$CurrentStation( -&gt; )}&nbsp;&nbsp;&nbsp;&nbsp;当前位置“()”中间是分隔符</td>                    
</tr>
<tr>
	<td class=tablerow2>{$ReadFriendLink(24,8,1,1)}<br>&nbsp;&nbsp;友情连接标签,1、显示最多连接数，2、每行显示连接数，3、连接类型，1=LOGO连接，0=文字连接，4、排序方式，1=是升序，0=降序</td>
</tr>
<tr>
	<td class=tablerow1>{$ReadClassMenu(1,0,8,8,|,navbar)}<br>&nbsp;&nbsp;分类菜单标签，1、频道ID；2、分类ID，0=所有分类；3、显示多少分类名称；4、每行显示多少分类名称；5、每个分类名称中间的分隔符；6、调用样式名</td>
</tr>
<tr>
	<td class=tablerow2><font face="宋体">{$ReadClassMenu({$ChannelID},alltree,10,2,┣,0)}以树形方式显示所有的菜单</font></td>
</tr>
<tr>
	<td class=tablerow2>
	<p><font face="宋体">{$ReadClassMenu({$ChannelID},all,10,2,│,0)}以列表方式显示所有的菜单</font></p>
    </td>
</tr>
<tr>
	<td class=tablerow2>{$AnnounceContent()}<br>&nbsp;&nbsp;内容公告，如果()中间没有参数，就是最新内容公告，中间的参数是指定内容公告的标题。</td>
</tr>
<tr>
	<td class=tablerow1>{$ReadAnnounceList(0,12,22,1,1,2,1)}<br>&nbsp;&nbsp;列表公告标签：参数1、频道ID，0=所有频道，2、显示多少条公告，3、显示公告字符数，4、是否新窗口打开 1=是，0=否，5、是否显示时间 1=是，0=否                    
                                      6、时间模式，7、是否树型显示，1=是，0=否</td>
</tr>
<tr>
	<td class=tablerow2>{$ReadClassMenubar({$ChannelID},{$ClassID},35,2,· )}<br>&nbsp;&nbsp;列表分类菜单，参数，1、频道ID，2、分类ID，3、最多显示分类数，4、每行显示分类数，5、分类名称前面的符号</td>                    
</tr>
<tr>
	<td class=tablerow1>{$ReadPopularArticle(1,0,3,24,10,1,_blank,·,showlist)}<br>&nbsp;&nbsp;文章排行函数标签,参数海1、频道ID，2、分类ID，3、调用类型（0=所有热门，1=本日排行，2=本周排行,3=本月排行，4=推荐排行），                    
							   4、显示字符长度，5、显示多少行，6、是否显示点击数，7、连接目标，8、样式名称</td>
</tr>
<tr>
	<td class=tablerow2>{$ReadPopularSoft(2,0,0,24,10,1,_blank,·,showlist)}<br>&nbsp;&nbsp;软件排行函数标签,参数,1、频道ID，2、分类ID，3、调用类型（0=所有热门，1=本日排行，2=本周排行,3=本月排行，4=推荐排行），                    
							   4、显示字符长度，5、显示多少行，6、是否显示点击数，7、连接目标，8、样式名称</td>
</tr>
<tr>
	<td class=tablerow1>{$ReadSoftType(2,国产软件,24,10,1,1,5,1,showlist)}<br>&nbsp;&nbsp;软件类型函数标签，参数：1、频道ID，2、软件类型名称，3、显示字符长度，4、显示多少行，5、是否显示分类（1=是，0=否）                    
						      6、是否显示日期（1=显示，0=不显示） 7、显示日期模式 8、是否新窗口打开连接（1=新窗口，0=本窗口） 9、样式名称</td>                    
</tr>
<tr>
	<td class=tablerow2>{$ReadGuestList(12,22,1,1,5,showlist)}<br>&nbsp;&nbsp;留言列表标签；参数：1、显示列表数，2、显示字符数，3、是否新窗口打开（1=新窗口打开，0=本窗口打开），4、是否显示日期（1=是，0=否），5、日期格式，6、样式名称</td>
</tr>
</table>
<%
End Sub
Private Sub Label_ContentList()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan="2">生成<%=sModuleName%>模板标签</th>
</tr>
<form action="?action=list" method="post" name="myform" id="myform">
<tr>
	<td class=tablerow1 align="right">所属频道：</td>
	<td class=tablerow1><select name=ChannelID size=1 onchange="javascript:submit()">
<%
Set Rs = enchiasp.Execute("Select ChannelID,ChannelName,ModuleName from ECCMS_Channel where ChannelType < 2 And ChannelID <> 4 Order By ChannelID Asc")
Do While Not Rs.EOF
%>
		<option value='<%=Rs("ChannelID")%>'<%If Rs("ChannelID") = ChannelID Then Response.Write " selected"%>><%=Rs("ChannelName")%></option>
<%
	Rs.movenext
Loop
Set Rs = Nothing
%>
	</select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">所属分类：</td>
	<td class=tablerow2><select name="ClassID" id="ClassID">
          <option value="0">所有<%=sModuleName%>分类</option>
<%
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Response.Write strSelectClass
	Set Re = Nothing
%>
        </select><font color="#0066CC">&nbsp;&nbsp;如指定了分类,请选择类型分类</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">所属专题：</td>
	<td class=tablerow1><select name=SpecialID>
	<option value="0">不选择专题</option>
<%
If ChannelID <> 0 And ChannelID <> "" Then
	Set Rs = enchiasp.Execute("Select SpecialID,SpecialName from ECCMS_Special where ChannelID="& ChannelID &" And ChangeLink=0")
	Do While Not Rs.EOF
%>
		<option value='<%=Rs("SpecialID")%>'><%=Rs("SpecialName")%></option>
<%
		Rs.movenext
	Loop
	Set Rs = Nothing
End If
%>
	</select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">调用类型：</td>
	<td class=tablerow2><select name="ShowType" id="ShowType">
          <option value="0">所有最新<%=sModuleName%></option>
          <option value='1'>所有推荐<%=sModuleName%></option>
	  <option value='2'>所有热门<%=sModuleName%></option>
<%
If CInt(enchiasp.modules) = 1 Then
%>
	  <option value='3'>所有图文<%=sModuleName%></option>
	  <option value='4'>分类最新<%=sModuleName%></option>
	  <option value='5'>分类推荐<%=sModuleName%></option>
	  <option value='6'>分类热门<%=sModuleName%></option>
	  <option value='7'>分类图文<%=sModuleName%></option>
<%
Else
%>
	  <option value='3'>分类最新<%=sModuleName%></option>
	  <option value='4'>分类推荐<%=sModuleName%></option>
	  <option value='5'>分类热门<%=sModuleName%></option>
<%
End If
%>
        </select><font color="#0066CC">&nbsp;&nbsp;如选择了所有,分类选择无效</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">显示<%=sModuleName%>列表数：</td>
	<td class=tablerow1><input name="MaxListNum" type="text" id="MaxListNum" value="12" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">标题名称最多字符数：</td>
	<td class=tablerow2><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow1 align="right">是否显示分类名称：</td>
	<td class=tablerow1><select name="ShowClass" id="ShowClass">
          <option value='0'>不显示分类名称</option>
	  <option value='1'>显示分类名称</option>
        </select></td>
</tr>
<%
If CInt(enchiasp.modules) = 1 or CInt(enchiasp.modules) = 6 Then
%>
<tr>
	<td class=tablerow2 align="right">是否显示图文标题：</td>
	<td class=tablerow2><select name="ShowPic" id="ShowPic">
          <option value='1'>显示图文标题</option>
	  <option value='0'>不显示图文标题</option>
        </select></td>
</tr>
<%
End If
%>
<tr>
	<td class=tablerow1 align="right">是否显示日期：</td>
	<td class=tablerow1><select name="ShowDate" id="ShowDate">
          <option value='1'>显示日期</option>
	  <option value='0'>不显示日期</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">显示日期的格式：</td>
	<td class=tablerow2><select name="DateMode" id="DateMode">
<%
For i = 1 To 9
	Response.Write "<option value='" & i & "'"
	If i = 5 Then Response.Write " selected"
	Response.Write ">"
	Response.Write enchiasp.FormatDate(Now(),i)
	Response.Write "</option>" & vbCrLf
Next

%>
          
        </select></td>
</tr>
<tr>
	<td class=tablerow1 align="right">是否新窗口打开连接：</td>
	<td class=tablerow1><select name="newindow" id="newindow">
          <option value='0' selected>本窗口打开</option>
	  <option value='1'>新窗口打开</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">列表调用样式名称：</td>
	<td class=tablerow2><input name="StyleName" type="text" id="StyleName" value="showlist" size="15" maxlength="20"></td>
</tr>
<tr align=center>
	<td class=tablerow1 colspan="2"><input type=text name=SkinCode size=65 id="SkinCode"></td>
</tr>
<tr align=center>
	<td class=tablerow2 colspan="2">
		<input name="MakeJS" type="button" id="MakeJS" onclick="MakeCode();" value="&nbsp;生成<%=sModuleName%>模板标签&nbsp;" class=button> 
		<input name="Copy" type="button" id="Copy" value="&nbsp;复制到剪贴板&nbsp;" onclick="copy();" class=button>
		<input type="button" value="&nbsp;关闭窗口&nbsp;" onclick="selflabel();" class=button> 
		</td>
</tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function MakeCode()
{
if(document.myform.ClassID.value=="")
{
	alert("<%=sModuleName%>分类不能指定外部栏目！");
	document.myform.ClassID.focus();
	return false;
}
if(document.myform.MaxListNum.value=="")
{
	alert("显示<%=sModuleName%>列表数不能为空！");
	document.myform.MaxListNum.focus();
	return false;
}
if(document.myform.TitleMaxLen.value=="")
{
	alert("标题名称最多字符数不能为空！");
	document.myform.TitleMaxLen.focus();
	return false;
}
if(document.myform.StyleName.value=="")
{
	alert("列表样式名称不能为空！");
	document.myform.StyleName.focus();
	return false;
}
var strCode;
<%
Select Case CInt(enchiasp.modules)
	Case 1
		Response.Write "strCode=""{$ReadArticleList("";"
	Case 2
		Response.Write "strCode=""{$ReadSoftList("";"
	Case 3
		Response.Write "strCode=""{$ReadShopList("";"
	Case 5
		Response.Write "strCode=""{$ReadFlashList("";"
	Case Else
		Response.Write "strCode=""{$ReadArticleList("";"
End Select
%>
strCode+=document.myform.ChannelID.value+","
strCode+=document.myform.ClassID.value+","
strCode+=document.myform.SpecialID.value+","
strCode+=document.myform.ShowType.value+","
strCode+=document.myform.MaxListNum.value+","
strCode+=document.myform.TitleMaxLen.value+","
strCode+=document.myform.ShowClass.value+","
<%
If CInt(enchiasp.modules) = 1 or CInt(enchiasp.modules) = 6 Then
%>
strCode+=document.myform.ShowPic.value+","
<%
End If
%>
strCode+=document.myform.ShowDate.value+","
strCode+=document.myform.DateMode.value+","
strCode+=document.myform.newindow.value+","
strCode+=document.myform.StyleName.value
strCode+=")}";
document.myform.SkinCode.value=strCode;
}
</script>

<%
End Sub

Private Sub Label_ImageUse()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan="2">生成<%=sModuleName%>模板图片调用标签</th>
</tr>
<form action="?action=image" method="post" name="myform" id="myform">

<tr>
	<td class=tablerow1 align="right">所属频道：</td>
	<td class=tablerow1><select  name=ChannelID size=1 onchange="javascript:submit()">
<%
Set Rs = enchiasp.Execute("Select ChannelID,ChannelName,ModuleName From ECCMS_Channel where ChannelType < 2 And ChannelID <> 4 Order By ChannelID Asc")
Do While Not Rs.EOF
%>
		<option value='<%=Rs("ChannelID")%>'<%If Rs("ChannelID") = ChannelID Then Response.Write " selected"%>><%=Rs("ChannelName")%></option>
<%
	Rs.movenext
Loop
Set Rs = Nothing
%>
	</select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">所属分类：</td>
	<td class=tablerow2><select name="ClassID" id="ClassID">
          <option value="0">所有<%=sModuleName%>分类</option>
<%
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Response.Write strSelectClass
	Set Re = Nothing
%>
        </select><font color="#0066CC">&nbsp;&nbsp;如指定了分类,请选择类型分类</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">所属专题：</td>
	<td class=tablerow1><select name=SpecialID>
	<option value="0">不选择专题</option>
<%
If ChannelID <> 0 And ChannelID <> "" Then
	Set Rs = enchiasp.Execute("Select SpecialID,SpecialName from ECCMS_Special where ChannelID="& ChannelID &" And ChangeLink=0")
	Do While Not Rs.EOF
%>
		<option value='<%=Rs("SpecialID")%>'><%=Rs("SpecialName")%></option>
<%
		Rs.movenext
	Loop
	Set Rs = Nothing
End If
%>
	</select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">调用类型：</td>
	<td class=tablerow2><select name="ShowType" id="ShowType">
          <option value="0">所有最新<%=sModuleName%>图片</option>
          <option value='1'>所有推荐<%=sModuleName%>图片</option>
	  <option value='2'>所有热门<%=sModuleName%>图片</option>
	  <option value='3'>分类最新<%=sModuleName%>图片</option>
	  <option value='4'>分类推荐<%=sModuleName%>图片</option>
	  <option value='5'>分类热门<%=sModuleName%>图片</option>
        </select><font color="#0066CC">&nbsp;&nbsp;如选择了所有,分类选择无效</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">最多显示多少张<%=sModuleName%>图片：</td>
	<td class=tablerow1><input name="MaxPicNum" type="text" id="MaxPicNum" value="12" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">每行显示多少张<%=sModuleName%>图片：</td>
	<td class=tablerow2><input name="PerRowNum" type="text" id="PerRowNum" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow1 align="right">标题名称最多字符数：</td>
	<td class=tablerow1><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">是否新窗口打开连接：</td>
	<td class=tablerow2><select name="newindow" id="newindow">
          <option value='0' selected>本窗口打开</option>
	  <option value='1'>新窗口打开</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow1 align="right">图片的宽度：</td>
	<td class=tablerow1><input name="picwidth" type="text" id="picwidth" value="120" size="5" maxlength="3"> 像素</td>
</tr>
<tr>
	<td class=tablerow2 align="right">图片的高度：</td>
	<td class=tablerow2><input name="picheight" type="text" id="picheight" value="100" size="5" maxlength="3"> 像素</td>
</tr>
<tr>
	<td class=tablerow1 align="right">是否显示<%=sModuleName%>标题名称：</td>
	<td class=tablerow1><select name="showtopic" id="showtopic">
          <option value='1'>显示</option>
	  <option value='0'>不显示</option>
	  <option value='2'>特殊显示</option>
        </select>对于特殊显示可以使用2</td>
</tr>
<tr align=center>
	<td class=tablerow2 colspan="2"><input type=text name=SkinCode size=65 id="SkinCode"></td>
</tr>
<tr align=center>
	<td class=tablerow1 colspan="2">
		<input name="MakeJS" type="button" id="MakeJS" onclick="MakeCode();" value="&nbsp;生成<%=sModuleName%>模板标签&nbsp;" class=button> 
		<input name="Copy" type="button" id="Copy" value="&nbsp;复制到剪贴板&nbsp;" onclick="copy();" class=button>
		<input type="reset" name="button" value="&nbsp;关闭窗口&nbsp;" onclick="copy(); window.close()" class=button></td>
</tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function MakeCode()
{
if(document.myform.ClassID.value=="")
{
	alert("<%=sModuleName%>分类不能指定外部栏目！");
	document.myform.ClassID.focus();
	return false;
}
if(document.myform.MaxPicNum.value=="")
{
	alert("显示<%=sModuleName%>图片数不能为空！");
	document.myform.MaxPicNum.focus();
	return false;
}
if(document.myform.TitleMaxLen.value=="")
{
	alert("标题名称最多字符数不能为空！");
	document.myform.TitleMaxLen.focus();
	return false;
}
var strCode;
<%
Select Case CInt(enchiasp.modules)
	Case 1
		Response.Write "strCode=""{$ReadArticlePic("";"
	Case 2
		Response.Write "strCode=""{$ReadSoftPic("";"
	Case 3
		Response.Write "strCode=""{$ReadShopPic("";"
	Case 5
		Response.Write "strCode=""{$ReadFlashPic("";"
	Case Else
		Response.Write "strCode=""{$ReadArticlePic("";"
End Select
%>
strCode+=document.myform.ChannelID.value+","
strCode+=document.myform.ClassID.value+","
strCode+=document.myform.SpecialID.value+","
strCode+=document.myform.ShowType.value+","
strCode+=document.myform.MaxPicNum.value+","
strCode+=document.myform.PerRowNum.value+","
strCode+=document.myform.TitleMaxLen.value+","
strCode+=document.myform.newindow.value+","
strCode+=document.myform.picwidth.value+","
strCode+=document.myform.picheight.value+","
strCode+=document.myform.showtopic.value
strCode+=")}";
document.myform.SkinCode.value=strCode;
}
</script>

<%	
End Sub
Private Sub Label_PicAndText()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan="2">生成<%=sModuleName%>模板图文混排标签</th>
</tr>
<form action="?action=list" method="post" name="myform" id="myform">
<tr>
	<td class=tablerow1 align="right">所属频道：</td>
	<td class=tablerow1><select name=ChannelID size=1 onchange="javascript:submit()">
<%
Set Rs = enchiasp.Execute("Select ChannelID,ChannelName,ModuleName from ECCMS_Channel where ChannelType < 2 And ChannelID <> 4 Order By ChannelID Asc")
Do While Not Rs.EOF
%>
		<option value='<%=Rs("ChannelID")%>'<%If Rs("ChannelID") = ChannelID Then Response.Write " selected"%>><%=Rs("ChannelName")%></option>
<%
	Rs.movenext
Loop
Set Rs = Nothing
%>
	</select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">所属分类：</td>
	<td class=tablerow2><select name="ClassID" id="ClassID">
          <option value="0">所有<%=sModuleName%>分类</option>
<%
	Dim strSelectClass,re
	strSelectClass = enchiasp.LoadSelectClass(ChannelID)
	Set re = New RegExp
        re.IgnoreCase = True
        re.Global = True
	Re.Pattern="(value=)(.*)("" )"
	strSelectClass = Re.Replace(strSelectClass,"")
	Re.Pattern="({ClassID=)(.*)(}>)"
	strSelectClass = Re.Replace(strSelectClass,"value=""$2"">")
	Response.Write strSelectClass
	Set Re = Nothing
%>
        </select><font color="#0066CC">&nbsp;&nbsp;如指定了分类,请选择类型分类</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">所属专题：</td>
	<td class=tablerow1><select name=SpecialID>
	<option value="0">不选择专题</option>
<%
If ChannelID <> 0 And ChannelID <> "" Then
	Set Rs = enchiasp.Execute("Select SpecialID,SpecialName from ECCMS_Special where ChannelID="& ChannelID &" And ChangeLink=0")
	Do While Not Rs.EOF
%>
		<option value='<%=Rs("SpecialID")%>'><%=Rs("SpecialName")%></option>
<%
		Rs.movenext
	Loop
	Set Rs = Nothing
End If
%>
	</select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">调用类型：</td>
	<td class=tablerow2><select name="ShowType" id="ShowType">
          <option value="0">所有最新<%=sModuleName%></option>
          <option value='1'>所有推荐<%=sModuleName%></option>
	  <option value='2'>所有热门<%=sModuleName%></option>
<%
If CInt(enchiasp.modules) = 1 Then
%>
	  <option value='3'>所有图文<%=sModuleName%></option>
	  <option value='4'>分类最新<%=sModuleName%></option>
	  <option value='5'>分类推荐<%=sModuleName%></option>
	  <option value='6'>分类热门<%=sModuleName%></option>
	  <option value='7'>分类图文<%=sModuleName%></option>
<%
Else
%>
	  <option value='3'>分类最新<%=sModuleName%></option>
	  <option value='4'>分类推荐<%=sModuleName%></option>
	  <option value='5'>分类热门<%=sModuleName%></option>
<%
End If
%>
        </select><font color="#0066CC">&nbsp;&nbsp;如选择了所有,分类选择无效</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">显示<%=sModuleName%>列表数：</td>
	<td class=tablerow1><input name="MaxListNum" type="text" id="MaxListNum" value="12" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">标题名称最多字符数：</td>
	<td class=tablerow2><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow1 align="right">是否显示分类名称：</td>
	<td class=tablerow1><select name="ShowClass" id="ShowClass">
          <option value='0'>不显示分类名称</option>
	  <option value='1'>显示分类名称</option>
        </select></td>
</tr>
<%
If CInt(enchiasp.modules) = 1 Then
%>
<tr>
	<td class=tablerow2 align="right">是否显示图文标题：</td>
	<td class=tablerow2><select name="ShowPic" id="ShowPic">
          <option value='1'>显示图文标题</option>
	  <option value='0'>不显示图文标题</option>
        </select></td>
</tr>
<%
End If
%>
<tr>
	<td class=tablerow1 align="right">是否显示日期：</td>
	<td class=tablerow1><select name="ShowDate" id="ShowDate">
          <option value='1'>显示日期</option>
	  <option value='0'>不显示日期</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">显示日期的格式：</td>
	<td class=tablerow2><select name="DateMode" id="DateMode">
<%
For i = 1 To 9
	Response.Write "<option value='" & i & "'"
	If i = 5 Then Response.Write " selected"
	Response.Write ">"
	Response.Write enchiasp.FormatDate(Now(),i)
	Response.Write "</option>" & vbCrLf
Next

%>
          
        </select></td>
</tr>
<tr>
	<td class=tablerow1 align="right">是否新窗口打开连接：</td>
	<td class=tablerow1><select name="newindow" id="newindow">
          <option value='0' selected>本窗口打开</option>
	  <option value='1'>新窗口打开</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">列表调用样式名称：</td>
	<td class=tablerow2><input name="StyleName" type="text" id="StyleName" value="showlist" size="15" maxlength="20"></td>
</tr>
<tr align=center>
	<td class=tablerow1 colspan="2"><input type=text name=SkinCode size=65 id="SkinCode"></td>
</tr>
<tr align=center>
	<td class=tablerow2 colspan="2">
		<input name="MakeJS" type="button" id="MakeJS" onclick="MakeCode();" value="&nbsp;生成<%=sModuleName%>模板标签&nbsp;" class=button> 
		<input name="Copy" type="button" id="Copy" value="&nbsp;复制到剪贴板&nbsp;" onclick="copy();" class=button>
		<input type="button" value="&nbsp;关闭窗口&nbsp;" onclick="selflabel();" class=button> 
		</td>
</tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function MakeCode()
{
if(document.myform.ClassID.value=="")
{
	alert("<%=sModuleName%>分类不能指定外部栏目！");
	document.myform.ClassID.focus();
	return false;
}
if(document.myform.MaxListNum.value=="")
{
	alert("显示<%=sModuleName%>列表数不能为空！");
	document.myform.MaxListNum.focus();
	return false;
}
if(document.myform.TitleMaxLen.value=="")
{
	alert("标题名称最多字符数不能为空！");
	document.myform.TitleMaxLen.focus();
	return false;
}
if(document.myform.StyleName.value=="")
{
	alert("列表样式名称不能为空！");
	document.myform.StyleName.focus();
	return false;
}
var strCode;
<%
Select Case CInt(enchiasp.modules)
	Case 1
		Response.Write "strCode=""{$ReadArticleList("";"
	Case 2
		Response.Write "strCode=""{$ReadSoftList("";"
	Case 3
		Response.Write "strCode=""{$ReadShopList("";"
	Case 5
		Response.Write "strCode=""{$ReadFlashList("";"
	Case Else
		Response.Write "strCode=""{$ReadArticleList("";"
End Select
%>
strCode+=document.myform.ChannelID.value+","
strCode+=document.myform.ClassID.value+","
strCode+=document.myform.SpecialID.value+","
strCode+=document.myform.ShowType.value+","
strCode+=document.myform.MaxListNum.value+","
strCode+=document.myform.TitleMaxLen.value+","
strCode+=document.myform.ShowClass.value+","
<%
If CInt(enchiasp.modules) = 1 Then
%>
strCode+=document.myform.ShowPic.value+","
<%
End If
%>
strCode+=document.myform.ShowDate.value+","
strCode+=document.myform.DateMode.value+","
strCode+=document.myform.newindow.value+","
strCode+=document.myform.StyleName.value
strCode+=")}";
document.myform.SkinCode.value=strCode;
}
</script>

<%
End Sub
%>