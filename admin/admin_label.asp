<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<%
Admin_header
Response.Write "<base target=""_self"">" & vbNewLine
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
	<td class=tablerow1>�뽫���ϱ�ǩ���Ƶ�ģ����Ӧ��λ��</td>
</tr>
</table>
<%
CloseConn
Private Sub showmain()
	%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th>����ģ���ǩ</th>
</tr>
<tr>
	<td class=tablerow1>{$showuserinfo}&nbsp;&nbsp;&nbsp;&nbsp;��ʾ��Ա��½״̬(���ģ����ģ�泣��������27�޸�)</td> 
	             
</tr>




<tr>
	<td class=tablerow1>{$InstallDir}&nbsp;&nbsp;&nbsp;&nbsp;ϵͳ��װ·��  ��ϵͳ�Զ����ɣ�</td>       
	             
</tr>
<tr>
	<td class=tablerow2>{$SkinPath}&nbsp;&nbsp;&nbsp;&nbsp;Ƥ��ͼƬ·��</td>
</tr>
<tr>
	<td class=tablerow1>{$ChannelRootDir}&nbsp;&nbsp;&nbsp;&nbsp;Ƶ��Ŀ¼·��</td>
</tr>
<tr>
	<td class=tablerow2>{$Version}&nbsp;&nbsp;&nbsp;&nbsp;ϵͳ�汾��Ϣ</td>
</tr>
<tr>
	<td class=tablerow1>{$WebSiteName}&nbsp;&nbsp;&nbsp;&nbsp;��վ���� ���ڻ��������޸ģ�</td>                    
</tr>
<tr>
	<td class=tablerow2>{$WebSiteUrl}&nbsp;&nbsp;&nbsp;&nbsp;��վURL ���ڻ��������޸ģ�</td>                    
</tr>
<tr>
	<td class=tablerow1>{$MasterMail}&nbsp;&nbsp;&nbsp;&nbsp;����ԱE-Mail���ڻ��������޸ģ�</td>
</tr>
<tr>
	<td class=tablerow2>{$Keyword}&nbsp;&nbsp;&nbsp;&nbsp;��վ�ؼ��� ���ڻ��������޸ģ�</td>                    
</tr>
<tr>
	<td class=tablerow1>{$Copyright}&nbsp;&nbsp;&nbsp;&nbsp;��վ��Ȩ��Ϣ ���ڻ��������޸ģ�</td>                    
</tr>
<tr>
	<td class=tablerow2>{$Width}&nbsp;&nbsp;&nbsp;&nbsp;����������� </td>
</tr>
<tr>
	<td class=tablerow1>{$IndexPage}&nbsp;&nbsp;&nbsp;&nbsp;Ĭ����ҳ�ļ���</td>
</tr>
<tr>
	<td class=tablerow2>{$Style_CSS}&nbsp;&nbsp;&nbsp;&nbsp;CSS��ʽ</td>
</tr>
<tr>
	<td class=tablerow1>{$PageTitle}&nbsp;&nbsp;&nbsp;&nbsp;HTML�ļ�����</td>
</tr>
<tr>
	<td class=tablerow2>{$TotalStatistics}&nbsp;&nbsp;&nbsp;&nbsp;ϵͳ��ͳ��</td>
</tr>
<tr>
	<td class=tablerow1>{$RenewStatistics}&nbsp;&nbsp;&nbsp;&nbsp;������Ϣͳ��</td>
</tr>
<tr>
	<td class=tablerow2>{$ChannelMenu}&nbsp;&nbsp;&nbsp;&nbsp;����Ƶ���˵���ǩ</td>
</tr>
<tr>
	<td class=tablerow2>{$ShowHotArticle}&nbsp;&nbsp;&nbsp;&nbsp;�������ŵ��</td>
</tr>
<tr>
	<th>�ű��������ǩ����</th>
</tr>

<tr>
	<td class=tablerow1>&lt;script language=&quot;javascript&quot;       
      type=&quot;text/javascript&quot; src=&quot;{$WebSiteUrl}{$InstallDir}qqonline/qq.asp&quot;&gt;&lt;/script&gt;       
      ����QQ������ͨ��ģ�������ӣ�ע���еĵط������޷����ã������ڸ�����Ŀ��ʹ�ã��ڸ�Ŀ¼�¿��ܻ�����޷����õ�������޸�QQ����������������</td>                
</tr>

<tr>
	<td class=tablerow1>{$vod} Ŀǰ��֧��MEDIA      
      PLAY��ʽ�ļ����޸���ģ�峣������26���޸�</td>                
</tr>

<tr>
	<td class=tablerow1>{$tupianhuan}    
      ͼƬ����������÷���������Ҫ���õĵط����ر�ǩ,���Ҫ�޸ĸ�FLASH��ͼƬ��С�Ȳ�������ͨ��ģ��<font COLOR="#800000" face="����">��������</font>��24�޸ġ�</td>                
</tr>

<tr>
	<td class=tablerow1>{$dibuhuan}    
      ͼƬ���ҹ���������÷���������Ҫ���õĵط����ر�ǩ,���Ҫ�޸ĸ�FLASH��ͼƬ��С�Ȳ�������ͨ��ģ�����������25�޸ġ���������10��ͼƬ������ʹ��JPGͼƬ</td>                
</tr>

<tr>
	<td class=tablerow1><font face="����">  
    &lt;script src=&quot;{$WebSiteUrl}{$InstallDir}count/count.asp&quot;&gt;&lt;/script><br>        
    ��ͨ����Ϊҳ��ͳ��</font></td>                
</tr>

<tr>
	<td class=tablerow1>{$ReadShopPic(3,0,0,3,4,4,23,0,100,100,2)}   
      �̳�ͼƬ������ʾ����ʾ������ͨ��ģ��<font COLOR="#800000" face="����">��������</font>��28�޸ġ����һ��Ϊ2</td>                
</tr>

<tr>
	<td class=tablerow1>{$ReadArticlePic(1,0,0,0,1,1,120,0,105,79,2)} ����ͼƬ������ʾ����ʾ������ͨ��ģ��<font COLOR="#800000" face="����">��������</font>��29�޸ġ����һ��Ϊ2</td>                  
</tr>

<tr>
	<td class=tablerow1>�����������Ҫ������ʾ�����ڸ�ģ�����޸�����</td>                
</tr>

<tr>
	<td class=tablerow1><font face="����"><b>     
    &lt;head&gt;&lt;/head&gt;�м�ͨ�ô���</b></font>
      <p><font face="����">&lt;meta http-equiv=&quot;Content-Type&quot; content=&quot;text/html;   
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
  <p align="left"><font face="����">���λ��JS�ļ� &lt;script language=javascript   
  src={$InstallDir}adfile/banner.js&gt;&lt;/script&gt;</font></td></tr>


<tr align="left"><td class=tablerow1>
  <p align="left"><font face="����">�˵�JS�ļ�&lt;script src=&quot;{$InstallDir}inc/menu.js&quot; type=&quot;text/javascript&quot;&gt;&lt;/script&gt;</font></td></tr>  


<tr align="left"><td class=tablerow1><font face="����">��������ʽ&lt;table width=&quot;{$Width}&quot; border=&quot;0&quot; align=&quot;center&quot;   
  cellpadding=&quot;0&quot; cellspacing=&quot;0&quot; class=&quot;tableborder&quot;&gt;</font></td>         
</tr>


<tr align="left"><td class=tablerow1><font face="����">��������ʽ&lt;td height=&quot;25&quot; align=&quot;right&quot; class=&quot;tablebody&quot;&gt;</font></td></tr>   


<tr align="left"><td class=tablerow1><font face="����">��ʾ���弰�˵�&lt;a href=&quot;{$InstallDir}index_gb.asp&quot;   
  class=navmenu&gt;�� ҳ&lt;/a&gt; {$ChannelMenu} �� &lt;a name=&quot;StranLink&quot; style=&quot;color:red&quot;&gt;���w����&lt;/a&gt;</font>         
      <p><font face="����">�����йز˵���ʾ������ģ��ͨ�������������޸�</font></p>
      <p><font face="����">��������´���</font></p>
  <p><font face="����">&lt;script language=&quot;javascript&quot; src=&quot;{$InstallDir}inc/Std_StranJF.Js&quot;&gt;&lt;/script&gt;</font></p>  
</td>
</tr>


<tr align="left"><td class=tablerow1><font face="����">ģ�����ʱ�滻·����{$InstallDir}{$SkinPath}</font></td></tr>


<tr align="left"><td class=tablerow1><font face="����">&lt;a onclick=&quot;this.style.behavior='url(#default#homepage)';this.sethomepage('{$WebSiteUrl}');return               
	false;&quot; href=&quot;{$WebSiteUrl}&quot; title=&quot;����վ��Ϊ�����ҳ&quot;&gt;��Ϊ��ҳ&lt;/a&gt;</font></td></tr>   


<tr align="left"><td class=tablerow1><font face="����">&lt;a href=&quot;javascript:window.external.AddFavorite(location.href,document.title)&quot;               
	title=&quot;����վ���뵽����ղؼ�&quot;&gt;�����ղ�&lt;/a&gt;</font></td></tr>


<tr align="left"><td class=tablerow1><font face="����">&lt;a href=&quot;mailto:{$MasterMail}&quot;&gt;��ϵ����E-MAIL&lt;/a&gt;</font></td></tr>  


<tr align="left"><td class=tablerow1><font face="����">����ǰ��λ�ã�&lt;a  
    href=&quot;{$InstallDir}index_gb.asp&quot;&gt;{$WebSiteName}&lt;/a&gt;              
	-&amp;gt; ��ҳ</font></td></tr>  





<tr align="left"><td class=tablerow1><font face="����">&lt;a href=&quot;{$InstallDir}user/logout.asp&quot;&gt;�˳���¼&lt;/a&gt;</font></td></tr> 





<tr align="left"><td class=tablerow1><font face="����">&lt;a href=&quot;{$InstallDir}user/&quot;&gt;�û�����&lt;/a&gt;</font></td></tr> 





<tr align="left"><td class=tablerow1><font face="����">&lt;marquee  
    scrollAmount=3&gt;{$ReadAnnounceList(0,12,22,1,1,2,0)}&lt;/marquee&gt;վ�ڹ���</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">&lt;marquee onmouseover=this.stop()  
    onmouseout=this.start() scrollAmount=1 scrollDelay=3 direction=up width=&quot;98%&quot; height=&quot;130&quot;              
	align=&quot;left&quot;&gt;{$ReadAnnounceList({$ChannelID},12,22,1,1,2,1)}&lt;/marquee&gt;ĳ��Ƶ��վ�ڹ���</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">���и��������Ϣ��{$ReadStatistic(1,{$ChannelID},0,0)}��<br>
      ���������Ѷ�� {$ReadStatistic(1,{$ChannelID},21,0)}��<br>        
      ������Ƹ��Ϣ��  {$ReadStatistic(1,{$ChannelID},22,0)}��<br>        
      ���з�����Ϣ��  {$ReadStatistic(1,{$ChannelID},23,0)}��        
      </font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">&lt;iframe src=&quot;vote/vote.htm&quot; border=&quot;0&quot; width=&quot;100%&quot;              
	height=&quot;220&quot; frameborder=&quot;0&quot; scrolling=&quot;no&quot;&gt;&lt;/iframe&gt; ͶƱ����</font></td></tr>  





<tr align="left"><td class=tablerow1><font face="����">������������{$ReadPopularArticle(1,0,3,22,12,0,_blank,��,showlist2)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">���¸�������{$ReadArticleList(1,0,0,0,12,24,0,1,1,5,1,showlist)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">����ͼ����Ϣ{$ReadArticlePic(1,0,0,0,4,4,12,0,120,90,1)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">�û�����{$ReadUserRank(0,0,10,showlist)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">�����Ķ�{$ReadArticleList(1,0,0,0,10,24,0,1,1,5,1,showlist)}</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">��������{$ReadFriendLink(24,8,3,1)}</font></td></tr>





<tr align="left"><td class=tablerow1>ȫ������
    <p><font face="����">&lt;form onsubmit=&quot;window.location=this.field.options[this.field.selectedIndex].value+this.keyword.value;              
	return false;&quot;&gt;<br>             
	&lt;td bgcolor=&quot;#EFEFEF&quot; height=&quot;25&quot; nowrap&gt;<br>             
	&lt;input name=&quot;keyword&quot; size=&quot;30&quot; value='�ؼ���'  
    maxlength='50' onFocus='this.select();'&gt;            
	<br>
	&lt;select name=&quot;field&quot;&gt;<br>             
	&lt;option value=&quot;soft/search.asp?act=topic&amp;keyword=&quot;&gt;�������&lt;/option&gt;<br>            
	&lt;option value=&quot;article/search.asp?act=topic&amp;keyword=&quot;&gt;������Ѷ&lt;/option&gt;<br>            
	&lt;option value=&quot;flash/search.asp?act=topic&amp;keyword=&quot;&gt;FLASH����&lt;/option&gt;<br>            
	&lt;option value=&quot;article/search.asp?act=isWeb&amp;keyword=&quot;&gt;��ҳ����&lt;/option&gt;<br>            
	&lt;/select&gt;<br>
	&lt;input name=&quot;Submit&quot; src=&quot;skin/default/d_search.gif&quot; type=&quot;image&quot;              
	value=&quot;Submit&quot; width=&quot;60&quot; height=&quot;20&quot; align=&quot;absmiddle&quot; border=&quot;0&quot;&gt;&lt;/td&gt;<br>             
	&lt;/form&gt;</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">&lt;!--ר��˵�--&gt;&lt;script src=&quot;{$ChannelRootDir}js/specmenu.js&quot;              
	type=&quot;text/javascript&quot;&gt;&lt;/script&gt;</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">&lt;!--������--&gt;&lt;script src=&quot;{$ChannelRootDir}js/search.js&quot;              
	type=&quot;text/javascript&quot;&gt;&lt;/script&gt;</font></td></tr>





<tr align="left"><td class=tablerow1><font face="����">&lt;table width=&quot;100%&quot; border=&quot;0&quot;  
    cellspacing=&quot;0&quot; cellpadding=&quot;0&quot;&gt;<br>            
	&lt;td height=&quot;165&quot; valign=&quot;top&quot;&gt;&lt;div id=rolllink              
	style=overflow:hidden;height:165;width:180&gt;&lt;div id=rolllink1&gt;<br>             
	{$ReadFriendLink(20,1,1,0)}<br>
	&lt;table width=&quot;100%&quot; border=0 cellpadding=1 cellspacing=3 class=FriendLink1&gt;<br>             
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='������������'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='������������'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='������������'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;tr&gt;<br>
	&lt;td align=center class=FriendLink2&gt;&lt;a href='{$InstallDir}link/'              
	target=_blank title='������������'&gt;&lt;img src='{$InstallDir}images/link.gif'              
	width=88 height=31 border=0&gt;&lt;/a&gt;&lt;/td&gt;<br>             
	&lt;/tr&gt;<br>
	&lt;/table&gt;&lt;/div&gt;&lt;div id=rolllink2&gt;&lt;/div&gt;&lt;/div&gt;<br>             
	&lt;script&gt;<br>
	var rollspeed=30<br>            
	rolllink2.innerHTML=rolllink1.innerHTML //��¡rolllink1Ϊrolllink2<br>             
	function Marquee(){<br>             
	if(rolllink2.offsetTop-rolllink.scrollTop&lt;=0) //��������rolllink1��rolllink2����ʱ<br>             
    rolllink.scrollTop-=rolllink1.offsetHeight //rolllink�������<br>            
	else{<br>
    rolllink.scrollTop++<br>
	}<br>
	}<br>
	var MyMar=setInterval(Marquee,rollspeed) //���ö�ʱ��<br>             
    rolllink.onmouseover=function() {clearInterval(MyMar)}//�������ʱ�����ʱ���ﵽ����ֹͣ��Ŀ��<br>            
    rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//����ƿ�ʱ���趨ʱ��<br> 
	&lt;/script&gt;&lt;/td&gt;</font></td></tr>





<tr align="left"><td class=tablerow1>��������ת�����ӣ���������ת�����ӣ���������ҳ���޷�ʹ�õģ�����������һҳ��ת�򣬶����������Ŀ���ʹ��#����</td></tr>





<tr align="left"><td class=tablerow1></td></tr>





<th>����ʽ��ǩ��()���м��ǲ���,�á�,���ֿ�</th>

<tr>
	<td class=tablerow1>{$CurrentStation( -&gt; )}&nbsp;&nbsp;&nbsp;&nbsp;��ǰλ�á�()���м��Ƿָ���</td>                    
</tr>
<tr>
	<td class=tablerow2>{$ReadFriendLink(24,8,1,1)}<br>&nbsp;&nbsp;�������ӱ�ǩ,1����ʾ�����������2��ÿ����ʾ��������3���������ͣ�1=LOGO���ӣ�0=�������ӣ�4������ʽ��1=������0=����</td>
</tr>
<tr>
	<td class=tablerow1>{$ReadClassMenu(1,0,8,8,|,navbar)}<br>&nbsp;&nbsp;����˵���ǩ��1��Ƶ��ID��2������ID��0=���з��ࣻ3����ʾ���ٷ������ƣ�4��ÿ����ʾ���ٷ������ƣ�5��ÿ�����������м�ķָ�����6��������ʽ��</td>
</tr>
<tr>
	<td class=tablerow2><font face="����">{$ReadClassMenu({$ChannelID},alltree,10,2,��,0)}�����η�ʽ��ʾ���еĲ˵�</font></td>
</tr>
<tr>
	<td class=tablerow2>
	<p><font face="����">{$ReadClassMenu({$ChannelID},all,10,2,��,0)}���б�ʽ��ʾ���еĲ˵�</font></p>
    </td>
</tr>
<tr>
	<td class=tablerow2>{$AnnounceContent()}<br>&nbsp;&nbsp;���ݹ��棬���()�м�û�в����������������ݹ��棬�м�Ĳ�����ָ�����ݹ���ı��⡣</td>
</tr>
<tr>
	<td class=tablerow1>{$ReadAnnounceList(0,12,22,1,1,2,1)}<br>&nbsp;&nbsp;�б����ǩ������1��Ƶ��ID��0=����Ƶ����2����ʾ���������棬3����ʾ�����ַ�����4���Ƿ��´��ڴ� 1=�ǣ�0=��5���Ƿ���ʾʱ�� 1=�ǣ�0=��                    
                                      6��ʱ��ģʽ��7���Ƿ�������ʾ��1=�ǣ�0=��</td>
</tr>
<tr>
	<td class=tablerow2>{$ReadClassMenubar({$ChannelID},{$ClassID},35,2,�� )}<br>&nbsp;&nbsp;�б����˵���������1��Ƶ��ID��2������ID��3�������ʾ��������4��ÿ����ʾ��������5����������ǰ��ķ���</td>                    
</tr>
<tr>
	<td class=tablerow1>{$ReadPopularArticle(1,0,3,24,10,1,_blank,��,showlist)}<br>&nbsp;&nbsp;�������к�����ǩ,������1��Ƶ��ID��2������ID��3���������ͣ�0=�������ţ�1=�������У�2=��������,3=�������У�4=�Ƽ����У���                    
							   4����ʾ�ַ����ȣ�5����ʾ�����У�6���Ƿ���ʾ�������7������Ŀ�꣬8����ʽ����</td>
</tr>
<tr>
	<td class=tablerow2>{$ReadPopularSoft(2,0,0,24,10,1,_blank,��,showlist)}<br>&nbsp;&nbsp;������к�����ǩ,����,1��Ƶ��ID��2������ID��3���������ͣ�0=�������ţ�1=�������У�2=��������,3=�������У�4=�Ƽ����У���                    
							   4����ʾ�ַ����ȣ�5����ʾ�����У�6���Ƿ���ʾ�������7������Ŀ�꣬8����ʽ����</td>
</tr>
<tr>
	<td class=tablerow1>{$ReadSoftType(2,�������,24,10,1,1,5,1,showlist)}<br>&nbsp;&nbsp;������ͺ�����ǩ��������1��Ƶ��ID��2������������ƣ�3����ʾ�ַ����ȣ�4����ʾ�����У�5���Ƿ���ʾ���ࣨ1=�ǣ�0=��                    
						      6���Ƿ���ʾ���ڣ�1=��ʾ��0=����ʾ�� 7����ʾ����ģʽ 8���Ƿ��´��ڴ����ӣ�1=�´��ڣ�0=�����ڣ� 9����ʽ����</td>                    
</tr>
<tr>
	<td class=tablerow2>{$ReadGuestList(12,22,1,1,5,showlist)}<br>&nbsp;&nbsp;�����б��ǩ��������1����ʾ�б�����2����ʾ�ַ�����3���Ƿ��´��ڴ򿪣�1=�´��ڴ򿪣�0=�����ڴ򿪣���4���Ƿ���ʾ���ڣ�1=�ǣ�0=�񣩣�5�����ڸ�ʽ��6����ʽ����</td>
</tr>
</table>
<%
End Sub
Private Sub Label_ContentList()
%>
<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th colspan="2">����<%=sModuleName%>ģ���ǩ</th>
</tr>
<form action="?action=list" method="post" name="myform" id="myform">
<tr>
	<td class=tablerow1 align="right">����Ƶ����</td>
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
	<td class=tablerow2 align="right">�������ࣺ</td>
	<td class=tablerow2><select name="ClassID" id="ClassID">
          <option value="0">����<%=sModuleName%>����</option>
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
        </select><font color="#0066CC">&nbsp;&nbsp;��ָ���˷���,��ѡ�����ͷ���</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">����ר�⣺</td>
	<td class=tablerow1><select name=SpecialID>
	<option value="0">��ѡ��ר��</option>
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
	<td class=tablerow2 align="right">�������ͣ�</td>
	<td class=tablerow2><select name="ShowType" id="ShowType">
          <option value="0">��������<%=sModuleName%></option>
          <option value='1'>�����Ƽ�<%=sModuleName%></option>
	  <option value='2'>��������<%=sModuleName%></option>
<%
If CInt(enchiasp.modules) = 1 Then
%>
	  <option value='3'>����ͼ��<%=sModuleName%></option>
	  <option value='4'>��������<%=sModuleName%></option>
	  <option value='5'>�����Ƽ�<%=sModuleName%></option>
	  <option value='6'>��������<%=sModuleName%></option>
	  <option value='7'>����ͼ��<%=sModuleName%></option>
<%
Else
%>
	  <option value='3'>��������<%=sModuleName%></option>
	  <option value='4'>�����Ƽ�<%=sModuleName%></option>
	  <option value='5'>��������<%=sModuleName%></option>
<%
End If
%>
        </select><font color="#0066CC">&nbsp;&nbsp;��ѡ��������,����ѡ����Ч</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">��ʾ<%=sModuleName%>�б�����</td>
	<td class=tablerow1><input name="MaxListNum" type="text" id="MaxListNum" value="12" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">������������ַ�����</td>
	<td class=tablerow2><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow1 align="right">�Ƿ���ʾ�������ƣ�</td>
	<td class=tablerow1><select name="ShowClass" id="ShowClass">
          <option value='0'>����ʾ��������</option>
	  <option value='1'>��ʾ��������</option>
        </select></td>
</tr>
<%
If CInt(enchiasp.modules) = 1 or CInt(enchiasp.modules) = 6 Then
%>
<tr>
	<td class=tablerow2 align="right">�Ƿ���ʾͼ�ı��⣺</td>
	<td class=tablerow2><select name="ShowPic" id="ShowPic">
          <option value='1'>��ʾͼ�ı���</option>
	  <option value='0'>����ʾͼ�ı���</option>
        </select></td>
</tr>
<%
End If
%>
<tr>
	<td class=tablerow1 align="right">�Ƿ���ʾ���ڣ�</td>
	<td class=tablerow1><select name="ShowDate" id="ShowDate">
          <option value='1'>��ʾ����</option>
	  <option value='0'>����ʾ����</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">��ʾ���ڵĸ�ʽ��</td>
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
	<td class=tablerow1 align="right">�Ƿ��´��ڴ����ӣ�</td>
	<td class=tablerow1><select name="newindow" id="newindow">
          <option value='0' selected>�����ڴ�</option>
	  <option value='1'>�´��ڴ�</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">�б������ʽ���ƣ�</td>
	<td class=tablerow2><input name="StyleName" type="text" id="StyleName" value="showlist" size="15" maxlength="20"></td>
</tr>
<tr align=center>
	<td class=tablerow1 colspan="2"><input type=text name=SkinCode size=65 id="SkinCode"></td>
</tr>
<tr align=center>
	<td class=tablerow2 colspan="2">
		<input name="MakeJS" type="button" id="MakeJS" onclick="MakeCode();" value="&nbsp;����<%=sModuleName%>ģ���ǩ&nbsp;" class=button> 
		<input name="Copy" type="button" id="Copy" value="&nbsp;���Ƶ�������&nbsp;" onclick="copy();" class=button>
		<input type="button" value="&nbsp;�رմ���&nbsp;" onclick="selflabel();" class=button> 
		</td>
</tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function MakeCode()
{
if(document.myform.ClassID.value=="")
{
	alert("<%=sModuleName%>���಻��ָ���ⲿ��Ŀ��");
	document.myform.ClassID.focus();
	return false;
}
if(document.myform.MaxListNum.value=="")
{
	alert("��ʾ<%=sModuleName%>�б�������Ϊ�գ�");
	document.myform.MaxListNum.focus();
	return false;
}
if(document.myform.TitleMaxLen.value=="")
{
	alert("������������ַ�������Ϊ�գ�");
	document.myform.TitleMaxLen.focus();
	return false;
}
if(document.myform.StyleName.value=="")
{
	alert("�б���ʽ���Ʋ���Ϊ�գ�");
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
	<th colspan="2">����<%=sModuleName%>ģ��ͼƬ���ñ�ǩ</th>
</tr>
<form action="?action=image" method="post" name="myform" id="myform">

<tr>
	<td class=tablerow1 align="right">����Ƶ����</td>
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
	<td class=tablerow2 align="right">�������ࣺ</td>
	<td class=tablerow2><select name="ClassID" id="ClassID">
          <option value="0">����<%=sModuleName%>����</option>
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
        </select><font color="#0066CC">&nbsp;&nbsp;��ָ���˷���,��ѡ�����ͷ���</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">����ר�⣺</td>
	<td class=tablerow1><select name=SpecialID>
	<option value="0">��ѡ��ר��</option>
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
	<td class=tablerow2 align="right">�������ͣ�</td>
	<td class=tablerow2><select name="ShowType" id="ShowType">
          <option value="0">��������<%=sModuleName%>ͼƬ</option>
          <option value='1'>�����Ƽ�<%=sModuleName%>ͼƬ</option>
	  <option value='2'>��������<%=sModuleName%>ͼƬ</option>
	  <option value='3'>��������<%=sModuleName%>ͼƬ</option>
	  <option value='4'>�����Ƽ�<%=sModuleName%>ͼƬ</option>
	  <option value='5'>��������<%=sModuleName%>ͼƬ</option>
        </select><font color="#0066CC">&nbsp;&nbsp;��ѡ��������,����ѡ����Ч</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">�����ʾ������<%=sModuleName%>ͼƬ��</td>
	<td class=tablerow1><input name="MaxPicNum" type="text" id="MaxPicNum" value="12" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">ÿ����ʾ������<%=sModuleName%>ͼƬ��</td>
	<td class=tablerow2><input name="PerRowNum" type="text" id="PerRowNum" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow1 align="right">������������ַ�����</td>
	<td class=tablerow1><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">�Ƿ��´��ڴ����ӣ�</td>
	<td class=tablerow2><select name="newindow" id="newindow">
          <option value='0' selected>�����ڴ�</option>
	  <option value='1'>�´��ڴ�</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow1 align="right">ͼƬ�Ŀ�ȣ�</td>
	<td class=tablerow1><input name="picwidth" type="text" id="picwidth" value="120" size="5" maxlength="3"> ����</td>
</tr>
<tr>
	<td class=tablerow2 align="right">ͼƬ�ĸ߶ȣ�</td>
	<td class=tablerow2><input name="picheight" type="text" id="picheight" value="100" size="5" maxlength="3"> ����</td>
</tr>
<tr>
	<td class=tablerow1 align="right">�Ƿ���ʾ<%=sModuleName%>�������ƣ�</td>
	<td class=tablerow1><select name="showtopic" id="showtopic">
          <option value='1'>��ʾ</option>
	  <option value='0'>����ʾ</option>
	  <option value='2'>������ʾ</option>
        </select>����������ʾ����ʹ��2</td>
</tr>
<tr align=center>
	<td class=tablerow2 colspan="2"><input type=text name=SkinCode size=65 id="SkinCode"></td>
</tr>
<tr align=center>
	<td class=tablerow1 colspan="2">
		<input name="MakeJS" type="button" id="MakeJS" onclick="MakeCode();" value="&nbsp;����<%=sModuleName%>ģ���ǩ&nbsp;" class=button> 
		<input name="Copy" type="button" id="Copy" value="&nbsp;���Ƶ�������&nbsp;" onclick="copy();" class=button>
		<input type="reset" name="button" value="&nbsp;�رմ���&nbsp;" onclick="copy(); window.close()" class=button></td>
</tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function MakeCode()
{
if(document.myform.ClassID.value=="")
{
	alert("<%=sModuleName%>���಻��ָ���ⲿ��Ŀ��");
	document.myform.ClassID.focus();
	return false;
}
if(document.myform.MaxPicNum.value=="")
{
	alert("��ʾ<%=sModuleName%>ͼƬ������Ϊ�գ�");
	document.myform.MaxPicNum.focus();
	return false;
}
if(document.myform.TitleMaxLen.value=="")
{
	alert("������������ַ�������Ϊ�գ�");
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
	<th colspan="2">����<%=sModuleName%>ģ��ͼ�Ļ��ű�ǩ</th>
</tr>
<form action="?action=list" method="post" name="myform" id="myform">
<tr>
	<td class=tablerow1 align="right">����Ƶ����</td>
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
	<td class=tablerow2 align="right">�������ࣺ</td>
	<td class=tablerow2><select name="ClassID" id="ClassID">
          <option value="0">����<%=sModuleName%>����</option>
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
        </select><font color="#0066CC">&nbsp;&nbsp;��ָ���˷���,��ѡ�����ͷ���</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">����ר�⣺</td>
	<td class=tablerow1><select name=SpecialID>
	<option value="0">��ѡ��ר��</option>
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
	<td class=tablerow2 align="right">�������ͣ�</td>
	<td class=tablerow2><select name="ShowType" id="ShowType">
          <option value="0">��������<%=sModuleName%></option>
          <option value='1'>�����Ƽ�<%=sModuleName%></option>
	  <option value='2'>��������<%=sModuleName%></option>
<%
If CInt(enchiasp.modules) = 1 Then
%>
	  <option value='3'>����ͼ��<%=sModuleName%></option>
	  <option value='4'>��������<%=sModuleName%></option>
	  <option value='5'>�����Ƽ�<%=sModuleName%></option>
	  <option value='6'>��������<%=sModuleName%></option>
	  <option value='7'>����ͼ��<%=sModuleName%></option>
<%
Else
%>
	  <option value='3'>��������<%=sModuleName%></option>
	  <option value='4'>�����Ƽ�<%=sModuleName%></option>
	  <option value='5'>��������<%=sModuleName%></option>
<%
End If
%>
        </select><font color="#0066CC">&nbsp;&nbsp;��ѡ��������,����ѡ����Ч</font></td>
</tr>
<tr>
	<td class=tablerow1 align="right">��ʾ<%=sModuleName%>�б�����</td>
	<td class=tablerow1><input name="MaxListNum" type="text" id="MaxListNum" value="12" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow2 align="right">������������ַ�����</td>
	<td class=tablerow2><input name="TitleMaxLen" type="text" id="TitleMaxLen" value="26" size="5" maxlength="3"></td>
</tr>
<tr>
	<td class=tablerow1 align="right">�Ƿ���ʾ�������ƣ�</td>
	<td class=tablerow1><select name="ShowClass" id="ShowClass">
          <option value='0'>����ʾ��������</option>
	  <option value='1'>��ʾ��������</option>
        </select></td>
</tr>
<%
If CInt(enchiasp.modules) = 1 Then
%>
<tr>
	<td class=tablerow2 align="right">�Ƿ���ʾͼ�ı��⣺</td>
	<td class=tablerow2><select name="ShowPic" id="ShowPic">
          <option value='1'>��ʾͼ�ı���</option>
	  <option value='0'>����ʾͼ�ı���</option>
        </select></td>
</tr>
<%
End If
%>
<tr>
	<td class=tablerow1 align="right">�Ƿ���ʾ���ڣ�</td>
	<td class=tablerow1><select name="ShowDate" id="ShowDate">
          <option value='1'>��ʾ����</option>
	  <option value='0'>����ʾ����</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">��ʾ���ڵĸ�ʽ��</td>
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
	<td class=tablerow1 align="right">�Ƿ��´��ڴ����ӣ�</td>
	<td class=tablerow1><select name="newindow" id="newindow">
          <option value='0' selected>�����ڴ�</option>
	  <option value='1'>�´��ڴ�</option>
        </select></td>
</tr>
<tr>
	<td class=tablerow2 align="right">�б������ʽ���ƣ�</td>
	<td class=tablerow2><input name="StyleName" type="text" id="StyleName" value="showlist" size="15" maxlength="20"></td>
</tr>
<tr align=center>
	<td class=tablerow1 colspan="2"><input type=text name=SkinCode size=65 id="SkinCode"></td>
</tr>
<tr align=center>
	<td class=tablerow2 colspan="2">
		<input name="MakeJS" type="button" id="MakeJS" onclick="MakeCode();" value="&nbsp;����<%=sModuleName%>ģ���ǩ&nbsp;" class=button> 
		<input name="Copy" type="button" id="Copy" value="&nbsp;���Ƶ�������&nbsp;" onclick="copy();" class=button>
		<input type="button" value="&nbsp;�رմ���&nbsp;" onclick="selflabel();" class=button> 
		</td>
</tr>
</form>
</table>
<script language="JavaScript" type="text/JavaScript">
function MakeCode()
{
if(document.myform.ClassID.value=="")
{
	alert("<%=sModuleName%>���಻��ָ���ⲿ��Ŀ��");
	document.myform.ClassID.focus();
	return false;
}
if(document.myform.MaxListNum.value=="")
{
	alert("��ʾ<%=sModuleName%>�б�������Ϊ�գ�");
	document.myform.MaxListNum.focus();
	return false;
}
if(document.myform.TitleMaxLen.value=="")
{
	alert("������������ַ�������Ϊ�գ�");
	document.myform.TitleMaxLen.focus();
	return false;
}
if(document.myform.StyleName.value=="")
{
	alert("�б���ʽ���Ʋ���Ϊ�գ�");
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