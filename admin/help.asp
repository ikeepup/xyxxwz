<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<head>
<title>CMS安装帮助指南</title>
</head>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th><font face="宋体">山西通达集团CMS使用帮助</font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">一、什么是CMS？</font></p>
      <p align="left" style="line-height: 200%; margin: 0"><font face="宋体">&nbsp;&nbsp;&nbsp;   
      CMS为内容管理系统，可以实现网站DIY，包含以下几个模块：新闻、下载、商城、动画、采集、单页面图文、论坛整合。采用模板和标签及缓存和JS。拥有强大的后台和安全机制。支持首页动画定制。</font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">二、产品安装指南</p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">1、将CMS内容拷贝到一个目录下</p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">2、修改两处数据库连接及数据库相关内容（注意数据库类型，数据库类型分ACCESS和SQL两种）</p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（1）主站数据库连接密码 conn.asp&nbsp;</font></p>            
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="宋体">第31行</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="宋体">
    SqlUsername = &quot;sa&quot;          '用户名&nbsp;</font>         
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="宋体">
	SqlPassword = &quot;sql&quot;          '用户密码&nbsp;</font></p>         
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="宋体">
	SqlLocalName = "(local)"        '连接名（本地用local，外地用IP）</font></p>          
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="宋体">修改为</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体"><font color="#800000">正确的用户名和密码，注意不要用SA帐号，请从SQL中分配一个用户</font></font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体" color="#000000">(2)BBS数据库连接密码   
      bbs/conn.asp</font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体" color="#000000">第23行</font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">Const SqlPassword =   
      &quot;sql&quot;  '数据库密码&nbsp;</font></p>          
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">
Const SqlUsername = &quot;sa&quot;   '数据库用户名&nbsp;</font></p>          
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">
Const SqlLocalName = "(local)" '数据库地址</font></p>           
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">为保证系统安全，建议与CMS数据库为不同的用户和密码，不要使用SA帐号</font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">(3)API整合接口，此系统为BBS接口专用</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">在后台登陆后，选择其他管理--论坛管理，将其中<u>整合程序的接口文件路径：http://localhost/test2/bbs/dv_dpo.asp修改为</u></font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><a href="http://www.ltzxw.com/bbs/dv_dpo.asp"><font face="宋体">http://www.sxtongda.com.cn/bbs/dv_dpo.asp</font></a></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（4）BBS整合接口参数修改</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">BBS/Api_Config.asp 修改相关参数 </p>        
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">Const DvApi_Urls	=   
      &quot;http://localhost/test2/API/API_Response.asp&quot;        
      <font color="#800000">修改</font>为Const DvApi_Urls	= &quot;http://www.sxtongda.com.cn/API/API_Response.asp&quot; </p>      
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">3、数据库恢复<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">（1）、新建数据库EC_CMS_TONGD并恢复<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">（2）、新建数据库dv_bbs并恢复<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">4、设置IIS，注意开启写入功能。如IIS为6.0则去除IIS6.0最大附件上传下载限制<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;<font color="#FF0000"> 
      在 IIS 6.0 中，默认设置是特别严格和安全的，最大只能传送 204,800 个字节，这样可以最大限度地减少因以前太宽松的超时和限制而造成的攻击。（在 IIS 6.0 之前的版本中无此限制）<br> 
      解决办法：</font><br>
      A、先在服务里关闭 iis admin service 服务。<br>   
      B、找到 windows\system32\inetsrv\ 下的 metabase.xml 文件。<br>   
      C、用纯文本方式打开，找到 ASPMaxRequestEntityAllowed 把它修改为需要的值（可修改为10M即：10240000），默认为：204800，即：200K。<br>   
      C、存盘，然后重启 iis admin service 服务。<br>   
      &nbsp;<font color="#FF0000">&nbsp; 在 IIS 6.0 中，无法下载超过4M的附件时，可以按以下步骤解决：</font><br> 
      A、先在服务里关闭 iis admin service 服务。<br>   
      B、找到 windows\system32\inetsrv\ 下的 metabase.xml 文件。<br>   
      C、用纯文本方式打开，找到 AspBufferingLimit 把它修改为需要的值（可修改为20M即：20480000）。<br>   
      D、存盘，然后重启 iis admin service 服务。<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">5、安装相关组件<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">（1）、JMAIL组件，请合理使用JMAIL组件，并设置邮件和用户注册选项<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">（2）、ASPJPEG组件，请注意何时开启此组件，在顶部广告文件上传时应关闭此组件。<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">（3）、脚本解释引擎：VBScript/5.6.7426以上  如版本过低请更新新版本<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">（4）、无组件上传开启。并注意水印设置，如果没有水印功能请关闭水印否则提示不支持。重新开启ADODB.Stream组件的办法：在开始—运行里执行：regsvr32 "C:\Program Files\Common Files\System\ado\msado15.dll<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">6、<font face="宋体">如果想将记数器清零，请选择提供的数据库进行覆盖，覆盖位置/count/mbdata/mbcount.asp</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">7、进入后台更新基本设置，重建缓存、注册，请选择一个类型进行注册</font>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册网址：http://www.sxtongda.com.cn/</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册日期：2007-7-16</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册人：liuyunfan</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册模块：所有模块</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册码：7afd3a9260d27973-29814027117a20805aaf3386d9a84126-ee943564637f0609545c83c526ffc6a84e5fda348606ab88-c8ce9bfb624d3ba9</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注意不要有空格，内容按照上述进行注册,注册后更新缓存</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">　
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册网址：http://127.0.0.1/</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册日期：2007-7-16</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册人：liuyunfan</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册模块：所有模块</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册码：943a2b843477baff-29814027117a20805aaf3386d9a84126-ee943564637f0609545c83c526ffc6a84e5fda348606ab88-c8ce9bfb624d3ba9</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注意不要有空格，内容按照上述进行注册,注册后更新缓存</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">　
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册网址：http://192.168.0.219/</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册日期：2007-7-16</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册人：liuyunfan</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册模块：所有模块</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注册码：cf3c5eae7398c7a4-29814027117a20805aaf3386d9a84126-ee943564637f0609545c83c526ffc6a84e5fda348606ab88-c8ce9bfb624d3ba9</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">注意不要有空格，内容按照上述进行注册,注册后更新缓存</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">　<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">8、其它注意事项</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（1）、论坛可以设置先关闭，登陆论坛后设置，用户：admin   
      密码：admin888 需输前台和后台</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（2）、会员注册可以先关闭</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（3）、友情连接先关闭</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">　<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">9、注意主机的SQL连接数、WINDOWS补丁、IIS安全设置，适当的时候将数据库服务器屏蔽1433端口</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">10、可制定数据库自动备份功能。</font></th>  
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">三、系统登陆</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">1、为保证系统安全登陆文件已经改名，防止恶意猜测登陆地址。</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">2、登陆地址：admin/admin_klogin.asp(如开启二次码则系统会自动转向admin_loginx.asp)</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">3、如果系统开启了二次码，那么必须转向成功后才能登陆，否则系统不让登陆</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">4、主站<font color="#800000">默认用户名</font>：admin&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000"> 
      默认密码 
      </font> admin&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#800000">默认二次密码开启键</font> 空格</font></p>           
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体"><font color="#800000">默认二次密码基码</font> liuyunfan&nbsp;          
      <font color="#800000">默认密码规则</font> 第1个位置和第1个位置加法运算，插如基码第一个位置之后</font></p>            
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">5、有关二次码是否开启，在后台登陆后基本设置中[管理员安全]修改。每个管理员可以设置不同的二次码规则，二次码基码和开启键相同</font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">四、关于管理员设置</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">建议只开启一个高级管理员，另一个为普通管理员。普通管理员请设置普通管理员的二次码规则和权限。可设置管理员的IP绑定。请设置好IP地址，并设置其他管理员的权限,并可以制定该管理员的<font color="#FF0000">IP规则</font></font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">五、需要注意的几个问题</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（1）由于系统采用模板级标签控制，因此，当不熟悉标签和HTML相关知识，请勿修改</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">&nbsp;(2)请定期清除无效数据、上传文件等，清理文件时请谨慎使用</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（3）请定期备份系统，数据库文件可通过计划进行定时备份，以避免不必要的损失</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="宋体">（4）论坛采用动网论坛，由于采用一站通行，因此，论坛关闭了会员注册，会员在主站注册后会自动在论坛注册，会员主站登陆后，论坛可以登陆</font></th>
</tr>
<tr>
	<th>
      <p align="left">六、留言版可设置是否允许匿名留言，具体修改请更改模版设置。</p>
    </th>
</tr>
<tr>
	<th>
      <p align="left">七、上传图片时无法出现确定按纽，请修改设置WINDOWS的显示模式，设置为标准
    </th>
</tr>
<tr>
	<th>
      <p align="left">八、新闻中可以上传其他文件，在工具菜单旁边
    </th>
</tr>
<tr>
	<th>
      <p align="left">九、设置首页图片新闻时，注意将HTML字符去除，如果不希望在首页显示该图片请去掉首页图片
    </th>
</tr>
<tr>
	<th>
      <p align="left"></p>
    </th>
</tr>
<tr>
	<th><font face="宋体">常用JS脚本</font></th>
</tr>
<tr>
	<td class=tablerow1><font face="宋体">1、禁止拷贝和粘帖，同时禁止右键</font>
      <p><font face="宋体">&lt;body oncontextmenu=&quot;return false&quot; ondragstart=&quot;return            
      false&quot; onselectstart =&quot;return false&quot; onselect=&quot;document.selection.empty()&quot;            
      oncopy=&quot;document.selection.empty()&quot; onbeforecopy=&quot;return&nbsp;<br>           
      false&quot;onmouseup=&quot;document.selection.empty()&quot; leftmargin=&quot;0&quot;            
      topmargin=&quot;0&quot; bgcolor=&quot;#ffffff&quot;&gt;<br>            
      &lt;script&gt;&nbsp;<br>
      function stop(){&nbsp;<br>           
      alert(&quot;欢迎您登陆山西通达集团&quot;);&nbsp;<br>
      return false;&nbsp;<br>           
      }&nbsp;<br>
      document.oncontextmenu=stop;&nbsp;<br>
      &lt;/script&gt;&nbsp;<br>
      <font color="#800000">注意：</font>此段代码加在通栏模版中，加此段代码后会出现搜索时无法显示，可以只屏蔽右键加如下代码</font>
      <p><font face="宋体">&lt;SCRIPT language=javascript><br>          
      &lt;!--<br>
      function Click()&nbsp;<br>         
      {<br>
	if(event.button!=1)<br>
	{<br>
	alert("欢迎您光临『山西通达集团』！");<br>
	}<br>
      }<br>
      document.onmousedown=Click;<br>
      //--><br>
      &lt;/SCRIPT></font>
      <p><font face="宋体">禁止拷贝</font>
      <p><font face="宋体">&lt;body oncopy=&quot;document.selection.empty()&quot;&gt;</font>       
      <p><font face="宋体">　</font></td>                
</tr>
<tr>
	<td class=tablerow1><font face="宋体">2、可以加跑马灯</font>
      <p><font COLOR="#0000c0" face="宋体">&lt;SCRIPT LANGUAGE=JAVASCRIPT&gt;</font></p>        
      <p><font COLOR="#800000" face="宋体">&lt;!--</font></p>
      <p><font COLOR="#800000" face="宋体">var msg = &quot;欢迎光临</font><font face="宋体">山西通达集团</font><font COLOR="#800000" face="宋体">&quot;;</font></p>       
      <p><font COLOR="#800000" face="宋体">var speed = 300;</font></p> 
      <p><font COLOR="#800000" face="宋体">var msgud = &quot; &quot; + msg;</font></p>
      <p><font COLOR="#800000" face="宋体">function statusScroll() {</font></p> 
      <p><font COLOR="#800000" face="宋体">if (msgud.length &lt;msg.length) msgud  
      += &quot; - &quot; + msg;</font></p>
      <p><font COLOR="#800000" face="宋体">msgud = msgud.substring(1, 
      msgud.length);</font></p>
      <p><font COLOR="#800000" face="宋体">window.status = msgud.substring(0, 
      msg.length);</font></p>
      <p><font COLOR="#800000" face="宋体">window.setTimeout(&quot;statusScroll()&quot;,  
      speed);</font></p>
      <p><font COLOR="#800000" face="宋体">}</font></p>
      <p><font COLOR="#800000" face="宋体">--&gt;</font></p>
      <p><font COLOR="#0000c0" face="宋体">&lt;/SCRIPT&gt;</font></p>
      <p><font face="宋体">位置放在通栏模板&lt;head&gt;&lt;/head&gt;中间</font></p>
      <p><font face="宋体">在&lt;body&gt; 中加入<font COLOR="#0000c0">onload=</font>&quot;window.setTimeout('statusScroll()',        
      500)&quot;</font></td>                 
</tr>
<tr>
	<td class=tablerow1><font color="#800000" face="宋体">3、显示星期几</font>
      <p><font face="宋体" COLOR="#0000c0">&lt;script language=JavaScript&gt;</font></p>       
      <p><font face="宋体" COLOR="#800000">today=new Date();</font></p> 
      <p><font face="宋体" COLOR="#800000">function initArray(){</font></p> 
      <p><font face="宋体" COLOR="#800000">this.length=initArray.arguments.length</font></p>
      <p><font face="宋体" COLOR="#800000">for(var i=0;i&lt;this.length;i++)</font></p> 
      <p><font face="宋体" COLOR="#800000">this[i+1]=initArray.arguments[i] }</font></p> 
      <p><font face="宋体" COLOR="#800000">var d=new initArray(</font></p> 
      <p><font face="宋体" COLOR="#800000">&quot;星期日&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;星期一&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;星期二&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;星期三&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;星期四&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;星期五&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;星期六&quot;);</font></p>
      <p><font face="宋体" COLOR="#800000">document.write(</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;&lt;font color=##000000 style='font-size:9pt;font-family: 宋体'&gt;        
      &quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">today.getYear(),&quot;年&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">today.getMonth()+1,&quot;月&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">today.getDate(),&quot;日&quot;,</font></p>
      <p><font face="宋体" COLOR="#800000">d[today.getDay()+1],</font></p>
      <p><font face="宋体" COLOR="#800000">&quot;&lt;/font&gt;&quot; );</font></p> 
      <p><font face="宋体" COLOR="#0000c0">&lt;/script&gt;</font></p>
      <p><font face="宋体">　</font></td>                 
</tr>
<tr>
	<td class=tablerow1><font color="#800000" face="宋体">4、自动滚屏</font>
      <p><font face="宋体" COLOR="#0000c0">&lt;script language</font><font face="宋体">=&quot;JavaScript&quot;<font SIZE="1" COLOR="#0000c0">&gt;</font></font></p>      
      <p><font face="宋体" COLOR="#800000">&lt;!--</font></p>
      <p><font face="宋体" COLOR="#800000">function click() {</font></p> 
      <p><font face="宋体" COLOR="#800000">if (event.button==2) {</font></p> 
      <p><font face="宋体" COLOR="#800000">if(document.all.auto.status==true){document.all.auto.status=false;alert(&quot;自动滚屏已经停止了！&quot;)}</font></p>
      <p><font face="宋体" COLOR="#800000">scroller();</font></p>
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">document.onmousedown=click</font></p>
      <p><font face="宋体" COLOR="#800000">var position = 0;</font></p> 
      <p><font face="宋体" COLOR="#800000">function scroller() {</font></p> 
      <p><font face="宋体" COLOR="#800000">if (document.all.auto.status==true){</font></p> 
      <p><font face="宋体" COLOR="#800000">position++;</font></p>
      <p><font face="宋体" COLOR="#800000">scroll(0,position);</font></p>
      <p><font face="宋体" COLOR="#800000">clearTimeout(timer);</font></p>
      <p><font face="宋体" COLOR="#800000">var timer = setTimeout(&quot;scroller()&quot;,50);</font></p> 
      <p><font face="宋体" COLOR="#800000">timer;</font></p>
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">else{</font></p>
      <p><font face="宋体" COLOR="#800000">clearTimeout(timer);</font></p>
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">function MM_callJS(jsStr) { //v2.0</font></p> 
      <p><font face="宋体" COLOR="#800000">return eval(jsStr)</font></p> 
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">//--&gt;</font></p>
      <p><font face="宋体" COLOR="#0000c0">&lt;/script&gt;</font></p>
      <p><font face="宋体" COLOR="#0000c0">&lt;SCRIPT language=javascript&gt;</font></p> 
      <p><font face="宋体" COLOR="#800000">&lt;!--</font></p>
      <p><font face="宋体" COLOR="#800000">function mOvr(src,clrOver) {</font></p> 
      <p><font face="宋体" COLOR="#800000">if (!src.contains(event.fromElement)) {</font></p> 
      <p><font face="宋体" COLOR="#800000">src.bgColor = clrOver;</font></p> 
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">function mOut(src,clrIn) {</font></p> 
      <p><font face="宋体" COLOR="#800000">if (!src.contains(event.toElement)) {</font></p> 
      <p><font face="宋体" COLOR="#800000">src.bgColor = clrIn;</font></p> 
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">}</font></p>
      <p><font face="宋体" COLOR="#800000">// --&gt;</font></p> 
      <p><font face="宋体" COLOR="#0000c0">&lt;/SCRIPT&gt;</font></p>
      <p><font face="宋体">　</font></td>                 
</tr>
</table>
