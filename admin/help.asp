<!--#include file="setup.asp"-->
<!--#include file="check.asp"-->
<head>
<title>CMS��װ����ָ��</title>
</head>

<table border=0 align=center cellpadding=3 cellspacing=1 class=tableborder>
<tr>
	<th><font face="����">ɽ��ͨ�Ｏ��CMSʹ�ð���</font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">һ��ʲô��CMS��</font></p>
      <p align="left" style="line-height: 200%; margin: 0"><font face="����">&nbsp;&nbsp;&nbsp;   
      CMSΪ���ݹ���ϵͳ������ʵ����վDIY���������¼���ģ�飺���š����ء��̳ǡ��������ɼ�����ҳ��ͼ�ġ���̳���ϡ�����ģ��ͱ�ǩ�������JS��ӵ��ǿ��ĺ�̨�Ͱ�ȫ���ơ�֧����ҳ�������ơ�</font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">������Ʒ��װָ��</p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">1����CMS���ݿ�����һ��Ŀ¼��</p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">2���޸��������ݿ����Ӽ����ݿ�������ݣ�ע�����ݿ����ͣ����ݿ����ͷ�ACCESS��SQL���֣�</p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��1����վ���ݿ��������� conn.asp&nbsp;</font></p>            
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="����">��31��</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="����">
    SqlUsername = &quot;sa&quot;          '�û���&nbsp;</font>         
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="����">
	SqlPassword = &quot;sql&quot;          '�û�����&nbsp;</font></p>         
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="����">
	SqlLocalName = "(local)"        '��������������local�������IP��</font></p>          
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000" face="����">�޸�Ϊ</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����"><font color="#800000">��ȷ���û��������룬ע�ⲻҪ��SA�ʺţ����SQL�з���һ���û�</font></font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����" color="#000000">(2)BBS���ݿ���������   
      bbs/conn.asp</font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����" color="#000000">��23��</font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">Const SqlPassword =   
      &quot;sql&quot;  '���ݿ�����&nbsp;</font></p>          
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">
Const SqlUsername = &quot;sa&quot;   '���ݿ��û���&nbsp;</font></p>          
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">
Const SqlLocalName = "(local)" '���ݿ��ַ</font></p>           
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font color="#800000">Ϊ��֤ϵͳ��ȫ��������CMS���ݿ�Ϊ��ͬ���û������룬��Ҫʹ��SA�ʺ�</font></p>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">(3)API���Ͻӿڣ���ϵͳΪBBS�ӿ�ר��</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">�ں�̨��½��ѡ����������--��̳����������<u>���ϳ���Ľӿ��ļ�·����http://localhost/test2/bbs/dv_dpo.asp�޸�Ϊ</u></font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><a href="http://www.ltzxw.com/bbs/dv_dpo.asp"><font face="����">http://www.sxtongda.com.cn/bbs/dv_dpo.asp</font></a></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��4��BBS���Ͻӿڲ����޸�</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">BBS/Api_Config.asp �޸���ز��� </p>        
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">Const DvApi_Urls	=   
      &quot;http://localhost/test2/API/API_Response.asp&quot;        
      <font color="#800000">�޸�</font>ΪConst DvApi_Urls	= &quot;http://www.sxtongda.com.cn/API/API_Response.asp&quot; </p>      
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">3�����ݿ�ָ�<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��1�����½����ݿ�EC_CMS_TONGD���ָ�<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��2�����½����ݿ�dv_bbs���ָ�<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">4������IIS��ע�⿪��д�빦�ܡ���IISΪ6.0��ȥ��IIS6.0��󸽼��ϴ���������<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">&nbsp;&nbsp;<font color="#FF0000"> 
      �� IIS 6.0 �У�Ĭ���������ر��ϸ�Ͱ�ȫ�ģ����ֻ�ܴ��� 204,800 ���ֽڣ�������������޶ȵؼ�������ǰ̫���ɵĳ�ʱ�����ƶ���ɵĹ��������� IIS 6.0 ֮ǰ�İ汾���޴����ƣ�<br> 
      ����취��</font><br>
      A�����ڷ�����ر� iis admin service ����<br>   
      B���ҵ� windows\system32\inetsrv\ �µ� metabase.xml �ļ���<br>   
      C���ô��ı���ʽ�򿪣��ҵ� ASPMaxRequestEntityAllowed �����޸�Ϊ��Ҫ��ֵ�����޸�Ϊ10M����10240000����Ĭ��Ϊ��204800������200K��<br>   
      C�����̣�Ȼ������ iis admin service ����<br>   
      &nbsp;<font color="#FF0000">&nbsp; �� IIS 6.0 �У��޷����س���4M�ĸ���ʱ�����԰����²�������</font><br> 
      A�����ڷ�����ر� iis admin service ����<br>   
      B���ҵ� windows\system32\inetsrv\ �µ� metabase.xml �ļ���<br>   
      C���ô��ı���ʽ�򿪣��ҵ� AspBufferingLimit �����޸�Ϊ��Ҫ��ֵ�����޸�Ϊ20M����20480000����<br>   
      D�����̣�Ȼ������ iis admin service ����<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">5����װ������<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��1����JMAIL����������ʹ��JMAIL������������ʼ����û�ע��ѡ��<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��2����ASPJPEG�������ע���ʱ������������ڶ�������ļ��ϴ�ʱӦ�رմ������<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��3�����ű��������棺VBScript/5.6.7426����  ��汾����������°汾<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��4����������ϴ���������ע��ˮӡ���ã����û��ˮӡ������ر�ˮӡ������ʾ��֧�֡����¿���ADODB.Stream����İ취���ڿ�ʼ��������ִ�У�regsvr32 "C:\Program Files\Common Files\System\ado\msado15.dll<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">6��<font face="����">����뽫���������㣬��ѡ���ṩ�����ݿ���и��ǣ�����λ��/count/mbdata/mbcount.asp</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">7�������̨���»������ã��ؽ����桢ע�ᣬ��ѡ��һ�����ͽ���ע��</font>  
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע����ַ��http://www.sxtongda.com.cn/</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע�����ڣ�2007-7-16</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע���ˣ�liuyunfan</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע��ģ�飺����ģ��</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע���룺7afd3a9260d27973-29814027117a20805aaf3386d9a84126-ee943564637f0609545c83c526ffc6a84e5fda348606ab88-c8ce9bfb624d3ba9</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע�ⲻҪ�пո����ݰ�����������ע��,ע�����»���</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע����ַ��http://127.0.0.1/</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע�����ڣ�2007-7-16</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע���ˣ�liuyunfan</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע��ģ�飺����ģ��</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע���룺943a2b843477baff-29814027117a20805aaf3386d9a84126-ee943564637f0609545c83c526ffc6a84e5fda348606ab88-c8ce9bfb624d3ba9</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע�ⲻҪ�пո����ݰ�����������ע��,ע�����»���</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע����ַ��http://192.168.0.219/</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע�����ڣ�2007-7-16</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע���ˣ�liuyunfan</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע��ģ�飺����ģ��</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע���룺cf3c5eae7398c7a4-29814027117a20805aaf3386d9a84126-ee943564637f0609545c83c526ffc6a84e5fda348606ab88-c8ce9bfb624d3ba9</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">ע�ⲻҪ�пո����ݰ�����������ע��,ע�����»���</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">8������ע������</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��1������̳���������ȹرգ���½��̳�����ã��û���admin   
      ���룺admin888 ����ǰ̨�ͺ�̨</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��2������Աע������ȹر�</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��3�������������ȹر�</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0">��<p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">9��ע��������SQL��������WINDOWS������IIS��ȫ���ã��ʵ���ʱ�����ݿ����������1433�˿�</font><p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">10�����ƶ����ݿ��Զ����ݹ��ܡ�</font></th>  
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">����ϵͳ��½</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">1��Ϊ��֤ϵͳ��ȫ��½�ļ��Ѿ���������ֹ����²��½��ַ��</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">2����½��ַ��admin/admin_klogin.asp(�翪����������ϵͳ���Զ�ת��admin_loginx.asp)</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">3�����ϵͳ�����˶����룬��ô����ת��ɹ�����ܵ�½������ϵͳ���õ�½</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">4����վ<font color="#800000">Ĭ���û���</font>��admin&nbsp;&nbsp;&nbsp;&nbsp;<font color="#800000"> 
      Ĭ������ 
      </font> admin&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#800000">Ĭ�϶������뿪����</font> �ո�</font></p>           
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����"><font color="#800000">Ĭ�϶����������</font> liuyunfan&nbsp;          
      <font color="#800000">Ĭ���������</font> ��1��λ�ú͵�1��λ�üӷ����㣬��������һ��λ��֮��</font></p>            
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">5���йض������Ƿ������ں�̨��½�����������[����Ա��ȫ]�޸ġ�ÿ������Ա�������ò�ͬ�Ķ�������򣬶��������Ϳ�������ͬ</font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">�ġ����ڹ���Ա����</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">����ֻ����һ���߼�����Ա����һ��Ϊ��ͨ����Ա����ͨ����Ա��������ͨ����Ա�Ķ���������Ȩ�ޡ������ù���Ա��IP�󶨡������ú�IP��ַ����������������Ա��Ȩ��,�������ƶ��ù���Ա��<font color="#FF0000">IP����</font></font></th>
</tr>
<tr>
	<th>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">�塢��Ҫע��ļ�������</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��1������ϵͳ����ģ�弶��ǩ���ƣ���ˣ�������Ϥ��ǩ��HTML���֪ʶ�������޸�</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">&nbsp;(2)�붨�������Ч���ݡ��ϴ��ļ��ȣ������ļ�ʱ�����ʹ��</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��3���붨�ڱ���ϵͳ�����ݿ��ļ���ͨ���ƻ����ж�ʱ���ݣ��Ա��ⲻ��Ҫ����ʧ</font></p>
      <p align="left" style="line-height: 200%; margin-top: 0; margin-bottom: 0"><font face="����">��4����̳���ö�����̳�����ڲ���һվͨ�У���ˣ���̳�ر��˻�Աע�ᣬ��Ա����վע�����Զ�����̳ע�ᣬ��Ա��վ��½����̳���Ե�½</font></th>
</tr>
<tr>
	<th>
      <p align="left">�������԰�������Ƿ������������ԣ������޸������ģ�����á�</p>
    </th>
</tr>
<tr>
	<th>
      <p align="left">�ߡ��ϴ�ͼƬʱ�޷�����ȷ����Ŧ�����޸�����WINDOWS����ʾģʽ������Ϊ��׼
    </th>
</tr>
<tr>
	<th>
      <p align="left">�ˡ������п����ϴ������ļ����ڹ��߲˵��Ա�
    </th>
</tr>
<tr>
	<th>
      <p align="left">�š�������ҳͼƬ����ʱ��ע�⽫HTML�ַ�ȥ���������ϣ������ҳ��ʾ��ͼƬ��ȥ����ҳͼƬ
    </th>
</tr>
<tr>
	<th>
      <p align="left"></p>
    </th>
</tr>
<tr>
	<th><font face="����">����JS�ű�</font></th>
</tr>
<tr>
	<td class=tablerow1><font face="����">1����ֹ������ճ����ͬʱ��ֹ�Ҽ�</font>
      <p><font face="����">&lt;body oncontextmenu=&quot;return false&quot; ondragstart=&quot;return            
      false&quot; onselectstart =&quot;return false&quot; onselect=&quot;document.selection.empty()&quot;            
      oncopy=&quot;document.selection.empty()&quot; onbeforecopy=&quot;return&nbsp;<br>           
      false&quot;onmouseup=&quot;document.selection.empty()&quot; leftmargin=&quot;0&quot;            
      topmargin=&quot;0&quot; bgcolor=&quot;#ffffff&quot;&gt;<br>            
      &lt;script&gt;&nbsp;<br>
      function stop(){&nbsp;<br>           
      alert(&quot;��ӭ����½ɽ��ͨ�Ｏ��&quot;);&nbsp;<br>
      return false;&nbsp;<br>           
      }&nbsp;<br>
      document.oncontextmenu=stop;&nbsp;<br>
      &lt;/script&gt;&nbsp;<br>
      <font color="#800000">ע�⣺</font>�˶δ������ͨ��ģ���У��Ӵ˶δ������������ʱ�޷���ʾ������ֻ�����Ҽ������´���</font>
      <p><font face="����">&lt;SCRIPT language=javascript><br>          
      &lt;!--<br>
      function Click()&nbsp;<br>         
      {<br>
	if(event.button!=1)<br>
	{<br>
	alert("��ӭ�����١�ɽ��ͨ�Ｏ�š���");<br>
	}<br>
      }<br>
      document.onmousedown=Click;<br>
      //--><br>
      &lt;/SCRIPT></font>
      <p><font face="����">��ֹ����</font>
      <p><font face="����">&lt;body oncopy=&quot;document.selection.empty()&quot;&gt;</font>       
      <p><font face="����">��</font></td>                
</tr>
<tr>
	<td class=tablerow1><font face="����">2�����Լ������</font>
      <p><font COLOR="#0000c0" face="����">&lt;SCRIPT LANGUAGE=JAVASCRIPT&gt;</font></p>        
      <p><font COLOR="#800000" face="����">&lt;!--</font></p>
      <p><font COLOR="#800000" face="����">var msg = &quot;��ӭ����</font><font face="����">ɽ��ͨ�Ｏ��</font><font COLOR="#800000" face="����">&quot;;</font></p>       
      <p><font COLOR="#800000" face="����">var speed = 300;</font></p> 
      <p><font COLOR="#800000" face="����">var msgud = &quot; &quot; + msg;</font></p>
      <p><font COLOR="#800000" face="����">function statusScroll() {</font></p> 
      <p><font COLOR="#800000" face="����">if (msgud.length &lt;msg.length) msgud  
      += &quot; - &quot; + msg;</font></p>
      <p><font COLOR="#800000" face="����">msgud = msgud.substring(1, 
      msgud.length);</font></p>
      <p><font COLOR="#800000" face="����">window.status = msgud.substring(0, 
      msg.length);</font></p>
      <p><font COLOR="#800000" face="����">window.setTimeout(&quot;statusScroll()&quot;,  
      speed);</font></p>
      <p><font COLOR="#800000" face="����">}</font></p>
      <p><font COLOR="#800000" face="����">--&gt;</font></p>
      <p><font COLOR="#0000c0" face="����">&lt;/SCRIPT&gt;</font></p>
      <p><font face="����">λ�÷���ͨ��ģ��&lt;head&gt;&lt;/head&gt;�м�</font></p>
      <p><font face="����">��&lt;body&gt; �м���<font COLOR="#0000c0">onload=</font>&quot;window.setTimeout('statusScroll()',        
      500)&quot;</font></td>                 
</tr>
<tr>
	<td class=tablerow1><font color="#800000" face="����">3����ʾ���ڼ�</font>
      <p><font face="����" COLOR="#0000c0">&lt;script language=JavaScript&gt;</font></p>       
      <p><font face="����" COLOR="#800000">today=new Date();</font></p> 
      <p><font face="����" COLOR="#800000">function initArray(){</font></p> 
      <p><font face="����" COLOR="#800000">this.length=initArray.arguments.length</font></p>
      <p><font face="����" COLOR="#800000">for(var i=0;i&lt;this.length;i++)</font></p> 
      <p><font face="����" COLOR="#800000">this[i+1]=initArray.arguments[i] }</font></p> 
      <p><font face="����" COLOR="#800000">var d=new initArray(</font></p> 
      <p><font face="����" COLOR="#800000">&quot;������&quot;,</font></p>
      <p><font face="����" COLOR="#800000">&quot;����һ&quot;,</font></p>
      <p><font face="����" COLOR="#800000">&quot;���ڶ�&quot;,</font></p>
      <p><font face="����" COLOR="#800000">&quot;������&quot;,</font></p>
      <p><font face="����" COLOR="#800000">&quot;������&quot;,</font></p>
      <p><font face="����" COLOR="#800000">&quot;������&quot;,</font></p>
      <p><font face="����" COLOR="#800000">&quot;������&quot;);</font></p>
      <p><font face="����" COLOR="#800000">document.write(</font></p>
      <p><font face="����" COLOR="#800000">&quot;&lt;font color=##000000 style='font-size:9pt;font-family: ����'&gt;        
      &quot;,</font></p>
      <p><font face="����" COLOR="#800000">today.getYear(),&quot;��&quot;,</font></p>
      <p><font face="����" COLOR="#800000">today.getMonth()+1,&quot;��&quot;,</font></p>
      <p><font face="����" COLOR="#800000">today.getDate(),&quot;��&quot;,</font></p>
      <p><font face="����" COLOR="#800000">d[today.getDay()+1],</font></p>
      <p><font face="����" COLOR="#800000">&quot;&lt;/font&gt;&quot; );</font></p> 
      <p><font face="����" COLOR="#0000c0">&lt;/script&gt;</font></p>
      <p><font face="����">��</font></td>                 
</tr>
<tr>
	<td class=tablerow1><font color="#800000" face="����">4���Զ�����</font>
      <p><font face="����" COLOR="#0000c0">&lt;script language</font><font face="����">=&quot;JavaScript&quot;<font SIZE="1" COLOR="#0000c0">&gt;</font></font></p>      
      <p><font face="����" COLOR="#800000">&lt;!--</font></p>
      <p><font face="����" COLOR="#800000">function click() {</font></p> 
      <p><font face="����" COLOR="#800000">if (event.button==2) {</font></p> 
      <p><font face="����" COLOR="#800000">if(document.all.auto.status==true){document.all.auto.status=false;alert(&quot;�Զ������Ѿ�ֹͣ�ˣ�&quot;)}</font></p>
      <p><font face="����" COLOR="#800000">scroller();</font></p>
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">document.onmousedown=click</font></p>
      <p><font face="����" COLOR="#800000">var position = 0;</font></p> 
      <p><font face="����" COLOR="#800000">function scroller() {</font></p> 
      <p><font face="����" COLOR="#800000">if (document.all.auto.status==true){</font></p> 
      <p><font face="����" COLOR="#800000">position++;</font></p>
      <p><font face="����" COLOR="#800000">scroll(0,position);</font></p>
      <p><font face="����" COLOR="#800000">clearTimeout(timer);</font></p>
      <p><font face="����" COLOR="#800000">var timer = setTimeout(&quot;scroller()&quot;,50);</font></p> 
      <p><font face="����" COLOR="#800000">timer;</font></p>
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">else{</font></p>
      <p><font face="����" COLOR="#800000">clearTimeout(timer);</font></p>
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">function MM_callJS(jsStr) { //v2.0</font></p> 
      <p><font face="����" COLOR="#800000">return eval(jsStr)</font></p> 
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">//--&gt;</font></p>
      <p><font face="����" COLOR="#0000c0">&lt;/script&gt;</font></p>
      <p><font face="����" COLOR="#0000c0">&lt;SCRIPT language=javascript&gt;</font></p> 
      <p><font face="����" COLOR="#800000">&lt;!--</font></p>
      <p><font face="����" COLOR="#800000">function mOvr(src,clrOver) {</font></p> 
      <p><font face="����" COLOR="#800000">if (!src.contains(event.fromElement)) {</font></p> 
      <p><font face="����" COLOR="#800000">src.bgColor = clrOver;</font></p> 
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">function mOut(src,clrIn) {</font></p> 
      <p><font face="����" COLOR="#800000">if (!src.contains(event.toElement)) {</font></p> 
      <p><font face="����" COLOR="#800000">src.bgColor = clrIn;</font></p> 
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">}</font></p>
      <p><font face="����" COLOR="#800000">// --&gt;</font></p> 
      <p><font face="����" COLOR="#0000c0">&lt;/SCRIPT&gt;</font></p>
      <p><font face="����">��</font></td>                 
</tr>
</table>
