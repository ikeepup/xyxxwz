<!--
function ShowListPage(page,Pcount,TopicNum,maxperpage,strLink,ListName){
	var alertcolor = '#FF0000';
	maxperpage=Math.floor(maxperpage);
	TopicNum=Math.floor(TopicNum);
	page=Math.floor(page);
	var n,p;
	if ((page-1)%5==0) {
		p=(page-1) /5
	}else{
		p=(((page-1)-(page-1)%5)/5)
	}
	if(TopicNum%maxperpage==0) {
		n=TopicNum/maxperpage;
	}else{
		n=(TopicNum-TopicNum%maxperpage)/maxperpage+1;
	}
	document.write ('<table border="0" cellpadding="0" cellspacing="1" class="Tableborder5">');
	document.write ('<form method=post action="?pcount='+Pcount+strLink+'">');
	document.write ('<tr align="center">');
	document.write ('<td class="tabletitle1" title="'+ListName+'">&nbsp;'+ListName+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="总数">&nbsp;'+TopicNum+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="每页">&nbsp;'+maxperpage+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="页次">&nbsp;'+page+'/'+Pcount+'页&nbsp;</td>');
	if (page==1){
		document.write ('<td class="tablebody1">&nbsp;<font face=webdings>9</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page=1'+strLink+'" title="首页"><font face=webdings>9</font></a>&nbsp;</td>');
	}
	if (p*5 > 0){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+p*5+strLink+'" title="上五页"><font face=webdings>7</font></a>&nbsp;</td>');
	}
	for (var i=p*5+1;i<p*5+6;i++){
		if (i==page){
			document.write ('<td class="tablebody2">&nbsp;<font class="normalTextSmall"><u><b>'+i+'</b></u></font>&nbsp;</td>');
		}else{
			document.write ('<td class="tablebody1">&nbsp;<a href="?page='+i+strLink+'">'+i+'</a>&nbsp;</td>');
		}
		if (i==n) break;
	}
	if (i<n){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+i+strLink+'" title="下五页"><font face=webdings>8</font></a>&nbsp;<td>');
	}
	if (page==n){
		document.write ('<td class="tablebody1">&nbsp;<Font face=webdings>:</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+n+strLink+'" title="尾页"><font face=webdings>:</font></a>&nbsp;</td>');
	}
		
	document.write ('<td class="tablebody1"><input class="PageInput" type=text name="page" size=1 maxlength=10  value="'+page+'"></td>');
	document.write ('<td class="tablebody1"><input type=submit value=Go name=submit class="PageInput"></td>');
	document.write ('</tr>');
	document.write ('</form></table>');
}
function ShowHtmlPage(page,Pcount,TopicNum,maxperpage,strLink,ExtName,ListName){
	var alertcolor = '#FF0000';
	maxperpage=Math.floor(maxperpage);
	TopicNum=Math.floor(TopicNum);
	page=Math.floor(page);
	var n,p;
	if ((page-1)%5==0) {
		p=(page-1) /5
	}else{
		p=(((page-1)-(page-1)%5)/5)
	}
	if(TopicNum%maxperpage==0) {
		n=TopicNum/maxperpage;
	}else{
		n=(TopicNum-TopicNum%maxperpage)/maxperpage+1;
	}
	document.write ('<table border="0" cellpadding="0" cellspacing="1" class="Tableborder5">');
	document.write ('<form method=post>');
	document.write ('<tr align="center">');
	document.write ('<td class="tabletitle1" title="'+ListName+'">&nbsp;'+ListName+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="总数">&nbsp;'+TopicNum+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="每页">&nbsp;'+maxperpage+'&nbsp;</td>');
	document.write ('<td class="tabletitle1" title="页次">&nbsp;'+page+'/'+Pcount+'页&nbsp;</td>');
	if (page==1){
		document.write ('<td class="tablebody1">&nbsp;<font face=webdings>9</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+subjoin(1)+ExtName+'" title="首页"><font face=webdings>9</font></a>&nbsp;</td>');
	}
	if (p*5 > 0){
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+subjoin(p*5)+ExtName+'" title="上五页"><font face=webdings>7</font></a>&nbsp;</td>');
	}
	for (var i=p*5+1;i<p*5+6;i++){
		if (i==page){
			document.write ('<td class="tablebody2">&nbsp;<font class="normalTextSmall"><u><b>'+i+'</b></u></font>&nbsp;</td>');
		}else{

			document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+subjoin(i)+ExtName+'">'+i+'</a>&nbsp;</td>');
		}
		if (i==n) break;
	}
	if (i<n){
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+subjoin(i)+ExtName+'" title="下五页"><font face=webdings>8</font></a>&nbsp;<td>');
	}
	if (page==n){
		document.write ('<td class="tablebody1">&nbsp;<Font face=webdings>:</font>&nbsp;</td>');
	}else{
		document.write ('<td class="tablebody1">&nbsp;<a href="'+strLink+subjoin(n)+ExtName+'" title="尾页"><font face=webdings>:</font></a>&nbsp;</td>');
	}
	document.write ('<td class="tabletitle1" title="转到">&nbsp;GO&nbsp;</td>');
	document.write ('<td class="tablebody1"><select class="PageInput" name="page" size="1" onchange="javascript:window.location=this.options[this.selectedIndex].value;">');
	for (var i=1;i<TopicNum;i++){
		if (i==page){
			document.write ('<option value="'+strLink+subjoin(i)+ExtName+'" selected>第'+i+'页</option>');
		}else{
			document.write ('<option value="'+strLink+subjoin(i)+ExtName+'">第'+i+'页</option>');
		}
		if (i==n) break;
	}
	document.write ('</select></td>');
	document.write ('</tr>');
	document.write ('</form></table>');
}
function subjoin(what) {
	if (what < 10){
		str = '00'+what;
	}else{
			if (what>9 && what<100)
			{
				str = '0'+what;
			}else{
				str = what;
			}
	}
	return str;
}
//-->