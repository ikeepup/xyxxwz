<!--
function ShowListPage(page,Pcount,TopicNum,maxperpage,strLink,ListName){
	var alertcolor = '#FF0000';
	maxperpage=Math.floor(maxperpage);
	TopicNum=Math.floor(TopicNum);
	page=Math.floor(page);
	var n,p;
	if ((page-1)%10==0) {
		p=(page-1) /10
	}else{
		p=(((page-1)-(page-1)%10)/10)
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
	if (p*10 > 0){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+p*10+strLink+'" title="上十页"><font face=webdings>7</font></a>&nbsp;</td>');
	}
	for (var i=p*10+1;i<p*10+11;i++){
		if (i==page){
			document.write ('<td class="tablebody2">&nbsp;<font class="normalTextSmall"><u><b>'+i+'</b></u></font>&nbsp;</td>');
		}else{
			document.write ('<td class="tablebody1">&nbsp;<a href="?page='+i+strLink+'">'+i+'</a>&nbsp;</td>');
		}
		if (i==n) break;
	}
	if (i<n){
		document.write ('<td class="tablebody1">&nbsp;<a href="?page='+i+strLink+'" title="下十页"><font face=webdings>8</font></a>&nbsp;<td>');
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
//-->