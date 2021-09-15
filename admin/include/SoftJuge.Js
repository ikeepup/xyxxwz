function ToRunsystem(addTitle) {
	var revisedTitle;
	var currentTitle;
	currentTitle = document.myform.RunSystem.value;
	revisedTitle = currentTitle+addTitle;
	document.myform.RunSystem.value=revisedTitle;
	document.myform.RunSystem.focus();
	return; 
}

function CheckForm(myform){
	if (!validateSubmit()) return (false);
	myform.content.value=getHTML(); 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<2){alert("软件简介不能小于2个字符！");return false;}

	if (myform.SoftName.value==""){
		alert("软件名称不能为空！");
		return false;
	}
	if (myform.ClassID.value==""){
		alert("该一级分类已经有下属分类，请选择其下属分类！");
		return false;
	}
	if (myform.ClassID.value=="0"){
		alert("该分类是外部连接，不能添加内容！");
		return false;
	}
	if (myform.RunSystem.value==""){
		alert("软件运行环境不能为空！");
		return false;
	}
	if (myform.SoftType.value==""){
		alert("软件类型不能为空！");
		return false;
	}
	if (myform.AllHits.value==""){
		alert("初始点击数不能为空！");
		return false;
	}
	//document.myform.submit1.disabled = true;
	//document.myform.submit();
}