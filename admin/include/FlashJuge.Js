function CheckForm(form1)
{
	if (!validateSubmit()) return (false);
	if (form1.title.value == "")
	{
		alert("请输入标题名称!");
		form1.title.focus();
		return (false);
	}
	if (form1.ClassID.value==""){
		alert("该一级分类已经有下属分类，请选择其下属分类！");
		form1.ClassID.focus();
		return (false);
	}
	if (form1.ClassID.value=="0"){
		alert("该分类是外部连接，不能添加内容！");
		form1.ClassID.focus();
		return (false);
	}
	if (form1.filesize.value==""){
		alert("动画大小不能为空！");
		form1.filesize.focus();
		return false;
	}
	form1.content.value=IframeID.document.body.innerHTML; 
	MessageLength=IframeID.document.body.innerHTML.length;
	if(MessageLength<2){alert("动画简介不能小于2个字符！");return false;}
}