function doChange(objText, objDrop){
	if(document.myform.BriefTopic.selectedIndex<2){
		document.myform.BriefTopic.selectedIndex+=1;
	}
	if (!objDrop) return;
	var str = objText.value;
	var arr = str.split("|");
	var nIndex = objDrop.selectedIndex;
	objDrop.length=1;
	for (var i=0; i<arr.length; i++){
		objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
		document.myform.ImageUrl.value=arr[i];
	}
	objDrop.selectedIndex = nIndex;
}

function doSubmit(){
	if (document.myform.title.value==""){
		alert("文章标题不能为空！");
		return false;
	}
	if (document.myform.ClassID.value==""){
		alert("该一级分类已经有下属分类或该分类不允许添加内容，请选择其下属分类！");
		return false;
	}
	if (document.myform.ClassID.value=="0"){
		alert("该分类是外部连接，不能添加内容！");
		return false;
	}
	if (document.myform.Author.value==""){
		alert("文章作者不能为空！");
		return false;
	}
	if (document.myform.ComeFrom.value==""){
		alert("文章来源不能为空！");
		return false;
	}
	if (HtmlEditor1.getHTML()==""){
		alert("文章内容不能为空！");
		return false;
	}
	if (document.myform.AllHits.value==""){
		alert("初始点击数不能为空！");
		return false;
	}
	document.myform.submit();
}