
function JugeComment(myform)
{
	if (document.myform.UserName.value==""){
		alert ("你的用户名不可为空！");
		document.myform.UserName.focus();
		return(false);
	}
	if (document.myform.content.value == "")
	{
		alert("评论内容不能为空！");
		document.myform.content.focus();
                return (false);
	}
}
function CheckAll(form) {  
	for (var i=0;i<form.elements.length;i++)  
	{  
		var e = form.elements[i];  
		if (e.name != 'chkall')  
		e.checked = true // form.chkall.checked;  
	}  
} 
 
function ContraSel(form) {
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
		if (e.name != 'chkall')
		e.checked=!e.checked;
	}
}
function bbimg(o){
	var zoom=parseInt(o.style.zoom, 10)||100;zoom+=event.wheelDelta/12;if (zoom>0) o.style.zoom=zoom+'%';
	return false;
}
function imgzoom(img,maxsize){
	var a=new Image();
	a.src=img.src
	if(a.width > maxsize * 4)
	{
		img.style.width=maxsize;
	}
	else if(a.width >= maxsize)
	{
		img.style.width=Math.round(a.width * Math.floor(4 * maxsize / a.width) / 4);
	}
	return false;
}

function storePage() {
	d=document;
	t=d.selection?(d.selection.type!='None'?d.selection.createRange().text:''):(d.getSelection?d.getSelection():'');
	void(vivi=window.open('http://vivi.sina.com.cn/collect/icollect.php?pid=52z.com&title='+escape(d.title)+'&url='+escape(d.location.href)+'&desc='+escape(t),'vivi','scrollbars=no,width=480,height=480,left=75,top=20,status=no,resizable=yes'));
	vivi.focus();
}