


function ResizeImage(objImage,swidth,sheight,maxWidth){ 
//三个参数 需满足高度和宽度再进行图片缩放，防止广告位高度宽度不足引起的误缩放
//swidth  满足缩放的宽度
//sheight  满足缩放的高度
//maxWidth将来图片缩放的宽度
try{ 
if(maxWidth>0){ 

if(objImage.width>swidth&&objImage.height>sheight){ 
objImage.width=maxWidth; 
    if (window.attachEvent) 
     {objImage.attachEvent('onclick', function(){try{window.open(objImage.src);}catch(e){window.open(objImage.src);}}); 
     objImage.attachEvent('onmouseover', function(){objImage.style.cursor='pointer';}); 
     } 
     if (window.addEventListener) 
     {objImage.addEventListener('click', function(){try{window.open(objImage.src);}catch(e){window.open(objImage.src);}},false); 
     objImage.addEventListener('mouseover', function(){objImage.style.cursor='pointer';},false); 
     }     

} 
}
}catch(e){}; 
}


function ResizeImage2(objImage,swidth,sheight,maxHeight){ 
//三个参数 需满足高度和宽度再进行图片缩放，防止广告位高度宽度不足引起的误缩放
//swidth  满足缩放的宽度
//sheight  满足缩放的高度
//maxWidth将来图片缩放的高度
try{ 
if(maxHeight>0){ 

if(objImage.width>swidth&&objImage.height>sheight){ 
objImage.height=maxHeight; 
    if (window.attachEvent) 
     {objImage.attachEvent('onclick', function(){try{window.open(objImage.src);}catch(e){window.open(objImage.src);}}); 
     objImage.attachEvent('onmouseover', function(){objImage.style.cursor='pointer';}); 
     } 
     if (window.addEventListener) 
     {objImage.addEventListener('click', function(){try{window.open(objImage.src);}catch(e){window.open(objImage.src);}},false); 
     objImage.addEventListener('mouseover', function(){objImage.style.cursor='pointer';},false); 
     }     

} 
}
}catch(e){}; 
}
