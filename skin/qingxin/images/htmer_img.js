


function ResizeImage(objImage,swidth,sheight,maxWidth){ 
//�������� ������߶ȺͿ���ٽ���ͼƬ���ţ���ֹ���λ�߶ȿ�Ȳ��������������
//swidth  �������ŵĿ��
//sheight  �������ŵĸ߶�
//maxWidth����ͼƬ���ŵĿ��
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
//�������� ������߶ȺͿ���ٽ���ͼƬ���ţ���ֹ���λ�߶ȿ�Ȳ��������������
//swidth  �������ŵĿ��
//sheight  �������ŵĸ߶�
//maxWidth����ͼƬ���ŵĸ߶�
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
