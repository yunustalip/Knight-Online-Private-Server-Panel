var xmlPost;
function xmlPost(url,text)
{
	var xmlspan = document.getElementById(text);
	var xmlhttp = new_xmlhttp();
	xmlhttp.open("get",url,true);
	xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded;charset=iso-8859-9");
	xmlhttp.send(text);
	xmlhttp.onreadystatechange = function() {
		
		if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
			setTimeout("refresh()", 7000); 
			setTimeout("loading()", 2500); 
			xmlspan.innerHTML = xmlhttp.responseText
				if (xmlspan.innerHTML == '') {
					
					rating.style.display = "block"
					ratings.style.display = "none";
					
					}
				else {
					rating.style.display = "none";
					ratings.style.display = "block";
				}
		}
		else
		{
			xmlspan.innerHTML = '<font class=orta>Hatalı İstek</font>';
		}
	}
	return false;
}
function refresh(){
   window.location.reload( false );
}
function loading(){
xmlspan.innerHTML = '<img src=tema/images/loading.gif /><font class=orta> Yükleniyor...</font>';
}
