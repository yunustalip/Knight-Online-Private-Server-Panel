<!--

	var img1 = new Image();
	//plus image
	img1.src = "images/bar.gif";
	var img2 = new Image();
	//minus image
	img2.src = "images/bar.gif";
	
	//create expand menu				
	function doOutline() {
	  var srcId, srcElement, targetElement;
	  srcElement = window.event.srcElement;
	  if (srcElement.className.toUpperCase() == "LEVEL1" || srcElement.className.toUpperCase() == "FAQ") {
			 srcID = srcElement.id.substr(0, srcElement.id.length-1);
			 targetElement = document.all(srcID + "s");
			 srcElement = document.all(srcID + "i");
					
		if (targetElement.style.display == "none") {			
					 targetElement.style.display = "";
					 if (srcElement.className == "LEVEL1") srcElement.src = img2.src;
			} else {
					 targetElement.style.display = "none";
					 if (srcElement.className == "LEVEL1") srcElement.src = img1.src;
		 }
	  }
	}
					
	document.onclick = doOutline;
-->