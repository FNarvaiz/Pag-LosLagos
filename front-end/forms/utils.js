
function setSelectElementValue(selectElem, val) 
{
	for (var i = 0; i < selectElem.options.length; i++) 
    if (selectElem.options[i].value == val)
    {
      selectElem.selectedIndex = i;
      return;
    }
}

function imgResizeToFit(imgElem, w, h) {
  var r = imgElem.naturalWidth / imgElem.naturalHeight;
  if (r < w / h) {
    imgElem.style.height = h + "px";
    imgElem.style.width = Math.round(h * r) + "px";
  }
  else {
    imgElem.style.width = w + "px";
    imgElem.style.height = Math.round(w / r) + "px";
  }
}

function elemPos(elem) {
  var elem2 = elem;
  var curtop = 0;
  var curleft = 0;
  do {
    curleft += elem.offsetLeft - elem.scrollLeft;
    curtop += elem.offsetTop - elem.scrollTop;
    elem = elem.offsetParent;
    elem2 = elem2.parentNode;
    while (elem2!=elem) 
    {
      curleft -= elem2.scrollLeft;
      curtop -= elem2.scrollTop;
      elem2 = elem2.parentNode;
    }
  } while (elem.offsetParent)
  return [curtop, curleft];
}

var fadeElem = null;
var onFadeFinished = null;
var fadeStart = 0;
var fadeDuration = 100;
var fadeIn = false;

function fade(elem, doFadeIn, onFinished)
{
  if (typeof(elem) == "string")
    fadeElem = document.getElementById(elem);
	else
	  fadeElem = elem;
	if ((doFadeIn && fadeElem.style.visibility == "visible") || (!doFadeIn && fadeElem.style.visibility == "hidden"))
	{
  	if (typeof(onFinished) == "function")
  	  onFinished()
  	else
			eval(onFinished);
	}
	else
	{
  	fadeIn = doFadeIn;
  	fadeStart = new Date();
  	onFadeFinished = onFinished;
    fadeElem.style.visibility = "visible";
  	doFadeStep();
	}
}

function doFadeStep()
{
  var elapsed = new Date().getTime() - fadeStart.getTime();
  if (elapsed < fadeDuration)
	{
	  if (fadeIn)
		{
    	fadeElem.style.opacity = (elapsed / fadeDuration).toPrecision(2);
    	fadeElem.style.filter = "alpha(opacity=" + Math.round(100 * elapsed / fadeDuration) + ")";
		}
		else
		{
    	fadeElem.style.opacity = 1 - (elapsed / fadeDuration).toPrecision(2);
    	fadeElem.style.filter = "alpha(opacity=" + Math.round(100 - 100 * elapsed / fadeDuration) + ")";
		}
  	setTimeout("doFadeStep()", 10);
	}
  else
	{
	  if (fadeIn)
		{
    	//fadeElem.style.opacity = "1";
    	//fadeElem.style.filter = "alpha(opacity=100)";
    	fadeElem.style.opacity = "";
    	fadeElem.style.filter = "";
		}
		else
		{
    	//fadeElem.style.opacity = "0";
    	//fadeElem.style.filter = "alpha(opacity=0)";
    	fadeElem.style.opacity = "";
    	fadeElem.style.filter = "";
      fadeElem.style.visibility = "hidden";
    }
		var s = onFadeFinished;
		onFadeFinished = null;
  	if (typeof(s) == "function")
  	  s()
  	else
			eval(s);
	}
}


