// Cookies FUNCTIONS =============================================================================================

function getCookie(cookieName)
{
  var result = "";
  if (document.cookie.length)
  {
    cookieStart = document.cookie.indexOf(cookieName + "=");
    if (cookieStart != -1)
    {
      cookieStart = cookieStart + cookieName.length + 1; 
      cookieEnd = document.cookie.indexOf(";", cookieStart);
      if (cookieEnd == -1)
        cookieEnd = document.cookie.length;
      result = unescape(document.cookie.substring(cookieStart, cookieEnd));
    }
  }
  return(result);
}

function setCookie(cookieName, value, expireDays)
{
  var expDate = new Date();
  expDate.setDate(expDate.getDate() + expireDays);
  document.cookie = cookieName+ "=" + escape(value) + ((expireDays == null) ? "" : ";expires=" + expDate.toGMTString());
}

// MISCELANEOUS FUNCTIONS =============================================================================================

function setupImg(imgSrc, containerDst, imgDst, borderWidth)
{
  var totalBorder = 0;
  if (borderWidth)
    totalBorder = borderWidth * 2;
  var imgS = imgSrc;
  if (typeof(imgS) == "string")
    imgS = document.getElementById(imgS);
  var container = containerDst;
  if (typeof(container) == "string")
    container = document.getElementById(container);
  if (!container) return;
  var imgD = imgDst;
  if (typeof(imgD) == "string")
    imgD = document.getElementById(imgD);
  var imgAR = imgS.offsetWidth / imgS.offsetHeight;
  var containerWidth = container.offsetWidth  - totalBorder
  var containerHeight = container.offsetHeight - totalBorder;
  if (imgAR > containerWidth / containerHeight)
  {
    imgD.style.left = "0px";
    var t = Math.round((containerHeight - (containerWidth / imgAR)) / 2);
    imgD.style.top = t + "px";
    imgD.style.width = containerWidth + "px";
    imgD.style.height = containerHeight - (t * 2) + "px";
  }
  else
  {
    var l = Math.round((containerWidth - (containerHeight * imgAR)) / 2)
    imgD.style.left = l + "px";
    imgD.style.top = "0px";
    imgD.style.width = containerWidth - (l * 2) + "px";
    imgD.style.height = containerHeight + "px";
  }
  if (imgD.src != imgS.src)
    imgD.src = imgS.src;
}

function setupImgViewerImg()
{
  var imgAR = currentDataViewerImgElem.offsetWidth / currentDataViewerImgElem.offsetHeight;
  var container = document.getElementById("dataViewer");
  var totalBorder = 12;
  var containerWidth = container.offsetWidth  - totalBorder
  var containerHeight = container.offsetHeight - totalBorder;
  var imgD = document.getElementById("dataViewerImg");
  if (imgAR > containerWidth / containerHeight)
  {
    imgD.style.left = "0px";
    var t = Math.round((containerHeight - (containerWidth / imgAR)) / 2);
    imgD.style.top = t + "px";
    imgD.style.width = containerWidth + "px";
    imgD.style.height = containerHeight - (t * 2) + "px";
  }
  else
  {
    var l = Math.round((containerWidth - (containerHeight * imgAR)) / 2)
    imgD.style.left = l + "px";
    imgD.style.top = "0px";
    imgD.style.width = containerWidth - (l * 2) + "px";
    imgD.style.height = containerHeight + "px";
  }
  imgD.style.visibility = "visible";
  document.getElementById("dataViewerImgTools").style.visibility = "visible";
}


function abortUIProcess()
{
  delayTask = null;
	ajaxAbort();
}

var delayTask = null;

function delay(time, task)
{
  delayTask = task;
  setTimeout(
	  function() {
		  if (typeof(delayTask) == "function")
			  delayTask();
			else
			  eval(delayTask);
		},
		time
	);
}

var scrollLeftFrom = 0;
var scrollLeftTo = 0;
var scrollTopFrom = 0;
var scrollTopTo = 0;
var scrollContainer = null;
var scrollStart = 0;
var scrollDuration = 800;

function scroll(elemId, direction, step)
{
	scrollContainer = document.getElementById(elemId);
  scrollLeftFrom = scrollContainer.scrollLeft;
  scrollTopFrom = scrollContainer.scrollTop;
  switch(direction)
	{
    case "left":
	    scrollLeftTo = scrollContainer.scrollLeft + step;
		  if (scrollLeftTo + scrollContainer.offsetWidth > scrollContainer.scrollWidth)
				scrollLeftTo = scrollContainer.scrollWidth - scrollContainer.offsetWidth
		  scrollTopTo = 0;
		  break;
    case "right":
  	  scrollLeftTo = scrollContainer.scrollLeft - step;
		  if (scrollLeftTo < 0)
			  scrollLeftTo = 0;
		  scrollTopTo = 0;
		  break;
    case "up":
		  scrollLeftTo = 0;
  	  scrollTopTo = scrollContainer.scrollTop + step;
		  if (scrollTopTo + scrollContainer.offsetHeight > scrollContainer.scrollHeight)
				scrollTopTo = scrollContainer.scrollHeight - scrollContainer.offsetHeight
		  break;
    case "down":
		  scrollLeftTo = 0;
  	  scrollTopTo = scrollContainer.scrollTop - step;
		  if (scrollTopTo < 0)
			  scrollTopTo = 0;
		  break;
	}
  scrollStart = new Date();
  doScroll();
}

function doScroll()
{
  var elapsed = new Date().getTime() - scrollStart.getTime();
  if (elapsed < scrollDuration)
	{
	  var factor = (1 - Math.cos(3.14 * elapsed / scrollDuration)) / 2;
    scrollContainer.scrollLeft = scrollLeftFrom + Math.round((scrollLeftTo - scrollLeftFrom) * factor);
    scrollContainer.scrollTop = scrollTopFrom + Math.round((scrollTopTo - scrollTopFrom) * factor);
    delay(10, "doScroll()");
	}
	else
	{
    scrollContainer.scrollLeft = scrollLeftTo;
    scrollContainer.scrollTop = scrollTopTo;
	}
}

var fadeElem = null;
var onFadeFinished = null;
var fadeStart = 0;
var fadeDuration = 200;
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

function centerElem(elem)
{
  elem.style.position = "relative";
  elem.style.left = Math.round((elem.offsetParent.clientWidth - elem.offsetWidth) / 2);
	elem.style.top = Math.round((elem.offsetParent.clientHeight - elem.offsetHeight) / 2);
  elem.style.visibility = "visible";
}

function formatDate(dateStr)
{
  if (lang == "EN")
  {
    var d = dateStr.split("/");
    var s = d[1];
    d[1] = d[0];
    d[0] = s;
    return(d.join("-"));
  }
  else
    return(dateStr);
}

function a (){
  try{
  } catch (e) { alert("error: " + e.name + " - " + e.message); }
}
