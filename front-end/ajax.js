var xmlHttp;

function createXmlHttpObject()
{
  if (!xmlHttp)
	{
    try
  	{
      xmlHttp = new ActiveXObject("Msxml2.xmlHttp");
    }
  	catch (e)
    {
      try
    	{
        xmlHttp = new ActiveXObject("Microsoft.xmlHttp");
      }
      catch (e)
    	{
    	  try
    		{
          xmlHttp = new XMLHttpRequest();
    		}
    		catch (e)
    		{
    		  xmlHttp = null;
    			alert("Su navegador no tiene capacidades suficientes para mostrar este sitio.");
    		}
  		}
  	}
	}
}

function ajaxBusy()
{
  return(xmlHttp != null);
}

var container = null;
var indicator = null;
var onResponseReady = "";
var ajaxResponseText = "";
var autoIndicator = false;

function ajaxAbort()
{
  if (xmlHttp)
  {
	  xmlHttp.onreadystatechange = null;
	  xmlHttp.abort();
		xmlHttp = null;
		container = null;
    autoIndicator = false;
	}
}

function ajaxGetText(url, containerId, indicatorId, onResponse)
{
  try
	{
    if (!xmlHttp)
  	{
      createXmlHttpObject();
      if (xmlHttp)
      {
        if (containerId)
				{
				  container = document.getElementById(containerId);
          if (!container)
					{
            alert("ajax GET: el contenedor de destino no existe (" + containerId + ")");
						xmlHttp = null;
						return(false);
					}
				}
        onResponseReady = onResponse;
        if (indicatorId)
        {
          ajaxLoadingSignalOn(indicatorId);
          autoIndicator = true;
        }
        xmlHttp.onreadystatechange = function ()
        {
          if (xmlHttp.readyState == 4)
          {
            if (container)
						{
						  container.innerHTML = xmlHttp.responseText;
    				  container = null;
						}
						else
						{
						  ajaxResponseText = xmlHttp.responseText;
						}
     			  xmlHttp = null;
            if (autoIndicator)
            {
              ajaxLoadingSignalOff();
              autoIndicator = false;
            }
    				if (onResponseReady)
    				{
              var s = onResponseReady;
    				  onResponseReady = "";
							if (typeof(s) == "function")
							  s()
							else
  							eval(s);
    				}
          }
        }
    		xmlHttp.open("GET", url, true);
        xmlHttp.send(null);
    	}
  	}
  }
	catch (e)
	{ 
	  alert("ajaxGetText: " + e.name + " - " + e.message);
	}
}

function ajaxSynchronousGetText(url, containerId, indicatorId)
{
  try
	{
    if (!xmlHttp)
  	{
      createXmlHttpObject();
      if (xmlHttp)
      {
        if (containerId)
				{
				  container = document.getElementById(containerId);
          if (!container)
					{
            alert("ajax GET: el contenedor de destino no existe (" + containerId + ")");
						xmlHttp = null;
						return(false);
					}
				}
        if (indicatorId)
        {
          ajaxLoadingSignalOn(indicatorId);
          autoIndicator = true;
        }
    		xmlHttp.open("GET", url, false);
        xmlHttp.send(null);
        if (xmlHttp.readyState == 4)
        {
          if (container)
  				{
  				  container.innerHTML = xmlHttp.responseText;
  				  container = null;
  				}
  				else
  				  ajaxResponseText = xmlHttp.responseText;
   			  xmlHttp = null;
          if (autoIndicator)
          {
            ajaxLoadingSignalOff();
            autoIndicator = false;
          }
        }
    	}
  	}
  }
	catch (e)
	{ 
	  alert("ajaxGetText: " + e.name + " - " + e.message);
	}
}

function ajaxPostText(url, containerId, params, indicatorId, onResponse)
{
  try
	{
    if (!xmlHttp)
  	{
      createXmlHttpObject();
      if (xmlHttp)
      {
        if (containerId)
				{
				  container = document.getElementById(containerId);
          if (!container)
					{
            alert("ajax GET: el contenedor de destino no existe (" + containerId + ")");
						xmlHttp = null;
						return(false);
					}
				}
        onResponseReady = onResponse;
        if (indicatorId)
        {
          ajaxLoadingSignalOn(indicatorId);
          autoIndicator = true;
        }
        xmlHttp.onreadystatechange = function()
        {
          if (xmlHttp.readyState == 4)
    			{
            if (container)
						{
						  container.innerHTML = xmlHttp.responseText;
    				  container = null;
						}
						else
						{
						  ajaxResponseText = xmlHttp.responseText;
						}
    			  xmlHttp = null;
            if (autoIndicator)
            {
              ajaxLoadingSignalOff();
              autoIndicator = false;
            }
    				if (onResponseReady)
    				{
              var s = onResponseReady;
    				  onResponseReady = "";
							if (typeof(s) == "function")
							  s()
							else
  							eval(s);
  					}
    			}
        }
        xmlHttp.open("POST", url, true);
        
        //Send the proper header information along with the request
        xmlHttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        xmlHttp.setRequestHeader("Content-length", params.length);
        xmlHttp.setRequestHeader("Connection", "close");
        xmlHttp.send(params);
    	}
		}
  }
	catch (e)
	{ 
	  alert("ajaxPostText: " + e.name + " - " + e.message);
	}
}

function addField(queryString, name, value) { 
	if (queryString.length > 0) { 
		queryString += "&";
	}
	queryString += encodeURIComponent(name) + "=" + encodeURIComponent(value);
	return queryString;
}

function serializeForm (theform)
{
	var els = theform.elements;
	var len = els.length;
	var queryString = "";
	for (var i=0; i<len; i++) 
	{
		var el = els[i];
		if (!el.disabled) 
		{
			switch(el.type) 
			{
				case 'text': case 'password': case 'hidden': case 'textarea': 
					queryString = addField(queryString, el.name, el.value);
					break;
				case 'select-one':
					if (el.selectedIndex>=0) 
						queryString = addField(queryString, el.name, el.options[el.selectedIndex].value);
					break;
				case 'select-multiple':
					for (var j=0; j<el.options.length; j++) 
						if (el.options[j].selected)
							queryString = addField(queryString, el.name, el.options[j].value);
					break;
				case 'checkbox':
				  var multiple = false;
        	for (var j=0; j<len; j++) 
      		  if (j != i && els[j].name == el.name) {
						  multiple = true;
							break;
						}
					if (multiple) {
  					if (el.checked)
  						queryString = addField(queryString, el.name, el.value);
					}
					else {
  					if (el.checked)
  						queryString = addField(queryString, el.name, '1');
  					else
  						queryString = addField(queryString, el.name, '0');
					}
					break;
				case 'radio':
					if (el.checked)
						queryString = addField(queryString, el.name, el.value);
			}
		}
	}
	return queryString;
}

function ajaxSubmit(theform, containerId, loadingDivId, onResponse)
{
  ajaxPostText(theform.action, containerId, serializeForm(theform), loadingDivId, onResponse);
}

// Loading signal handling =============================================================================================

function ajaxLoadingSignalOn(indicatorId)
{
  indicator = document.getElementById(indicatorId);
  if (indicator)
    indicator.style.visibility = "visible";
}

function ajaxLoadingSignalOff()
{
  if (indicator)
  {
    indicator.style.visibility = "hidden";
    indicator = null;
  }
}

