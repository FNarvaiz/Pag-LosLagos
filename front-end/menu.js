
var menu = new menuManager();

function menuManager()
{
  this.submenus = new Array();
	this.expandedHeights = new Array();
  this.expandedSubmenuDivId = "";
	this.expandedSubmenu = -1;
	this.expandingSubmenu = -1;
	this.collapsingSubmenu = -1;
  this.start = 0;
  this.duration = 1000;
  this.lapse = 10;
  this.expanded = true;
  this.busy = false;
  this.divId = '';
	this.setup = menuSetup;
	this.expandSubmenu = submenuExpand;
	this.getSubmenuIdx = getSubmenuIdx;
	this.doExpandCollapse = mDoExpandCollapse;
	this.doExpand = mDoExpand;
	this.doCollapse = mDoCollapse;
	this.collapseAll = collapseAll;
	this.collapseExpanded = collapseExpanded;
	this.onReady = null;
}

function menuSetup()
{
  for (var i = 0;; i++)
	{
	  var divSubmenu = document.getElementById("portfolioSubmenu" + (i + 1));
		if (!divSubmenu) break;
    this.submenus[i] = divSubmenu;
 	  this.expandedHeights[i] = divSubmenu.offsetHeight;
	}
	for (var i = 1; i < this.submenus.length; i++)
	{
 	  this.submenus[i].style.height = "1px";
 	  this.submenus[i].style.display = "none";
	}
	this.expandedSubmenu = 0;
	this.expandedSubmenuDivId = this.submenus[0].id;
}
function a(){
try{
 } catch (e) { alert("error: " + e.name + " - " + e.message); }
}

function collapseAll()
{
	for (var i = 0; i < this.submenus.length; i++)
 	  this.submenus[i].style.height = "1px";
	this.expandedSubmenu = -1;
	this.expandedSubmenuDivId = "";
}

function collapseExpanded()
{
  if (!this.busy && this.expandedSubmenu >= 0)
	{
	  this.busy = true;
		this.collapsingSubmenu = this.expandedSubmenu;
  	this.start = new Date();
  	setTimeout("mDoCollapse()", this.lapse);
	}
}

function getSubmenuIdx(divId)
{
  var idx = -1;
  for (var i = 0; i < this.submenus.length; i++)
	{
		if (this.submenus[i].id == divId)
		{
		  idx = i;
			break;
		}
	}
  return(idx);
}

function mDoExpandCollapse()
{
  menu.submenus[menu.expandingSubmenu].blur(); // Mozilla
	var elapsed = new Date().getTime() - menu.start.getTime();
	if (elapsed < menu.duration)
  {
	  var factor = (1 - Math.cos(3.14 * elapsed / menu.duration)) / 2;
	  var h = Math.round(menu.expandedHeights[menu.expandingSubmenu] * factor);
  	if (h <= 1)
  	  h = 1; // for IE
    menu.submenus[menu.expandingSubmenu].style.height = h + "px";
		if (menu.collapsingSubmenu >= 0)
		{
		  var h = Math.round(menu.expandedHeights[menu.collapsingSubmenu] * (1 - factor));
			if (h <= 1)
			  h = 1; // for IE
      menu.submenus[menu.collapsingSubmenu].style.height = h + "px";
		}
		setTimeout("mDoExpandCollapse()", menu.lapse);
	}
	else
	{
    menu.submenus[menu.expandingSubmenu].style.height = menu.expandedHeights[menu.expandingSubmenu] + "px";
		menu.expandedSubmenu = menu.expandingSubmenu;
	  menu.expandedSubmenuDivId = menu.submenus[menu.expandedSubmenu].id;
		menu.expandingSubmenu = -1;
		menu.collapsingSubmenu = -1;
		menu.busy = false;
		if (menu.onReady)
		  eval(menu.onReady);
	}
}

function submenuExpand(divId)
{
  if (!this.busy && (divId != this.expandedSubmenuDivId))
	{
	  this.busy = true;
		this.expandingSubmenu = this.getSubmenuIdx(divId);
    this.submenus[this.expandingSubmenu].style.display = "block";
		this.collapsingSubmenu = this.expandedSubmenu;
  	this.start = new Date();
  	setTimeout("mDoExpandCollapse()", this.lapse);
	}
	else if (menu.onReady)
		eval(menu.onReady);
}

function mDoCollapse()
{
	var elapsed = new Date().getTime() - menu.start.getTime();
	if (elapsed < menu.duration)
  {
	  var factor = (1 - Math.cos(3.14 * elapsed / menu.duration)) / 2;
    menu.submenus[menu.collapsingSubmenu].style.height = 
		  Math.round(menu.expandedHeights[menu.collapsingSubmenu] * (1 - factor)) + "px";
		setTimeout("mDoCollapse()", menu.lapse);
	}
	else
	{
    menu.submenus[menu.collapsingSubmenu].style.display = "none";
		if (menu.collapsingSubmenu == menu.expandedSubmenu)
		{
		  menu.expandedSubmenu = -1;
  	  menu.expandedSubmenuDivId = "";
		}
		menu.collapsingSubmenu = -1;
		menu.busy = false;
	}
}

function mDoExpand()
{
	var elapsed = new Date().getTime() - menu.start.getTime();
	if (elapsed < menu.duration) {
    menu.submenus[menu.expandingSubmenu].style.height =
		  (menu.expandedHeights[menu.expandingSubmenu] * elapsed / menu.duration) + "px";
		setTimeout("mDoExpand()", menu.lapse);
	}
	else
	{
    menu.submenus[menu.expandingSubmenu].style.height = menu.expandedHeights[menu.expandingSubmenu] + "px";
		menu.expandedSubmenu = menu.expandingSubmenu;
  	menu.expandedSubmenuDivId = menu.submenus[menu.expandedSubmenu].id;
		menu.expandingSubmenu = -1;
		busy = false;
	}
}

