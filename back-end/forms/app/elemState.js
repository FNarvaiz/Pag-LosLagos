
function setElemState(elem, cssStateKeyword)
{
  if (!elem.className) return;
  var cssClasses = elem.className.split(" ");
  var i = 1;
  while (i < cssClasses.length)
  {
    if (cssClasses[i].indexOf(cssClasses[0]) == 0)
      cssClasses.splice(i, 1);
    else
      i++;
  }
  cssClasses[cssClasses.length] = cssClasses[0] + cssStateKeyword;
  elem.className = cssClasses.join(" ");
}

function isElemSelected(elem)
{
  return(elem.className.indexOf(elem.className.split(" ")[0] + "Selected") >= 0);
}

function mouseOver(elem)
{
  if (!isElemSelected(elem))
    setElemState(elem, "Pointed");
}

function mouseOut(elem)
{
  if (!isElemSelected(elem))
    setElemState(elem, "Normal");
}

