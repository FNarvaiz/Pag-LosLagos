function toggleSectionItems(btnElem, sectionElemId)
{
  var sectionElem = document.getElementById(sectionElemId);
  if (sectionElem.className)
  {
    btnElem.src = btnElem.src.replace("expand", "collapse");
    sectionElem.className = '';
  }
  else
  {
    btnElem.src = btnElem.src.replace("collapse", "expand");
    sectionElem.className = 'hidden';
  }
}
