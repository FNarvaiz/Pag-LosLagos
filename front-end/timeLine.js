var timeLineColumnSelected = -1;
var timeLineRowSelected = -1;

function setIconSelected(iconImage)
{
  iconImage.style.padding = "1px 2px";
  try
  {
    iconImage.style.boxShadow = "5px 5px 18px #000000";
  }
  catch(e)
  {
    iconImage.style.borderColor = "#cc0000";
  }
}

function setIconUnselected(iconImage)
{
  iconImage.style.padding = "";
  try
  {
    iconImage.style.boxShadow = "";
  }
  catch(e)
  {
    iconImage.style.borderColor = "";
  }
}

function timeLineColumnSelectedChanged(row, rowQty, column, detailKeyValues)
{
  if (timeLineColumnSelected != column)
  {
    if (timeLineColumnSelected > 0)
    {
      var e = document.getElementById("timeLineColumnLabel" + timeLineColumnSelected);
      track('cartelera-periodo', e.innerHTML.replace("<br>", "/"), 'Familia ' + usrName);
      e.style.fontWeight = "";
      document.getElementById("timeLineIcons" + timeLineColumnSelected).className = "timeLineIcons";
      for (var i = 1; i <= rowQty; i++)
      {
        var img = document.getElementById("timeLineIcon" + i + "_" + timeLineColumnSelected);
        if (img)
          setIconUnselected(img);
      }
    }
    timeLineColumnSelected = column;
    document.getElementById("timeLineColumnLabel" + timeLineColumnSelected).style.fontWeight = "bold";
    document.getElementById("timeLineIcons" + timeLineColumnSelected).className = "timeLineIconsSelected";
    for (var i = 1; i <= rowQty; i++)
    {
      var img = document.getElementById("timeLineIcon" + i + "_" + timeLineColumnSelected);
      if (img && i == row)
        setIconSelected(img);
    }
  }
  
  if (timeLineRowSelected != row)
  {
    if (timeLineRowSelected > 0)
      document.getElementById("timeLineRowLabel" + timeLineRowSelected).style.color = "";
    timeLineRowSelected = row;
    var e = document.getElementById("timeLineRowLabel" + timeLineRowSelected);
    track('cartelera-categoria', e.innerHTML, 'Familia ' + usrName);
    e.style.color = "#ffffff";
    for (var i = 1; i <= rowQty; i++)
    {
      img = document.getElementById("timeLineIcon" + i + "_" + timeLineColumnSelected);
      if (img)
      {
        if (i == row)
          setIconSelected(img);
        else
          setIconUnselected(img);
      }
    }
  }

  updateTimeLineDetail(row, detailKeyValues);
}
  
function updateTimeLineDetail(timeLineRow, detailKeyValues)
{
  ajaxGetText(appServer + "?lang=" + lang + "&sessionId=" + sessionId + "&content=timeLineDetail" + 
    "&timeLineRow=" + timeLineRow + "&detailKeyValues=" + detailKeyValues, "timeLineDetailIcons", loadingSignalId,
    function() {
      document.getElementById("timeLineDetailIcons").scrollLeft = 0;
    }
  );
}

