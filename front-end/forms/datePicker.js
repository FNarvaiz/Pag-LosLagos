var datePickerEditField = null;
var datePickerBusy = false;
var datePickerCurrentDate = null;
var datePickerVisible = false;

function datePickerSetPosition()
{
  var point = elemPos(datePickerEditField);
  var pickerStyle = document.getElementById("datePicker").style;
  pickerStyle.left = point[1] + "px";
  if (point[0] - datePicker.offsetHeight - 3 > 0)
    pickerStyle.top = point[0] - datePicker.offsetHeight - 3 + "px";
  else
    pickerStyle.top = "0px";
}

function datePickerShow(inputElem)
{
  if (datePickerBusy) return;
  datePickerBusy = true;
  if (datePickerVisible)
    fade("datePicker", false, 
      function() {
        datePickerVisible = false;
        datePickerBusy = false;
        datePickerShow(inputElem);
      }
    );
  else
  {
    datePickerEditField = inputElem;
    if (datePickerEditField.value)
    {
      var dateParts = datePickerEditField.value.split("/");
      if (!isNaN(dateParts[0]) && !isNaN(dateParts[1]) && !isNaN(dateParts[2]))
        datePickerSetDate(new Date(dateParts[2],dateParts[1] - 1, dateParts[0]), true);
    }
    if (!datePickerCurrentDate) datePickerSetDate(new Date(), false);
    dayPickerUpdate();
    datePickerSetPosition();
    delay(100,
      function() {
        fade("datePicker", true, 
          function() {
            datePickerVisible = true;
            datePickerBusy = false;
          }
        );
      }                                                 
    );
  }
}

function datePickerHide()
{
  if (datePickerBusy || !datePickerVisible) return;
  datePickerBusy = true;
  fade("datePicker", false, 
    function() {
      datePickerCurrentDate = null;
      datePickerVisible = false;
      datePickerEditField.focus();
      datePickerEditField = null;
      datePickerBusy = false;
    }
  );
}

function datePickerToggle(elem)
{
  if (datePickerVisible && datePickerEditField == elem)
    datePickerHide();
  else
    datePickerShow(elem);
}

function datePickerSetDate(d, updateField)
{
  datePickerCurrentDate = d;
  document.getElementById("datePickerMonthComboBox").selectedIndex = d.getMonth();
  document.getElementById("datePickerYearComboBox").value = d.getFullYear();
  if (updateField){
    datePickerEditField.setAttribute('value',("0" + d.getDate()).slice(-2) + "/" + ("0" + (d.getMonth() + 1)).slice(-2) + "/" + d.getFullYear());
    if(document.getElementById("bookingsFormPlace") != undefined)
      bookingResourceSelectedHOME();
  }
}

function datePickerDaySelected(dayElem)
{
  datePickerCurrentDate.setDate(parseInt(dayElem.innerHTML));
  datePickerSetDate(datePickerCurrentDate, true);
  datePickerHide();
}

function dayPickerUpdate()
{
  var day = datePickerCurrentDate.getDate();
  var month = datePickerCurrentDate.getMonth();
  var year = datePickerCurrentDate.getFullYear();
  var monthDays = 32 - new Date(year, month, 32).getDate();
  if (day > monthDays) day = null;
  var d = new Date(year, month, 1);
  var cells = document.getElementById("dayPicker").getElementsByTagName("td");
  var rows = document.getElementById("dayPicker").getElementsByTagName("tr");
  for (var i = 0; i < cells.length; i++)
  {
    cell = cells[i];
    if (i < d.getDay() || i >= d.getDay() + monthDays)
    {
      cell.className = "";
      cell.innerHTML = "&nbsp;";
      cell.onclick = null;
    }
    else 
    {
      cell.innerHTML = 1 + i - d.getDay();
      cell.onclick = function(){datePickerDaySelected(this)};
      if (day && i == d.getDay() + day - 1)
        cell.className = "anchor selected";
      else
        cell.className = "anchor";
    }
  }
  for (var i = 1; i < rows.length; i++)
    if (d.getDay() + monthDays <= 7 * (i - 1))
      rows[i].style.display = "none";
    else
      rows[i].style.display = "";
}

function datePickerSetMonth(monthComboBox)
{
  var day = datePickerCurrentDate.getDate();
  var month = monthComboBox.selectedIndex;
  var year = datePickerCurrentDate.getFullYear();
  var monthDays = 32 - new Date(year, month, 32).getDate();
  if (day > monthDays) day = 1;
  datePickerCurrentDate = new Date(year, month, day);
  dayPickerUpdate();
}

function datePickerSetYear(yearComboBox)
{
  var day = datePickerCurrentDate.getDate();
  var month = datePickerCurrentDate.getMonth();
  var year = yearComboBox.value;
  var monthDays = 32 - new Date(year, month, 32).getDate();
  if (day > monthDays) day = 1;
  datePickerCurrentDate = new Date(year, month, day);
  datePickerUpdate();
}



