<%

function renderDatePicker
  %>
    <div id="datePicker" onclick="return(false)">
      <div id="monthPicker">
        <select size="1" id="datePickerMonthComboBox" onclick="event.cancelBubble=true;" onchange="datePickerSetMonth(this)">
          <option value="1">ENE</option>
          <option value="2">FEB</option>
          <option value="3">MAR</option>
          <option value="4">ABR</option>
          <option value="5">MAY</option>
          <option value="6">JUN</option>
          <option value="7">JUL</option>
          <option value="8">AGO</option>
          <option value="9">SET</option>
          <option value="10">OCT</option>
          <option value="11">NOV</option>
          <option value="12">DIC</option>
        </select>
        <select size="1" id="datePickerYearComboBox" onclick="event.cancelBubble=true;" onchange="datePickerSetYear(this)">
          <option value="<%= year(date) - 2 %>"><%= year(date) - 2 %></option>
          <option value="<%= year(date) - 1 %>"><%= year(date) - 1 %></option>
          <option value="<%= year(date) %>"><%= year(date) %></option>
          <option value="<%= year(date) + 1 %>"><%= year(date) + 1 %></option>
        </select>
        <span class="datePickerBtn" onclick="datePickerSetDate(new Date(), true); datePickerHide();">Hoy</span>
        <span class="datePickerBtn" onclick="datePickerSetDate(new Date(new Date().getTime() - 86400000), true); datePickerHide();">Ayer</span>
      </div>
      <table cellpadding="0" cellspacing="1" id="dayPicker" width="100%">
        <thead>
          <tr><th>D</th><th>L</th><th>M</th><th>M</th><th>J</th><th>V</th><th>S</th></tr>
        </thead>
        <tbody>
          <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
          <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
          <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
          <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
          <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
          <tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>
        </tbody>
      </table>
    </div>
  <%
end function

%>
