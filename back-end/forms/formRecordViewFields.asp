
<%

const HTMLFieldNamePrefix = "field"
const HTMLFieldCurrentValueSuffix = "Value"
const HTMLFieldNewValueSuffix = "NewValue"
const HTMLFieldLabelSuffix = "Label"

function isFieldReadOnly(fieldIndex)
  if recordViewReadOnly or (usrAccessLevel < usrPermissionInsert) or (usrAccessLevel < usrPermissionUpdate and not inserting) or _
      (nullRecord and not inserting) then
	  isFieldReadOnly = true
  elseif isNull(recordViewReadOnlyFields) then
	  isFieldReadOnly = false
	else
	  isFieldReadOnly = recordViewReadOnlyFields(fieldIndex) or _
      (recordViewIdFieldIsIdentity and recordViewFields(recordViewCurrentField) = "ID")
  end if
end function

function renderBooleanAttr(attrName, booleanValue)
  if booleanValue then
    renderBooleanAttr = attrName & "=" & dQuotes(attrName)
	else
    renderBooleanAttr = ""
	end if
end function

function renderValidationAttribute
	if recordViewNullableFields(recordViewCurrentField) or isFieldReadOnly(recordViewCurrentField) then
	  renderValidationAttribute = ""
	else
	  renderValidationAttribute = "onblur=" & _
      dQuotes("checkNull('" & formId & "','field" & recordViewCurrentField & "')") & _
      " onkeyup=" & dQuotes("formElementKeyUp(event)") '  & " onchange=" & dQuotes("formValueChanged()")
	end if
end function

function renderRecordViewFieldLabel
  dim cssClass
  if recordViewDisabled then
    cssClass = "labelDisabled"
  elseif isFieldReadOnly(recordViewCurrentField) then
    cssClass = "labelReadOnly"
	elseif recordViewNullableFields(recordViewCurrentField) then
    cssClass = ""
  else
    cssClass = "labelRequired"
  end if
  dim elemId: elemId = formId & HTMLFieldNamePrefix & recordViewCurrentField & HTMLFieldLabelSuffix
  %>
	<div class="fieldLabel <%= cssClass %>" id="<%= elemId %>"
	  style="left: <%= recordViewLabelLeftPos %>px; top: <%= recordViewFieldTopPos %>px;
      width: <%= recordViewFieldLeftPos - recordViewLabelLeftPos - 6 %>px"><%= recordViewFieldLabels(recordViewCurrentField) %></div>
  <%
end function

function renderRecordViewFieldHiddenLabel
  dim elemId: elemId = formId & HTMLFieldNamePrefix & recordViewCurrentField & HTMLFieldLabelSuffix
  dim labelText: labelText = recordViewFieldLabels(recordViewCurrentField)
  %><div class="hidden" id="<%= elemId %>"><%= labelText %></div><%
end function

function renderHTMLFieldCurrentValueName(field)
  dim i
  if isNumeric(field) then
    i = field
  else
    i = getFieldIndex(field)
  end if
  renderHTMLFieldCurrentValueName = HTMLFieldNamePrefix & i & HTMLFieldCurrentValueSuffix
end function

function renderHTMLFieldNewValueName(field)
  dim i
  if isNumeric(field) then
    i = field
  else
    i = getFieldIndex(field)
  end if
  renderHTMLFieldNewValueName = HTMLFieldNamePrefix & i & HTMLFieldNewValueSuffix
end function

function renderHTMLFieldValue
	if isNull(fieldCurrentValues(recordViewCurrentField)) then
	  renderHTMLFieldValue = ""
	elseif varType(fieldCurrentValues(recordViewCurrentField)) = 11 then
    if fieldCurrentValues(recordViewCurrentField) then
      renderHTMLFieldValue = 1
    else
      renderHTMLFieldValue = 0
    end if
	else
  	renderHTMLFieldValue = replace(replace(replace(fieldCurrentValues(recordViewCurrentField), chr(34), "&quot;"), "<", "&lt;"), ">", "&gt;")
	end if
end function

function renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldCurrentValueName(recordViewCurrentField)
  %><input class="hidden" name="<%= elemName %>" type="text" maxlength = "100" value="<%= renderHTMLFieldValue %>"><%
end function

function renderRecordViewHiddenField
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %><input class="hidden" name="<%= elemName %>" type="text" maxlength="100" value="<%= renderHTMLFieldValue %>"><%
  renderRecordViewHiddenField = 0
end function

function renderRecordViewIdentityField
  renderRecordViewIdentityField = renderRecordViewHiddenField
end function

function renderFieldBackgroundColor
  if isFieldReadOnly(recordViewCurrentField) then
    renderFieldBackgroundColor = "background-color: transparent"
  else
    renderFieldBackgroundColor = ""
  end if
end function

function renderRecordViewLiteralField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="editbox" name="<%= elemName %>" type="text" maxlength="100" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px;
      <%= renderFieldBackgroundColor %>" 
      <%= renderBooleanAttr("disabled", recordViewDisabled) %>
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
    onFocus="this.select()">
  <%
  renderRecordViewLiteralField = recordViewDefaultFieldHeight
end function

function renderRecordViewPasswordField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="editbox" name="<%= elemName %>" type="password" maxlength="100" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px;
      <%= renderFieldBackgroundColor %>" 
      <%= renderBooleanAttr("disabled", recordViewDisabled) %>
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
    onFocus="this.select()">
  <%
  renderRecordViewPasswordField = recordViewDefaultFieldHeight
end function

function renderRecordViewNameField
  renderRecordViewNameField = renderRecordViewLiteralField
end function

function renderRecordViewTextField(lines)
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim h
  if len(lines) > 0 then
    h = lines * recordViewDefaultFieldHeight - 2
  else
    h = recordViewTextAreaHeight
  end if
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<textarea class="textarea" name="<%= elemName %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px;
      height: <%= h %>px; <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %> onFocus="this.select()"><%= renderHTMLFieldValue %></textarea>
  <%
  renderRecordViewTextField = h + 2
end function

function renderRecordViewNameTextField(lines)
  renderRecordViewNameTextField = renderRecordViewTextField(lines)
end function

function renderRecordViewInfoField(lines)
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim h
  if len(lines) > 0 then
    h = lines * recordViewDefaultFieldHeight - 2
  else
    h = recordViewTextAreaHeight
  end if
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<textarea class="textarea" name="<%= elemName %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px;
      height: <%= h %>px; <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %> onFocus="this.select()"><%= replace(renderHTMLFieldValue, "|", vbCrLf) %></textarea>
  <%
  renderRecordViewInfoField = h + 2
end function

function renderRecordViewNotesField(lines)
  renderRecordViewFieldHiddenLabel
  renderCurrentValueField
  dim h
  if len(lines) > 0 then
    h = lines * recordViewDefaultFieldHeight - 2
  else
    h = recordViewTextAreaHeight
  end if
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<textarea class="textarea" name="<%= elemName %>"
    style="left: <%= recordViewLabelLeftPos %>px; top: <%= recordViewFieldTopPos %>px; 
    width: <%= recordViewEditboxWidth + recordViewFieldLeftPos - recordViewLabelLeftPos %>px;
      height: <%= h %>px; <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>><%= renderHTMLFieldValue %></textarea>
  <%
  renderRecordViewNotesField = h + 2
end function

function renderRecordViewPostalCodeField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="editbox" name="<%= elemName %>" type="text" maxlength="10" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 70px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onFocus="this.select()">
  <%
  renderRecordViewPostalCodeField = recordViewDefaultFieldHeight
end function

function renderRecordViewNumericField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="12" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewNumericEditboxWidth %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewNumericField = recordViewDefaultFieldHeight
end function

function renderRecordViewVarNumericField(length)
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="<% = length %>" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= length * 8 + 6 %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewVarNumericField = recordViewDefaultFieldHeight
end function


function renderRecordViewNumericCode5Field
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="21" value="<%= zeroPad(renderHTMLFieldValue, 5) %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 40px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) )"
    onFocus="this.select()">
  <%
  renderRecordViewNumericCode5Field = recordViewDefaultFieldHeight
end function

function renderRecordViewNumericCode6Field
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="21" value="<%= zeroPad(renderHTMLFieldValue, 6) %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 45px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) )"
    onFocus="this.select()">
  <%
  renderRecordViewNumericCode6Field = recordViewDefaultFieldHeight
end function

function renderRecordViewBigNumericField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="21" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewNumericEditboxWidth * 3 %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewBigNumericField = recordViewDefaultFieldHeight
end function

function renderRecordViewDecimal2Field
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="decimalEditbox" name="<%= elemName %>" type="text" maxlength="12" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewNumericEditboxWidth %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==110) || (event.keyCode==188) || (event.keyCode==190) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewDecimal2Field = recordViewDefaultFieldHeight
end function

function renderRecordViewDecimal3Field
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="decimalEditbox" name="<%= elemName %>" type="text" maxlength="12" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewNumericEditboxWidth %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==110) || (event.keyCode==188) || (event.keyCode==190) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewDecimal3Field = recordViewDefaultFieldHeight
end function

function renderRecordViewMoneyField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="moneyEditbox" name="<%= elemName %>" type="text" maxlength="12" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewNumericEditboxWidth + 20 %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==110) || (event.keyCode==188) || (event.keyCode==190) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewMoneyField = recordViewDefaultFieldHeight
end function

function renderRecordViewMoney4Field
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="moneyEditbox" name="<%= elemName %>" type="text" maxlength="12" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewNumericEditboxWidth + 32 %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==110) || (event.keyCode==188) || (event.keyCode==190) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewMoney4Field = recordViewDefaultFieldHeight
end function

function renderRecordViewEANField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="editbox" name="<%= elemName %>" type="text" maxlength="13" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 89px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) )"
    onFocus="this.select()">
  <%
  renderRecordViewEANField = recordViewDefaultFieldHeight
end function

function renderRecordViewCUITField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="editbox" name="<%= elemName %>" type="text" maxlength="13" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 89px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode==109) )"
    onFocus="this.select()">
  <%
  renderRecordViewCUITField = recordViewDefaultFieldHeight
end function

function renderRecordViewDateField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  dim readOnly: readOnly = isFieldReadOnly(recordViewCurrentField)
  %>
    <input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="10" value="<%= renderHTMLFieldValue %>"
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewDateEditboxWidth %>px;
        <%= renderFieldBackgroundColor %>"
      <%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", readOnly) %>
      <%= renderBooleanAttr("disabled", recordViewDisabled) %>
      onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
        || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (34<event.keyCode && event.keyCode<40) 
        || (event.keyCode==46) || (event.keyCode==111) || (event.keyCode == 55 && event.shiftKey==true) )"
      onFocus="this.select();" onkeyup="if (this.value) datePickerHide(); else datePickerShow(this);">
  <%
  if not readOnly then
    %>
      <img src="forms/resource/btnDatePicker.png" class="anchor" title="Ver calendario"
        style="position: absolute; left: <%= recordViewFieldLeftPos + recordViewDateEditboxWidth + 3 %>px; top: <%= recordViewFieldTopPos %>px" 
        onclick="datePickerToggle(document.<%= formId %>.<%= elemName %>); event.cancelBubble=true;">
    <%
  end if
  renderRecordViewDateField = recordViewDefaultFieldHeight
end function

function renderRecordViewDateFieldPlus
  renderRecordViewDateFieldPlus = renderRecordViewDateField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
  <input class="button" type="button" value="Hoy"
	  style="left: <%= recordViewFieldLeftPos + recordViewDateEditboxWidth + 8 %>px; top: <%= recordViewFieldTopPos %>px; width: 30px"
    onclick="document.forms['<%= formId %>'].<%= elemName %>.value='<%= systemDateStr(date) %>';"
		<%= renderBooleanAttr("disabled", recordViewDisabled or isFieldReadOnly(recordViewCurrentField)) %>>
  <input class="button" type="button" value="Ayer"
	  style="left: <%= recordViewFieldLeftPos + recordViewDateEditboxWidth + 46 %>px; top: <%= recordViewFieldTopPos %>px; width: 30px"
    onclick="document.forms['<%= formId %>'].<%= elemName %>.value='<%= systemDateStr(dateAdd("d", -1, date)) %>';"
		<%= renderBooleanAttr("disabled", recordViewDisabled or isFieldReadOnly(recordViewCurrentField)) %>
    onFocus="this.select()">
  <%
end function

function renderRecordViewTimeField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="10" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewTimeEditboxWidth %>px;
      <%= renderFieldBackgroundColor %>"
		<%= renderValidationAttribute %> <%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>
		<%= renderBooleanAttr("disabled", recordViewDisabled) %>
    onkeydown= "return ( event.ctrlKey || event.altKey || (47<event.keyCode && event.keyCode<58 && event.shiftKey==false) 
      || (95<event.keyCode && event.keyCode<106) || (event.keyCode==8) || (event.keyCode==9) || (event.keyCode>34 && event.keyCode<40) 
      || (event.keyCode==46) || (event.keyCode == 190 && event.shiftKey==true) )"
    onFocus="this.select()">
  <%
  renderRecordViewTimeField = recordViewDefaultFieldHeight
end function

function renderRecordViewTimeStampField
  renderRecordViewFieldLabel
  renderCurrentValueField
  formatFieldCurrentValue(recordViewCurrentField)
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
	<input class="numericEditbox" name="<%= elemName %>" type="text" maxlength="19" value="<%= renderHTMLFieldValue %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewDateTimeEditboxWidth %>px;
      background-color: transparent"
		readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
  <%
  renderRecordViewTimeStampField = recordViewDefaultFieldHeight
end function

function renderRecordViewBooleanField
  renderRecordViewFieldLabel
  dim fvalue
  if recordViewDisabled then fValue = false else fValue = fieldCurrentValues(recordViewCurrentField)
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
  <input class="hidden" name="<%= renderHTMLFieldCurrentValueName(recordViewCurrentField) %>" type="checkbox" 
    <%= renderBooleanAttr("checked", fValue) %>>&nbsp;</input>
  <div class="checkbox" style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px;">
    <input type="checkbox" name="<%= elemName %>" <%= renderBooleanAttr("checked", fValue) %> 
		<%= renderBooleanAttr("readonly", isFieldReadOnly(recordViewCurrentField)) %>	
    <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
  </div>
  <%
  if isFieldReadOnly(recordViewCurrentField) then
    %>
    <div class="checkbox" 
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; background-color:white; opacity: .01; filter: alpha(opacity=1);">
    </div>
    <%
  end if
	renderRecordViewBooleanField = recordViewDefaultFieldHeight
end function

function renderRecordViewLookupField(paramsStr)
  renderRecordViewFieldLabel
  renderCurrentValueField
  if recordViewComboboxWidth = 0 then recordViewComboboxWidth = recordViewEditboxWidth
  dim paramList, lookupTable
  dim orderBy: orderBy = recordViewLookupFieldNameField
  dim searchExpr: searchExpr = ""
  dim onChangeEventHandler: onChangeEventHandler = ""
  dim dependantFieldName: dependantFieldName = ""
  paramList = split(paramsStr, ",")
  if isArray(paramList) then
    lookupTable = paramList(0)
    if uBound(paramList) > 0 then orderBy = paramList(1)
    if uBound(paramList) > 1 then searchExpr = paramList(2)
    if uBound(paramList) > 2 then onChangeEventHandler = paramList(3)
    if uBound(paramList) > 3 then dependantFieldName = paramList(4)
  else
    lookupTable = paramsStr
  end if
  dim s: s = ""
  dim b
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  if isFieldReadOnly(recordViewCurrentField) then
    if len(fieldCurrentValues(recordViewCurrentField)) > 0 then
      s = "&nbsp;&nbsp;(sin asignar)"
      if not recordViewDisabled then
        b = dbConnect
        if dbGetData("SELECT " & recordViewLookupFieldNameField & " FROM " & lookupTable & " WHERE ID=" & fieldCurrentValues(recordViewCurrentField)) then
          s = rs(recordViewLookupFieldNameField)
        end if
        dbReleaseData
        if b then dbDisconnect
      end if
    end if
    dim readOnlyElemName: readOnlyElemName = HTMLFieldNamePrefix & recordViewCurrentField & "ReadOnly"
    %>
    <input class="hidden" name="<%= elemName %>" type="text" maxlength="100"
      value="<%= renderHTMLFieldValue %>">
  	<input class="editbox" name="<%= readOnlyElemName %>" type="text" maxlength="100" value="<%= s %>"
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewComboboxWidth %>px; background-color: transparent"
  		readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
    <%
  else
    if len(onChangeEventHandler) > 0 then
      s = ""
      if len(dependantFieldName) > 0 then
        s = ", document." & formId & "." & renderHTMLFieldNewValueName(getFieldIndex(dependantFieldName))
      end if
      onChangeEventHandler = "onchange=" & dQuotes(onChangeEventHandler & "(" & sQuotes(formId) & ", this" & s & ")")
    end if
    %>
    <div class="selectContainer" 
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewComboboxWidth %>px">
      <select class="combobox" size="1" name="<%= elemName %>" style="width: <%= recordViewComboboxWidth + 3 %>px"
        <%= onChangeEventHandler %> <%= renderValidationAttribute %> <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
        <option value="">&nbsp;&nbsp;(sin asignar)</option>
        <%
        if not recordViewDisabled then
          if len(searchExpr) > 0 then
'            if len(fieldCurrentValues(recordViewCurrentField)) > 0 then
'              searchExpr = " WHERE (" & searchExpr & ") OR ID=" & fieldCurrentValues(recordViewCurrentField)
'            else
              searchExpr = " WHERE " & searchExpr
'            end if
          else
            searchExpr = ""
          end if
          b = dbConnect
          dbGetData("SELECT TOP 1000 ID, " & recordViewLookupFieldNameField & " FROM " & lookupTable & searchExpr & " ORDER BY " & orderBy)
          do while not rs.EOF
            if isNull(fieldCurrentValues(recordViewCurrentField)) then
              s = ""
            else
              s = renderBooleanAttr("selected", rs("ID") = fieldCurrentValues(recordViewCurrentField))
            end if
            %>
            <option <%= s %> value="<%= rs("ID") %>"><%= rs(recordViewLookupFieldNameField) %></option>
            <%
            rs.moveNext
          loop
          dbReleaseData
          if b then dbDisconnect
        end if
        %>
      </select>
    </div>
    <%
  end if
  renderRecordViewLookupField = recordViewDefaultFieldHeight
end function

function renderRecordViewEnumField(optionList)
  renderRecordViewFieldLabel
  renderCurrentValueField
  if recordViewComboboxWidth = 0 then recordViewComboboxWidth = recordViewEditboxWidth
  dim options, s: s = ""
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  if isFieldReadOnly(recordViewCurrentField) then
    if not nullRecord then
      s = "&nbsp;&nbsp;(sin asignar)"
      if not recordViewDisabled and not isNull(fieldCurrentValues(recordViewCurrentField)) then
        options = split(optionList, ",")
        s = options(fieldCurrentValues(recordViewCurrentField))
      end if
    end if
    dim readOnlyElemName: readOnlyElemName = HTMLFieldNamePrefix & recordViewCurrentField & "ReadOnly"
    %>
    <input class="hidden" name="<%= elemName %>" type="text" maxlength="100" value="<%= s %>">
  	<input class="editbox" name="<%= readOnlyElemName %>" type="text" maxlength="100" value="<%= s %>"
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewComboboxWidth %>px; background-color: transparent"
  		readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
    <%
  else
    %>
    <div class="selectContainer" 
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewComboboxWidth %>px">
      <select class="combobox" size="1" name="<%= elemName %>" style="width: <%= recordViewComboboxWidth + 3 %>px"
        <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
        <option value="">&nbsp;&nbsp;(sin asignar)</option>
        <%
        if not recordViewDisabled then
          options = split(optionList, ",")
          dim i
          for i = 0 to uBound(options)
            if isNull(fieldCurrentValues(recordViewCurrentField)) then
              s = ""
            else
              s = renderBooleanAttr("selected", i = cInt(fieldCurrentValues(recordViewCurrentField)))
            end if
            %>
            <option <%= s %> value="<%= i %>"><%= options(i) %></option>
            <%
          next
        end if
        %>
      </select>
    </div>
    <%
  end if
  renderRecordViewEnumField = recordViewDefaultFieldHeight
end function


function renderRecordViewOptionsField(optionList)
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  %>
  <div class="radioButtons" 
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px">
    <%
    dim options
    options = split(optionList, ",")
    dim i, j
    if isNull(fieldCurrentValues(recordViewCurrentField)) then
      j = -1
    else
      j = fieldCurrentValues(recordViewCurrentField)
      if varType(j) = vbBoolean then
        if j then
          j = 1
        else
          j = 0
        end if
      end if
    end if
    for i = 0 to uBound(options)
      %>
      <span class="radioLabel"><%= options(i) %></span><input type="radio" value="<%= i %>" name="<%= elemName %>" 
        <%= renderBooleanAttr("disabled", recordViewDisabled) %> <%= renderBooleanAttr("checked", i = j) %>>&nbsp;
    	<%
    next
    %>
  </div>
  <%
  renderRecordViewOptionsField = recordViewDefaultFieldHeight
end function

function renderRecordViewFileUploadField
  recordViewFieldTopPos = recordViewFieldTopPos + 8
  renderRecordViewFieldLabel
  dim uploadBtnDisabled: uploadBtnDisabled = nullRecord or recordViewDisabled or usrAccessLevel = usrPermissionReadOnly or _
    (recordViewReadOnly and usrAccessLevel <> usrPermissionUnrestricted)
  dim uploadBtnClass
  if uploadBtnDisabled then
    uploadBtnClass = "buttonDisabled"
  else
    uploadBtnClass = "button"
  end if
  dim contentType: contentType = fieldCurrentValues(recordViewCurrentField)
  dim showBtnDisabled: showBtnDisabled = nullRecord or isNull(contentType)
  dim showBtnClass
  if showBtnDisabled then
    showBtnClass = "buttonDisabled"
  else
    showBtnClass = "button"
  end if
  dim s: s = replace(recordViewFields(recordViewCurrentField), "_CONTENTTYPE", "")
  %>
  <input class="<%= uploadBtnClass %>" type="button" value="Cargar ..."
	  style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 60px"
    onclick="fileUploadDialog('<%= formId %>','<%= s %>', '<%= fileUploadFormDataType %>', <%= fileUploadFormMaxFilesize %>)"
		<%= renderBooleanAttr("disabled", uploadBtnDisabled) %>>
  <input class="<%= showBtnClass %>" type="button" value="Ver"
    style="left: <%= recordViewFieldLeftPos + 65 %>px; top: <%= recordViewFieldTopPos %>px; width: 30px"
    onclick="window.open('<%= formsServer %>?sessionId=<%= sessionId %>&verb=binaryData&formId=<%= formId %>&keyFields=<%= keyFieldNames %>&keyValues=<%= keyFieldValues %>&recordId=<%= recordId %>&dbFieldBaseName=<%= s %>&t=<%= timer() %>')"
    <%= renderBooleanAttr("disabled", showBtnDisabled) %>>
  <%
	renderRecordViewFileUploadField = recordViewDefaultFieldHeight
end function

function renderRecordViewFilenameField
  recordViewFieldTopPos = recordViewFieldTopPos + 8
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim fValue: fValue = fieldCurrentValues(recordViewCurrentField)
  dim fValueStr
  if isNull(fValue) then fValueStr = "(Ninguno)" else fValueStr = fValue
  dim uploadBtnDisabled: uploadBtnDisabled = nullRecord or recordViewDisabled or usrAccessLevel = usrPermissionReadOnly or _
    (recordViewReadOnly and usrAccessLevel <> usrPermissionUnrestricted)
  dim uploadBtnClass
  if uploadBtnDisabled then
    uploadBtnClass = "buttonDisabled"
  else
    uploadBtnClass = "button"
  end if
  dim showBtnDisabled: showBtnDisabled = nullRecord or isNull(fValue)
  dim showBtnClass
  if showBtnDisabled then
    showBtnClass = "buttonDisabled"
  else
    showBtnClass = "button"
  end if
  dim s: s = replace(recordViewFields(recordViewCurrentField), "_FILENAME", "")
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  if recordViewEditboxWidth - 94 > 100 then
    %>
    <input name="<%= elemName %>" type="hidden" value="<%= fValue %>">
    <input class="fileUploadEditBox" name="<%= elemName %>Visible" type="text" maxlength="100"
      value="<%= fValueStr %>" readonly="readonly"
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth - 98 %>px">
    <input class="<%= uploadBtnClass %>" type="button" value="Cargar ..."
      style="left: <%= recordViewFieldLeftPos + recordViewEditboxWidth - 94 %>px; top: <%= recordViewFieldTopPos %>px; width: 60px"
      onclick="fileUploadDialog('<%= formId %>','<%= s %>', '', <%= fileUploadFormMaxFilesize %>)"
      <%= renderBooleanAttr("disabled", uploadBtnDisabled) %>>
    <input class="<%= showBtnClass %>" type="button" value="Ver"
      style="left: <%= recordViewFieldLeftPos + recordViewEditboxWidth - 30 %>px; top: <%= recordViewFieldTopPos %>px; width: 30px; padding: 0"
      onclick="window.open('<%= formsServer %>?sessionId=<%= sessionId %>&verb=binaryFile&formId=<%= formId %>&keyFields=<%= keyFieldNames %>&keyValues=<%= keyFieldValues %>&recordId=<%= recordId %>&dbFieldBaseName=<%= s %>&t=<%= timer() %>')"
      <%= renderBooleanAttr("disabled", showBtnDisabled) %>>
    <%
    renderRecordViewFilenameField = recordViewDefaultFieldHeight
  else
    %>
    <input name="<%= elemName %>" type="hidden" value="<%= fValue %>">
    <input class="fileUploadEditBox" name="<%= elemName %>Visible" type="text" maxlength="100"
      value="<%= fValue %>" readonly="readonly"
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px">
    <input class="<%= uploadBtnClass %>" type="button" value="Cargar ..."
      style="left: <%= recordViewFieldLeftPos + recordViewEditboxWidth - 94 %>px; top: <%= recordViewFieldTopPos + recordViewDefaultFieldHeight %>px; width: 60px"
      onclick="fileUploadDialog('<%= formId %>','<%= s %>', '', <%= fileUploadFormMaxFilesize %>)"
      <%= renderBooleanAttr("disabled", uploadBtnDisabled) %>>
    <input class="<%= showBtnClass %>" type="button" value="Ver"
      style="left: <%= recordViewFieldLeftPos + recordViewEditboxWidth - 30 %>px; top: <%= recordViewFieldTopPos + recordViewDefaultFieldHeight %>px; width: 30px; padding: 0"
      onclick="window.open('<%= formsServer %>?sessionId=<%= sessionId %>&verb=binaryFile&formId=<%= formId %>&keyFields=<%= keyFieldNames %>&keyValues=<%= keyFieldValues %>&recordId=<%= recordId %>&dbFieldBaseName=<%= s %>&t=<%= timer() %>')"
      <%= renderBooleanAttr("disabled", showBtnDisabled) %>>
    <%
    renderRecordViewFilenameField = 2 * recordViewDefaultFieldHeight
  end if
end function

function renderRecordViewHTMLField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  dim b: b = isFieldReadOnly(recordViewCurrentField)
  dim btnClass
  if b then
    btnClass = "buttonDisabled"
  else
    btnClass = "button"
  end if
  %>
  <input class="hidden" name="<%= elemName %>" type="text" maxlength = "100" value="<%= renderHTMLFieldValue %>">
  <input class="<%= btnClass %>" type="button" value="Editar ..."
	  style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 70px;"
    onclick="htmlEditorDialog('<%= formId %>','<%= elemName %>','<%= recordViewFieldLabels(recordViewCurrentField) %>')" 
    <%= renderBooleanAttr("disabled", recordViewDisabled or b) %>>
  <%
	renderRecordViewHTMLField = recordViewDefaultFieldHeight
end function

function renderRecordViewNoteField(lines)
  dim noteHeight
  noteHeight = lines * recordViewDefaultFieldHeight
  %>
	<div class="fieldNote" style="left: <%= recordViewLabelLeftPos %>; top: <%= recordViewFieldTopPos %>px;
    width: <%= recordViewFieldLeftPos + recordViewEditboxWidth - recordViewLabelLeftPos %>px; height: <%= noteHeight %>px">
    <%= fieldCurrentValues(recordViewCurrentField) %>
  </div>
  <%
  renderRecordViewNoteField = noteHeight
end function

function renderRecordViewNoteLabel(lines)
  dim noteHeight
  noteHeight = lines * recordViewDefaultFieldHeight
  %>
	<div class="fieldNote" style="left: <%= recordViewLabelLeftPos %>; top: <%= recordViewFieldTopPos %>px;
    width: <%= recordViewFieldLeftPos + recordViewEditboxWidth - recordViewLabelLeftPos %>px; height: <%= noteHeight %>px">
    <%= recordViewFieldLabels(recordViewCurrentField) %>
  </div>
  <%
  renderRecordViewNoteLabel = noteHeight
end function

%>
