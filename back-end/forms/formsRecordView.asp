
<!--#include file="formRecordViewFields.asp"-->

<%

function renderRecordViewUsrField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim s
  if recordViewDisabled or isNull(fieldCurrentValues(recordViewCurrentField)) then
    s = ""
  else
    dim b: b = dbConnect
  	if dbGetData("SELECT NAME FROM USR WHERE ID=" & fieldCurrentValues(recordViewCurrentField)) then
      s = rs("NAME")
    else
      s = ""
    end if
    dbReleaseData
    if b then dbDisconnect
  end if
  %>
  <input class="hidden" name="<%= recordViewFields(recordViewCurrentField) %>NewValue" type="text" maxlength="100"
	  value="<%= fieldCurrentValues(recordViewCurrentField) %>">
	<input class="editbox" type="text" maxlength="100" value="<%= s %>"
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: 160px; background-color: transparent"
		<%= renderBooleanAttr("readonly", true) %> <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
  <%
  renderRecordViewUsrField = recordViewDefaultFieldHeight
end function

function renderRecordViewSeparator(separator)
  if not isNull(separator) then
    %>
    <div class="fieldSeparator" style="left: <%= recordViewLabelLeftPos - 4 %>; top: <%= recordViewFieldTopPos %>px;
      width: <%= recordViewFieldLeftPos + recordViewEditboxWidth - recordViewLabelLeftPos + 6 %>px"><%= separator %></div>
    <%
    renderRecordViewSeparator = recordViewDefaultFieldHeight
  else
    renderRecordViewSeparator = 0
  end if
end function

function renderDBFormControls
  %>
	<div class="canvas formRecordViewControls" id="<%= formId %>RecordViewControls">
  <%
	recordViewCurrentField = 0
	recordViewFieldTopPos = recordViewTopmostFieldTopPos
  dim i
  for i = 0 to UBound(recordViewFields)
    if not isNull(recordViewSeparators) and (recordViewFieldRenderFuncs(i) <> "renderRecordViewHiddenField") then
      recordViewFieldTopPos = recordViewFieldTopPos + renderRecordViewSeparator(recordViewSeparators(i))
    end if
 	  recordViewFieldTopPos = recordViewFieldTopPos + eval(recordViewFieldRenderFuncs(recordViewCurrentField))
		recordViewCurrentField = recordViewCurrentField + 1
	next
  %>
	</div>
  <%
end function

function renderDBFormToolbar
  dim i
  i = recordViewLabelLeftPos - 4
  %>
	<div class="canvas formRecordViewButtons" id="<%= formId %>RecordViewButtons">
    <%
    if useAuditData then
      %>
      <div class="recordViewAuditInfo">
        <%= auditUserName %><br><%= auditDate %>
      </div>
      <%
    end if
    if not recordViewDisabled then
      %>
      <div class="button" style="left: <%= i %>px"><img src="forms/resource/b-refresh.bmp" alt="Releer" onclick="refreshAllForms('<%= formId %>');"></div>
      <%
      i = i + 24
      if usrAccessLevel >= usrPermissionInsert and recordViewButtons(0) and (verb = "recordView") then
        %>
        <div class="button" style="left: <%= i %>px"><img src="forms/resource/b-add.bmp" alt="Agregar" onclick="formAddNew('<%= formId %>');"></div>
        <%
        i = i + 24
      end if
      if usrAccessLevel >= usrPermissionDelete and recordViewButtons(1) and (verb = "recordView") and not nullRecord and _
          not recordViewReadOnly then
        %>
        <div class="button" style="left: <%= i %>px"><img src="forms/resource/b-delete.bmp" alt="Eliminar" onclick="formDelete('<%= formId %>');"></div>
        <%
        i = i + 24
      end if
      if not recordViewReadOnly and (recordViewButtons(2) or inserting) and ( _
          ((usrAccessLevel >= usrPermissionUpdate) and (verb = "recordView") and not nullRecord) or _
          ((usrAccessLevel >= usrPermissionInsert) and (verb = "newRecordView"))) then
        %>
        <div class="button" style="left: <%= i %>px"><img src="forms/resource/b-save.bmp" alt="Guardar" onclick="formSave('<%= formId %>');"></div>
        <%
        i = i + 24
      end if
      dim s: s = ""
      select case usrAccessLevel
        case usrPermissionReadOnly: s = "L"
        case usrPermissionUnrestricted: s = "S/R"
      end select
      if len(s) > 0 then
        i = i + 4
        %>
        <div class="recordViewAccessLevel" style="left: <%= i %>px"><%= s %></div>
        <%
      end if
    end if
    %>
	</div>
  <%
end function

function systemDateStr(dateVal)
  if isNull(dateVal) then
    systemDateStr = ""
  else
    systemDateStr = zeroPad(day(dateVal), 2) & dbDateSeparator & zeroPad(month(dateVal), 2) & dbDateSeparator & year(dateVal)
  end if
end function

function systemDateFromNewValue(dateFieldNewValueIdx)
  systemDateFromNewValue = null
  dim d: d = fieldNewValues(dateFieldNewValueIdx)
  if len(d) > len(dbDatePrefix) then
    d = split(right(d, len(d) - len(dbDatePrefix)), dbDateSeparator)
    if uBound(d) >= 2 then
      systemDateFromNewValue = dateSerial(d(2), d(1), d(0))
    elseif uBound(d) = 1 then
      systemDateFromNewValue = dateSerial(year(date), d(1), d(0))
    else
      systemDateFromNewValue = dateSerial(year(date), month(date), d(0))
    end if
  end if
end function

function systemTimeStr(timeVal)
  systemTimeStr = zeroPad(hour(timeVal), 2) & dbTimeSeparator & zeroPad(minute(timeVal), 2)
end function

function formatFieldCurrentValue(fieldIdx)
  select case recordViewFieldRenderFuncs(fieldIdx)
    case "renderRecordViewMoneyField", "renderRecordViewDecimal2Field":
      fieldCurrentValues(fieldIdx) = formatDecimal(fieldCurrentValues(fieldIdx), 2)
    case "renderRecordViewMoney4Field":
      fieldCurrentValues(fieldIdx) = formatDecimal(fieldCurrentValues(fieldIdx), 4)
    case "renderRecordViewDateField", "renderRecordViewDateFieldPlus", "renderMvDateField":
      if not isNull(fieldCurrentValues(fieldIdx)) and len(fieldCurrentValues(fieldIdx)) > 0 then
        fieldCurrentValues(fieldIdx) = systemDateStr(fieldCurrentValues(fieldIdx))
      end if
    case "renderRecordViewTimeField":
      if not isNull(fieldCurrentValues(fieldIdx)) and len(fieldCurrentValues(fieldIdx)) > 0 then
        fieldCurrentValues(fieldIdx) = systemTimeStr(timeValue(fieldCurrentValues(fieldIdx)))
      end if
    case "renderRecordViewTimeStampField":
      if not isNull(fieldCurrentValues(fieldIdx)) and len(fieldCurrentValues(fieldIdx)) > 0 then
        fieldCurrentValues(fieldIdx) = systemDateStr(fieldCurrentValues(fieldIdx)) & " " & _ 
          systemTimeStr(timeValue(fieldCurrentValues(fieldIdx)))
      end if
    case "renderRecordViewDecimal3Field":
      fieldCurrentValues(fieldIdx) = formatDecimal(fieldCurrentValues(fieldIdx), 3)
  end select
end function

function getCurrentValues
  if isNull(recordViewFields) then exit function
	redim fieldCurrentValues(uBound(recordViewFields))
  if nullRecord or recordViewDisabled then exit function
  dim s, i, j
  if useAuditData then
    s = "(SELECT USRTABLE." & usrNameField & " FROM " & usrTable & " USRTABLE WHERE USRTABLE.ID=" & _
      formTable & "." & dbAuditUsrIdField & ") AS AUDIT_USR_NAME," & dbAuditDate & " AS AUDIT_DATE"
    j = 2
  else
    s = ""
    j = 0
  end if
	for i = 0 to uBound(recordViewFields)
    if len(s) > 0 then s = s & ","
		if recordViewDBFields(i) then
		  s = s & recordViewFields(i)
    else
      s = s & "NULL"
		end if
	next
  if len(s) = 0 then s = "*"
  dim b: b = dbConnect
  if getRecord(s, formTable, keyFieldNames, keyFieldValues, recordId) then
    if useAuditData then
      auditUserName = rs("AUDIT_USR_NAME")
      auditDate = replace(systemDateStr(rs("AUDIT_DATE")), dbDateSeparator, ".") & " " & _
        zeroPad(hour(rs("AUDIT_DATE")), 2) & ":" & zeroPad(minute(rs("AUDIT_DATE")), 2) ' & ":" & zeroPad(second(rs("AUDIT_DATE")), 2)
    else
      auditUserName = ""
      auditDate = ""
    end if
    for i = 0 to uBound(recordViewFields)
      fieldCurrentValues(i) = rs(i + j)
      formatFieldCurrentValue(i)
    next
  end if
  dbReleaseData
  if b then dbDisconnect
end function

function getDefaultValues
  if isNull(recordViewFields) then exit function
	redim fieldCurrentValues(uBound(recordViewFields))
  dim i
  for i = 0 to uBound(recordViewFields)
    fieldCurrentValues(i) = recordViewFieldDefaults(i)
    formatFieldCurrentValue(i)
  next
end function

function getNewValues
  if isNull(recordViewFields) then exit function
	redim fieldNewValues(uBound(recordViewFields))
	redim fieldChanged(uBound(recordViewFields))
  dim i, s, t, j, fieldRenderFunc, newValue, currentValue
  for i = 0 to UBound(recordViewFields)
    s = renderHTMLFieldNewValueName(i)
    t = renderHTMLFieldCurrentValueName(i)
    newValue = getStringParam(s, 0)
    if isNull(newValue) then 
      newValue = ""
    else
      newValue = replace(replace(replace(newValue, " ", "#\#"), "##\", ""), "#\#", " ")
    end if
    currentValue = getStringParam(t, 0)
    if isNull(currentValue) then currentValue = ""
    fieldChanged(i) = (newValue <> currentValue)
    fieldRenderFunc = recordViewFieldRenderFuncs(i)
    j = inStr(fieldRenderFunc, "(")
    if j > 0 then fieldRenderFunc = left(fieldRenderFunc, j - 1)
    select case fieldRenderFunc
      case "renderRecordViewBooleanField", "renderRecordViewNumericField", "renderRecordViewBigNumericField", _
          "renderRecordViewNumericCode5Field", "renderRecordViewLookupField", "renderRecordViewEnumField", "renderRecordViewOptionsField":
        fieldNewValues(i) = getNumericParam(s)
      case "renderRecordViewNameField", "renderRecordViewNameTextField":
        fieldNewValues(i) = nameCase(newValue)
      case "renderRecordViewMoneyField":
        fieldNewValues(i) = getNumericParam(s)
        if isNull(fieldNewValues(i)) then
          fieldNewValues(i) = dbMoneyPrefix
        else
          fieldNewValues(i) = dbMoneyPrefix & cStr(round(CCur(fieldNewValues(i)), 2))
        end if
      case "renderRecordViewMoney4Field":
        fieldNewValues(i) = getNumericParam(s)
        if isNull(fieldNewValues(i)) then 
          fieldNewValues(i) = dbMoneyPrefix
        else
          fieldNewValues(i) = dbMoneyPrefix & cStr(round(CCur(fieldNewValues(i)), 4))
        end if
      case "renderRecordViewDecimal2Field", "renderRecordViewDecimal3Field":
        fieldNewValues(i) = getNumericParam(s)
        if isNull(fieldNewValues(i)) then
          fieldNewValues(i) = dbDecimalPrefix
        else
          fieldNewValues(i) = dbDecimalPrefix & cStr(fieldNewValues(i))
        end if
      case "renderRecordViewDateField", "renderRecordViewDateFieldPlus", "renderMvDateField": 
        fieldNewValues(i) = dbDatePrefix & getDateTimeParam(s)
      case "renderRecordViewTimeField":
        fieldNewValues(i) = dbTimePrefix & getDateTimeParam(s)
      case else
        fieldNewValues(i) = newValue
    end select
    if isNull(fieldNewValues(i)) then fieldNewValues(i) = recordViewFieldDefaults(i)
	next
end function

function getKeyValue(fieldName)
  getkeyValue = -1
  dim i, j
  for i = 0 to uBound(keyFields)
    if fieldName = keyFields(i) then
      j = strToInteger(keyValues(i))
      if not isNull(j) then getkeyValue = j
      exit function
    end if
  next
end function

function getFieldIndex(fieldName)
  getFieldIndex = -1
  if isNull(recordViewFields) then exit function
  dim i
	for i = 0 to uBound(recordViewFields)
    if fieldName = recordViewFields(i) then
      getFieldIndex = i
      exit function
		end if
	next
end function

function setFieldNewValue(fieldNameOrIndex, value)
  dim fieldIndex
  if isNumeric(fieldNameOrIndex) then
    fieldIndex = cInt(fieldNameOrIndex)
  else
    fieldIndex = getFieldIndex(fieldNameOrIndex)
  end if
  if fieldIndex < 0 then exit function
  fieldNewValues(fieldIndex) = value
  if not inserting then fieldChanged(fieldIndex) = true
  recordViewReadOnlyFields(fieldIndex) = false
end function

function setFieldsReadOnly(fromField, toField)
  if isNull(fromField) then 
    fromField = 0
  elseif not isNumeric(fromField) then
    fromField = getFieldIndex(fromField)
  end if
  if isNull(toField) then 
    toField = uBound(recordViewReadOnlyFields)
  elseif not isNumeric(toField) then
    toField = getFieldIndex(toField)
  end if
  dim i
  for i = fromField to toField
    recordViewReadOnlyFields(i) = true
  next
end function

function renderStandardRecordView
  renderDBFormToolbar
  renderDBFormControls
end function

function renderRecordView
  if inserting then
    getDefaultValues
  else
    getCurrentValues
  end if
  %>
  <form name="<%= formId %>">
    <%
    eval(formRecordViewRenderFunc)
    %>
  </form>
  <%
end function

%>

