<%

function renderRecordViewUserLookupField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim s: s = "&nbsp;&nbsp;(sin asignar)"
  dim b
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  dim id: id = -1
  if len(fieldCurrentValues(recordViewCurrentField)) > 0 then
    id = fieldCurrentValues(recordViewCurrentField)
    if not recordViewDisabled then
      b = dbConnect
      if dbGetData("SELECT NOMBRE FROM USUARIOS WHERE ID=" & id) then
        s = rs(0)
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
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth - 20 %>px; background-color: transparent"
    readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
  <img class="anchor" width="13" height="14" src="forms/resource/b-window.bmp" 
    style="position: absolute; left: <%= recordViewFieldLeftPos + recordViewEditboxWidth - 16 %>px; top: <%= recordViewFieldTopPos + 2 %>px"
    onclick="openForm(<%= id %>,'formAdminUsers')">
  <%
  renderRecordViewUserLookupField = recordViewDefaultFieldHeight
end function

function renderRecordViewSalesAccountLookupField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim s: s = "&nbsp;&nbsp;(sin asignar)"
  dim b
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  dim id: id = -1
  if len(fieldCurrentValues(recordViewCurrentField)) > 0 then
    id = fieldCurrentValues(recordViewCurrentField)
    if not recordViewDisabled then
      b = dbConnect
      if dbGetData("SELECT NOMBRE FROM CUENTAS WHERE ID=" & id) then
        s = rs(0)
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
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth - 20 %>px; background-color: transparent"
    readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
  <img class="anchor" width="13" height="14" src="forms/resource/b-window.bmp" 
    style="position: absolute; left: <%= recordViewFieldLeftPos + recordViewEditboxWidth - 16 %>px; top: <%= recordViewFieldTopPos + 2 %>px"
    onclick="openForm(<%= id %>,'formAdminSalesAccounts')">
  <%
  renderRecordViewSalesAccountLookupField = recordViewDefaultFieldHeight
end function

function renderRecordViewProjectLookupField
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim s: s = "&nbsp;&nbsp;(sin asignar)"
  dim b
  dim elemName: elemName = renderHTMLFieldNewValueName(recordViewCurrentField)
  dim id: id = -1
  if len(fieldCurrentValues(recordViewCurrentField)) > 0 then
    id = fieldCurrentValues(recordViewCurrentField)
    if not recordViewDisabled then
      b = dbConnect
      if dbGetData("SELECT NOMBRE FROM PROYECTOS WHERE ID=" & id) then
        s = rs(0)
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
    style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth - 20 %>px; background-color: transparent"
    readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
  <img class="anchor" width="13" height="14" src="forms/resource/b-window.bmp" 
    style="position: absolute; left: <%= recordViewFieldLeftPos + recordViewEditboxWidth - 16 %>px; top: <%= recordViewFieldTopPos + 2 %>px"
    onclick="openForm(<%= id %>,'formAdminTimeLine')">
  <%
  renderRecordViewProjectLookupField = recordViewDefaultFieldHeight
end function

function renderProjectVariantField
  renderProjectVariantField = renderRecordViewLookupField("PROYECTOS_VARIANTES,NOMBRE,ID_PROYECTO=" & _
    fieldCurrentValues(getFieldIndex("ID_PROYECTO")))
end function


' ================================================================================================

function renderRecordViewLookupSearchField(paramsStr)
  renderRecordViewFieldLabel
  renderCurrentValueField
  dim paramList, table, lookupTable
  dim orderBy: orderBy = recordViewLookupFieldNameField
  dim searchExpr: searchExpr = ""
  dim columnExpr: columnExpr = ""
  paramList = split(paramsStr, ",")
  if isArray(paramList) then
    lookupTable = paramList(0)
    if uBound(paramList) > 0 then orderBy = paramList(1)
    if uBound(paramList) > 1 then 
      if (len(paramList(2)) > 0) then searchExpr = paramList(2)
    end if
    if uBound(paramList) > 2 then 
      if (len(paramList(3)) > 0) then columnExpr = paramList(3)
    end if
  else
    lookupTable = paramsStr
  end if
  dim lookupValue: lookupValue = "&nbsp;&nbsp;(sin asignar)"
  dim b
  b = dbConnect
  if dbGetData("SELECT " & recordViewLookupFieldNameField & " FROM " & lookupTable & " WHERE ID=" & fieldCurrentValues(recordViewCurrentField)) then
    lookupValue = rs(0)
  end if
  dbReleaseData
  if isFieldReadOnly(recordViewCurrentField) then
    %>
    <input class="hidden" name="<%= recordViewFields(recordViewCurrentField) %>NewValue" type="text" maxlength="100"
      value="<%= renderHTMLFieldValue %>">
  	<input class="editbox" name="<%= recordViewFields(recordViewCurrentField) %>ReadOnly" type="text" maxlength="100" 
      value="<%= lookupValue %>"
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px; background-color: transparent"
  		readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>
    <%
  else
    %>
    <input class="hidden" name="<%= recordViewFields(recordViewCurrentField) %>NewValue" type="text" maxlength="100"
      value="<%= renderHTMLFieldValue %>">
  	<input class="editbox" name="<%= recordViewFields(recordViewCurrentField) %>ReadOnly" type="text" maxlength="100" 
      value="<%= lookupValue %>"
      style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos %>px; width: <%= recordViewEditboxWidth %>px"
  		readonly="readonly" <%= renderBooleanAttr("disabled", recordViewDisabled) %>>

    <input class="hidden" name="<%= recordViewFields(recordViewCurrentField) %>LookupTable" type="text" maxlength="100"
      value="<%= lookupTable %>">
    <input class="hidden" name="<%= recordViewFields(recordViewCurrentField) %>LookupOrder" type="text" maxlength="100"
      value="<%= orderBy %>">
    <input class="hidden" name="<%= recordViewFields(recordViewCurrentField) %>ColumnExpr" type="text" maxlength="1000"
      value="<%= columnExpr %>">
  	<input class="editbox" name="<%= recordViewFields(recordViewCurrentField) %>SearchValue" type="text" maxlength="100" 
      value="" style="left: <%= recordViewFieldLeftPos %>px; top: <%= recordViewFieldTopPos + recordViewDefaultFieldHeight %>px; width: 100px"
  		<%= renderBooleanAttr("disabled", recordViewDisabled) %> onkeyup="lookupSearchFieldChanged('<%= formId %>',this)">
    <select class="combobox" size="1" name="<%= recordViewFields(recordViewCurrentField) %>LookupList"
  	  style="position: absolute; left: <%= recordViewFieldLeftPos + 104 %>px; top: <%= recordViewFieldTopPos + recordViewDefaultFieldHeight %>px; width: <%= recordViewEditboxWidth - 104 %>px"
  		<%= renderValidationAttribute %> <%= renderBooleanAttr("disabled", recordViewDisabled) %>
      onchange="lookupSearchValueChanged('<%= formId %>',this)">
      <%
      renderRecordViewLookupSearchFieldOptions lookupTable, columnExpr, "", orderBy
      %>
    </select>
    <%
  end if
  if b then dbDisconnect
  renderRecordViewLookupSearchField = recordViewDefaultFieldHeight * 2
end function

function renderRecordViewLookupSearchFieldOptions(lookupTable, columnExpr, searchValue, lookupOrder)
  if not recordViewDisabled then
    dim searchExpr: searchExpr = ""
    if len(searchValue) > 0 then
      dim a: a = split(searchValue, " ")
      dim i
      for i = 0 to uBound(a)
        if len(a(i)) > 0 then
          if len(searchExpr) > 0 then searchExpr = searchExpr & " AND "
          searchExpr = searchExpr & " (" & columnExpr & " LIKE '%" & replace(a(i), "_", "[_]") & "%')"
        end if
      next
    end if
    if len(searchExpr) then
      searchExpr = " WHERE " & searchExpr & " "
    else
      searchExpr = ""
    end if
    dim b: b = dbConnect
    dbGetData("SELECT COUNT(*) FROM " & lookupTable & searchExpr)
    dim rowCount: rowCount = rs(0)
    dbReleaseData
    if rowCount > 100 then
      %>
    	<option value="">&nbsp;&nbsp;Más de 100 resultados (<%= rowCount %>). Modifique la búsqueda.</option>
      <%
    else
      dbGetData("SELECT TOP 100 ID, " & recordViewLookupFieldNameField & " FROM " & lookupTable & searchExpr & " ORDER BY " & lookupOrder)
      if rs.EOF then
        %>
        <option value="">&nbsp;&nbsp;Ningún resultado. Modifique la búsqueda.</option>
        <%
      else
        %>
        <option value="">&nbsp;&nbsp;seleccione un ítem (<%= rowCount %> resultados)</option>
        <%
        do while not rs.EOF
          dim s
          if isNull(fieldCurrentValues(recordViewCurrentField)) then
            s = ""
          else
            s = renderBooleanAttr("selected", rs("ID") = fieldCurrentValues(recordViewCurrentField))
          end if
          %>
          <option <%= s %> value="<%= rs("ID") %>"><%= rs(1) %></option>
          <%
          rs.moveNext
        loop
      end if
      dbReleaseData
    end if
    if b then dbDisconnect
  end if
end function

%>
