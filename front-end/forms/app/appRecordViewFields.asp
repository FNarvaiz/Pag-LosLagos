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

%>
