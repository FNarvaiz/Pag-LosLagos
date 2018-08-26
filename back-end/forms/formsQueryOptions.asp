<%

function renderStandardQueryOptions
  dim queryTypes: queryTypes = eval(formId & "QueryTypeNames")
  dim querySearchFieldNames: querySearchFieldNames = eval(formId & "QuerySearchFieldNames")
  dim querySearchFields: querySearchFields = eval(formId & "QuerySearchFields")
  dim queryLimits: queryLimits = eval(formId & "QueryLimitNames")
  dim i
  %>
  <div id="<%= formId %>QueryOptions" class="queryOptions">
    <%
    if not isEmpty(querySearchFields) then
      %>
      <table cellpadding="0" cellspacing="0" id="<%= formId %>QuerySearchControl">
        <%
          if uBound(querySearchFields) > 0 then
            %>
            <tr>
              <td valign="top">
                <div class="QuerySearchFieldNamesSelectContainer" id="<%= formId %>QuerySearchFieldNames">
                  <select class="combobox querySearchFieldNamesCombobox" size="1" id="<%= formId %>QuerySearchFieldNamesCombobox">
                  <%
                  for i = 0 to uBound(querySearchFields)
                    %>
                    <option value="<%= i %>" <%= renderBooleanAttr("selected", i = querySearchField) %>><%= querySearchFieldNames(i) %></option>
                    <%
                  next
                  %>
                  </select>
                </div>
              </td>
            </tr>
            <%
          end if
        %>
        <tr>
          <td valign="top">
            <input type="hidden" id="<%= formId %>QuerySearchField" value="<%= querySearchField %>">
            <input id="<%= formId %>QuerySearchValue" class="querySearchValue" type="text" maxlength="100"
              value="<%= querySearchValue %>" onkeydown="querySearchChanged(event, '<%= formId %>')" onFocus="this.select()">
            <input id="<%= formId %>querySearchGoBtn" class="queryOptionsButton" type="button"
              value="Buscar" onclick="querySearchGo('<%= formId %>')">
            <input id="<%= formId %>QuerySearchClearBtn" class="queryOptionsButton" type="button"
              value="Limpiar" onclick="querySearchClear('<%= formId %>')">
          </td>
        </tr>
      </table>
      <%
    end if
  
    if not isEmpty(queryTypes) then
      %>
      <table cellpadding="0" cellspacing="0" id="<%= formId %>QueryTypeControl" >
        <tr>
          <td class="queryOptionsLabel" id="<%= formId %>QueryTypeLabel">Consulta</td>
          <td>
            <div class="queryOptionsSelectContainer" id="<%= formId %>QueryType">
              <select class="combobox queryOptionsCombobox" size="1" id="<%= formId %>QueryTypeCombobox" onchange="refreshForm('<%= formId %>')">
              <%
              for i = 0 to uBound(queryTypes)
                %>
                <option value="<%= i %>" <%= renderBooleanAttr("selected", i = queryType) %>><%= queryTypes(i) %></option>
                <%
              next
              %>
              </select>
            </div>
          </td>
        </tr>
      </table>
      <%
    end if
    
    if not isEmpty(queryLimits) then
      %>
      <table cellpadding="0" cellspacing="0" id="<%= formId %>QueryLimitControl">
        <tr>
          <td class="queryOptionsLabel" id="<%= formId %>QueryLimitLabel">Mostrar</td>
          <td>
            <div class="queryOptionsSelectContainer" id="<%= formId %>QueryLimit">
              <select class="combobox queryOptionsCombobox" size="1" id="<%= formId %>QueryLimitCombobox" onchange="refreshForm('<%= formId %>')">
              <%
              for i = 0 to uBound(queryLimits)
                %>
                <option value="<%= i %>" <%= renderBooleanAttr("selected", i = queryLimit) %>><%= queryLimits(i) %></option>
                <%
              next
              %>
              </select>
            </div>
          </td>
        </tr>
      </table>
      <%
    end if
    
    %>
  </div>
  <%
end function

function standardPrepareQueryOptions
  dim queryTypes: queryTypes = eval(formId & "QueryTypeSearchConditions")
  if not isEmpty(queryTypes) then
    formGridViewShowTools = true
    addSearchExpr(queryTypes(queryType))
  end if
  dim querySearchFields: querySearchFields = eval(formId & "QuerySearchFields")
  if not isEmpty(querySearchFields) then
    formGridViewShowTools = true
    addSearchValues querySearchFields(querySearchField), querySearchValue
  end if
  dim queryLimitClauses: queryLimitClauses = eval(formId & "QueryLimitClauses")
  if isEmpty(queryLimitClauses) then
    if defaultQueryLimit > 0 then 
      gridViewQueryLimitClause = "TOP " & defaultQueryLimit
    end if
  else
    formGridViewShowTools = true
    gridViewQueryLimitClause = queryLimitClauses(queryLimit)
  end if
end function

%>
