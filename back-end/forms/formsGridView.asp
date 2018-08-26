<%

function formatColumnDecimal(decNum, decimals)
  if isNull(decNum) then
    formatColumnDecimal = ""
  else
    formatColumnDecimal = formatNumber(round(strToDecimal(decNum), decimals), decimals)
  end if
end function

dim thisYear: thisYear = year(date)
dim thisMonth: thisMonth = month(date)

function formatDateColumn(dateValue)
  if isNull(dateValue) then
    formatDateColumn = ""
'  elseif dateValue = date then
'    formatDateColumn = "hoy"
'  elseif dateDiff("d", dateValue, date) = 1 then
'    formatDateColumn = "ayer"
  else
    if year(dateValue) = thisYear then
      'if month(dateValue) = thisMonth then
      '  formatDateColumn = zeroPad(day(dateValue), 2)
      'else
        formatDateColumn = zeroPad(day(dateValue), 2) & dbDateSeparator & zeroPad(month(dateValue), 2)
      'end if
    else
      formatDateColumn = zeroPad(day(dateValue), 2) & dbDateSeparator & zeroPad(month(dateValue), 2) & dbDateSeparator & _
        right(year(dateValue), 2)
    end if
  end if
end function

function formatTimeColumn(timeValue)
  if isNull(timeValue) then
    formatTimeColumn = ""
  else
    formatTimeColumn = zeroPad(hour(timeValue), 2) & ":" & zeroPad(minute(timeValue), 2)
  end if
end function

function getFileExt(filename)
  if not isNull(filename) then
    dim i: i = InStrRev(filename, ".")
    if i = 0 then
      getFileExt = ""
    else
      getFileExt = right(filename, len(filename) - i)
    end if
  else
    getFileExt = ""
  end if
end function

dim recognisedFileTypes: recognisedFileTypes = array("txt", "dwg", "doc", "ppt", "xls", "pdf", "zip", "rar", "jpg")

function iconFilename(fileType)
  dim b: b = false
  dim i
  for i = 0 to uBound(recognisedFileTypes)
    if recognisedFileTypes(i) = fileType then
      b = true
      exit for
    end if
  next
  if b then
    iconFilename = "forms/app/resource/lst_file_icon_" & fileType & ".png"
  else
    iconFilename = "forms/app/resource/lst_file_icon_generic.png"
  end if    
end function

function cellHTML(gridRows, columnIdx, rowIdx, rowId, currentRow)
  cellHTML = ""
  dim val: val = gridRows(columnIdx, rowIdx)
  dim a
  select case formGridColumnTypes(columnIdx)
    case formGridColumnHidden: 
      exit function
    case formGridColumnGeneral:
      cellHTML = val
    case formGridColumnGeneralCenter:
      cellHTML = val
    case formGridColumnName:
      cellHTML = nameCase(val)
    case formGridColumnBoolean:
      if val then
        cellHTML = "SI"
      else
        cellHTML = "&nbsp;"
      end if
    case formGridColumnPercent:
      cellHTML = formatColumnDecimal(val, 2) & "&nbsp;%"
    case formGridColumnDecimal2:
      cellHTML = formatColumnDecimal(val, 2)
    case formGridColumnDecimal3:
      cellHTML = formatColumnDecimal(val, 3)
    case formGridColumnCurrency:
      cellHTML = "<u>$</u>" & formatColumnDecimal(val, 2)
    case formGridColumnCurrency4:
      cellHTML = "<u>$</u>" & formatColumnDecimal(val, 4)
    case formGridColumnDate:
      cellHTML = formatDateColumn(val)
    case formGridColumnTime:
      cellHTML = formatTimeColumn(val)
    case formGridColumnDateTime:
      cellHTML = formatDateColumn(val) & " " & formatTimeColumn(val)
    case formGridColumnImage:
      if isNull(val) then
        cellHTML = "(s/imagen)"
        val = ""
      elseif currentRow then
        cellHTML = "<img width=" & dQuotes(formGridColumnWidths(columnIdx)) & _
          " src=" & dQuotes("forms/forms.asp?sessionId=" & sessionId & "&verb=binaryData&formId=" & _
          formId & "&keyFields=" & keyFieldNames & "&keyValues=" & keyFieldValues & "&recordId=" & rowId & _
          "&dbFieldBaseName=" & uCase(formGridColumns(columnIdx)) & "&t=" & timer()) & ">"
      else
        cellHTML = "<img width=" & dQuotes(formGridColumnWidths(columnIdx)) & _
          " src=" & dQuotes("forms/forms.asp?sessionId=" & sessionId & "&verb=binaryData&formId=" & _
          formId & "&keyFields=" & keyFieldNames & "&keyValues=" & keyFieldValues & "&recordId=" & rowId & _
          "&dbFieldBaseName=" & uCase(formGridColumns(columnIdx))) & ">"
      end if
    case formGridColumnFileIcon:
      if not isNull(val) then
        cellHTML = "<img width=" & dQuotes(formGridColumnWidths(columnIdx)) & " src=" & dQuotes(iconFilename(getFileExt(val))) & ">"
      end if
    case else:
      cellHTML = val
  end select
  if isNull(val) then cellHTML = "&nbsp;"
  cellHTML = "<td>" & cellHTML & "</td>"
end function

dim rowCount: rowCount = 0
dim rows: rows = null
dim rowsTotals: rowsTotals = null
dim idColIdx
dim currentRecordIsVisible

function renderStandardGridView
  dim selectedRow:  selectedRow = false
  dim auxId:        auxId = -1
  dim w:            w = 0
  dim i
  dim rowNumber: rowNumber = 0
  if not isNull(formGridColumns) then
    for i = 0 to uBound(formGridColumns)
      if isNull(formGridColumnWidths) then
        w = w + 40
      else
        w = w + formGridColumnWidths(i) + 1 '2px padding.
      end if
    next
    'w = w + 1
  end if
  dim h, s
  if formGridViewRowCount = 0 then
    s = ""
  elseif not isNull(formGridTotals) then
    s = "height: " & cInt(formGridRowHeight * (formGridViewRowCount - 1))
  elseif formGridViewShowFooter then
    s = "height: " & cInt(formGridRowHeight * (formGridViewRowCount - 1) + 3)
  else
    s = "height: " & cInt(formGridRowHeight * (formGridViewRowCount) + 3)
  end if
  %>
  <center>
    <%
    if formGridViewShowTools then
      %>
      <div id="<%= formId %>GridTools" class="formGridTools" style="width: <%= w + 19 %>px">
      <%
        eval(formQueryOptionsRenderFunc)
      %>
      </div>
      <%
    end if
    if not isNull(formGridColumnLabels) then
      %>
      <div id="<%= formId %>GridTitle" class="formGridTitle" style="width: <%= w + 19 %>px">
        <table cellpadding="1" cellspacing="0" width="<%= w %>" units="pixels">
        <tr>
          <%
          dim columnLabel, columnLabelAlign
          for i = 0 to uBound(formGridColumns)
            if formGridColumnWidths(i) > 0 then
              select case formGridColumnTypes(i)
                case formGridColumnDecimal2, formGridColumnDecimal3, formGridColumnCurrency, formGridColumnCurrency4:
                  columnLabelAlign = "right"
                case formGridColumnGeneralCenter, formGridColumnBoolean, formGridColumnPercent, formGridColumnDate, formGridColumnTime, formGridColumnDateTime, formGridColumnImage:
                  columnLabelAlign = "center"
                case else:
                  columnLabelAlign = "left"
              end select
              if gridViewReordering then
                columnLabel = formGridColumnLabels(i)
                if cStr(i + 1) = gridViewOrderBy then
                  columnLabel = columnLabel & "▲"
                elseif (i + 1) & " DESC" = gridViewOrderBy then
                  columnLabel = columnLabel & "▼"
                end if
                %>
                <td style="width: <%= formGridColumnWidths(i) %>px; text-align: <%= columnLabelAlign %>" class="anchor"
                  onclick="gridColumnClick('<%= formId %>',<%= i + 1 %>)"><%= columnLabel %></td>
                <%
              else
                %>
                <td width="<%= formGridColumnWidths(i) %>" style="text-align: <%= columnLabelAlign %>"><%= formGridColumnLabels(i) %></td>
                <%
              end if
            end if
          next
          %>
        </tr>
        </table>
      </div>
      <%
    end if
    dim mozTableStyle: mozTableStyle = ""
    dim ieColsStyle: ieColsStyle = ""
    dim colsFormat: colsFormat = ""
    dim visibleColCount: visibleColCount = 1
    if not isNull(rows) and not isNull(formGridColumns) and not isNull(formGridColumnWidths) and not isNull(formGridColumnTypes) then
      dim mozPartialTableStyle: mozPartialTableStyle = "#" & formId & "Grid tr > td"
      dim colStyle
      visibleColCount = 0
      for i = 0 to uBound(formGridColumns)
        if (formGridColumnTypes(i) <> formGridColumnHidden) and (formGridColumnWidths(i) > 0) then
          visibleColCount = visibleColCount + 1
          select case formGridColumnTypes(i)
            case formGridColumnDecimal2, formGridColumnDecimal3, formGridColumnCurrency, formGridColumnCurrency4, formGridColumnPercent:
              colStyle = "text-align: right"
            case formGridColumnGeneralCenter, formGridColumnBoolean, formGridColumnDate, formGridColumnTime, formGridColumnDateTime, formGridColumnImage:
              colStyle = "text-align: center"
            case else:
              colStyle = "text-align: left"
          end select
          if visibleColCount > 1 then mozPartialTableStyle = mozPartialTableStyle & " + td"
          mozTableStyle = mozTableStyle & mozPartialTableStyle & " { " & colStyle & " } " & vbCrLf
          colsFormat = colsFormat & "<col width=" & dQuotes(formGridColumnWidths(i)) & " style=" & dQuotes(colStyle) & ">"
        end if
      next
    end if
    %>
  	<div id="<%= formId %>Grid" class="formGrid" style="width: <%= w + 19 %>px; <%= s %>">
      <%
      if not recordViewDisabled then
        %>
        <style><%= mozTableStyle %></style>
        <table class="formGrid" cellpadding="2" cellspacing="0" style="width:<%= w %>">
        <%= colsFormat %>
        <%
        if isNull(rows) then
          if not formGridViewShowFooter then
            %>
            <tr class="formGridRow off"><td align="center" colspan="<%= visibleColCount %>">(Sin datos)</td></tr>
            <%
          end if
        else
          auxId = recordId
          dim itemClass, itemId
          dim rowIdx, lastRowIdx
          lastRowIdx = uBound(rows, 2)
          for rowIdx = 0 to lastRowIdx
            itemId = rows(idColIdx, rowIdx)
            if 0 > auxId then auxId = itemId
            if isNull(formGridRowCssClassColumn) then
              itemClass = ""
            else
              itemClass = rows(formGridRowCssClassColumn, rowIdx)
            end if
            if formGridViewSelectable or len(gridViewDetailModule) > 0 then
              if (itemId = auxId or (rowIdx = 0 and not currentRecordIsVisible)) and (len(gridViewDetailModule) = 0) then
                auxId = itemId
                if len(itemClass) > 0 then
                  itemClass = itemClass & " selRow"
                else
                  itemClass = "selRow"
                end if
                response.write("<tr class=" & dQuotes(itemClass) & " id=" & dQuotes(formId & "@" & itemId) & _
                  " onclick=" & dQuotes("grc(this)") & ">")
                selectedRow = true
              else
                if rowIdx mod 2 = 1 then itemClass = "oddRow " & itemClass
                if len(itemClass) > 0 then itemClass = " class=" & dQuotes(itemClass)
                response.write("<tr id=" & dQuotes(formId & "@" & itemId) & " onclick=" & dQuotes("grc(this)") & itemClass & ">")
              end if
            else
              response.write("<tr class=" & dQuotes("off") & ">")
            end if
            for i = 0 to uBound(formGridColumns)
              response.write(cellHTML(rows, i, rowIdx, itemId, auxId = itemId))
            next
            response.write("</tr>" & vbCrLf)
            rowNumber = rowNumber + 1
            if rowNumber = 2000 then
              rowNumber = 0
              response.flush
            end if
          next
        end if
        %>
        </table>
        <%
      end if
      %>
  	</div>
    <%
    if not recordViewDisabled then
      if not isNull(formGridTotals) then
        %>
        <div id="<%= formId %>GridTotals" class="formGrid" style="width: <%= w + 19 %>px; overflow: hidden; text-align: left">
          <style><%= replace(mozTableStyle, "#" & formId & "Grid", "#" & formId & "GridTotals") %></style>
          <table class="formGrid" cellpadding="2" cellspacing="1" style="width:<%= w %>">
            <%= colsFormat %>
            <tr>
            <%
              for i = 0 to uBound(formGridColumns)
                response.write(cellHTML(rowsTotals, i, 0, null, false))
              next
            %>
            </tr>
          </table>
        </div>
        <%
      elseif formGridViewShowFooter then
        if isNull(rows) then
          %>
          <div class="formGridFooter" style="width: <%= w + 19 %>px">(vacío)</div>
          <%
        elseif lastRowIdx = 0 then
          %>
          <div class="formGridFooter" style="width: <%= w + 19 %>px">1 ítem</div>
          <%
        else
          %>
          <div class="formGridFooter" style="width: <%= w + 19 %>px"><%= lastRowIdx + 1 %> ítems de <%= rowCount %></div>
          <%
        end if
      end if
    end if
    %>
  </center>
 	<%
  renderStandardGridView = auxId
end function

function renderGridView
  eval(formPrepareQueryOptionsFunc)
	dim searchExpr
  searchExpr = getSearchExpr(dbSearchModeAll, keyFieldNames, keyFieldValues, null)
  if searchExpr <> "" then searchExpr = " WHERE " & searchExpr

  dim i
  dim cols: cols = ""
  idColIdx = -1
  if not isNull(formGridColumns) then
    for i = 0 to uBound(formGridColumns)
      if formGridColumns(i) = "ID" then idColIdx = i
      if len(cols) > 0 then cols = cols & ","
      if not isNull(formGridColumnTypes) then
        if formGridColumnTypes(i) = formGridColumnImage then
          cols = cols & formGridColumns(i) & "_CONTENTTYPE"
        else
          cols = cols & formGridColumns(i)
        end if
      end if
    next
  end if
  if idColIdx < 0 then
    if len(cols) > 0 then cols = cols & ","
    cols = cols & "ID"
    idColIdx = uBound(formGridColumns) + 1
  end if

  dim queryOrderBy
  dim firstOrderColumn: firstOrderColumn = ""
  if gridViewReordering then
    if len(queryOrder) > 0 then gridViewOrderBy = queryOrder
    for i = 0 to uBound(formGridColumns)
      if formGridColumnWidths(i) > 0 then
        firstOrderColumn = cStr(i + 1)
        exit for
      end if
    next
    'dbLog("formId=" & formId & " - firstOrderColumn=" & firstOrderColumn)
    if len(firstOrderColumn) > 0 then
      if len(gridViewOrderBy) = 0 then
        gridViewOrderBy = firstOrderColumn
        queryOrderBy = firstOrderColumn
      elseif gridViewOrderBy <> firstOrderColumn and gridViewOrderBy <> firstOrderColumn & " DESC" then
        queryOrderBy = gridViewOrderBy & "," & firstOrderColumn
      else
        queryOrderBy = gridViewOrderBy
      end if
    else
      queryOrderBy = ""
    end if
  else
    queryOrderBy = gridViewOrderBy
  end if
  'dbLog("formId=" & formId & " - gridViewOrderBy=" & gridViewOrderBy & " - queryOrderBy=" & queryOrderBy)
  'response.write("formId=" & formId & " - gridViewOrderBy=" & gridViewOrderBy & " - queryOrderBy=" & queryOrderBy)
  'exit function
  dim selectedItem
  dim b: b = dbConnect
  currentRecordIsVisible = true
  if not nullRecord then
    if len(searchExpr) = 0 then
      currentRecordIsVisible = dbGetData("SELECT ID FROM " & formTable & " WHERE ID=" & recordId)
    else
      currentRecordIsVisible = dbGetData("SELECT ID FROM " & formTable & searchExpr & " AND ID=" & recordId)
    end if
    dbReleaseData
  end if
 	'dbLog("SELECT " & gridViewQueryLimitClause & " " & cols & " FROM " & formTable & searchExpr & " ORDER BY " & queryOrderBy)
 	'response.write("SELECT " & gridViewQueryLimitClause & " " & cols & " FROM " & formTable & searchExpr & " ORDER BY " & queryOrderBy)
  'exit function
 	if dbGetData("SELECT " & gridViewQueryLimitClause & " " & cols & " FROM " & formTable & searchExpr & " ORDER BY " & queryOrderBy) then
    rows = rs.GetRows
  end if
	dbReleaseData
  if not isNull(formGridTotals) then
    cols = ""
    for i = 0 to uBound(formGridColumns)
      if len(cols) > 0 then cols = cols & ", "
      if isNull(formGridTotals(i)) then
        cols = cols & "NULL"
      else
        cols = cols & formGridTotals(i) & "(" & formGridColumns(i) & ")"
      end if
    next
    'dbLog("SELECT " & cols & " FROM " & formTable & searchExpr)
    dbGetData("SELECT " & cols & " FROM " & formTable & searchExpr)
    rowsTotals = rs.getRows
    dbReleaseData
  elseif formGridViewShowFooter and not isNull(rows) then
    'dbLog("SELECT COUNT(*) FROM " & formTable & searchExpr)
    dbGetData("SELECT COUNT(*) FROM " & formTable & searchExpr)
    rowCount = rs(0)
    dbReleaseData
  end if
	if b then dbDisconnect
  selectedItem = eval(formGridViewRenderFunc)
 	%>
 	<div class="hidden" id="<%= formId %>SelectedItem"><%= selectedItem %></div>
 	<div class="hidden" id="<%= formId %>IsEmpty"><%= isNull(rows) %></div>
 	<div class="hidden" id="<%= formId %>CurrentRecordIsVisible"><%= currentRecordIsVisible %></div>
 	<div class="hidden" id="<%= formId %>QueryOrder"><%= gridViewOrderBy %></div>
 	<div class="hidden" id="<%= formId %>DetailModule"><%= gridViewDetailModule %></div>
  <div class="hidden" id="<%= formId %>LiteralKeyField"><%= literalKeyField %>
  <%
end function

%>