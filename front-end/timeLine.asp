<%

function addSearchCondition(whereClause, condition)
  if len(whereClause) > 0 then
    addSearchCondition = whereClause & " OR " & condition
  else
    addSearchCondition = condition
  end if
end function

function renderTimeLineRowLabels
  if dbGetData("SELECT ID, NOMBRE FROM CATEGORIAS_LINEA_TIEMPO WHERE VIGENTE = 1 AND MOSTRAR = 1 ORDER BY ID") then
    %>
    <table cellpadding="0" cellspacing="0" width="100%" height="100%">
      <%
      do while not rs.EOF
        %>
        <tr><td id="timeLineRowLabel<%= rs("ID") %>"><%= rs("NOMBRE") %></td></tr>
        <%
        rs.moveNext
      loop
      %>
    </table>
    <%
  end if
  dbReleaseData
end function

function getVisibleRows
  if dbGetData("SELECT MOSTRAR FROM CATEGORIAS_LINEA_TIEMPO WHERE VIGENTE = 1 ORDER BY ID") then
    getVisibleRows = rs.getRows
  else
    getVisibleRows = array()
  end if
  dbReleaseData
end function

function getTimeLineData
  if dbGetData("SELECT L.ANIO, L.MES, L.ID_CATEGORIA FROM LINEA_TIEMPO L " & _
      "JOIN CATEGORIAS_LINEA_TIEMPO C ON C.ID = L.ID_CATEGORIA " & _
      "WHERE L.APROBADO=1 AND C.VIGENTE = 1 AND C.MOSTRAR = 1 " & _
      "GROUP BY L.ANIO, L.MES, L.ID_CATEGORIA ORDER BY L.ANIO, L.MES, L.ID_CATEGORIA") then
    getTimeLineData = rs.getRows
  else
    getTimeLineData = array()
  end if
  dbReleaseData
end function

function getCategoryQty
  dbGetData("SELECT COUNT(*) FROM CATEGORIAS_LINEA_TIEMPO WHERE VIGENTE = 1 AND MOSTRAR = 1")
  getCategoryQty = rs(0)
  dbReleaseData
end function

function renderTimeLineColumnBeginning(column, columnValue)
  %>
  <td class="timeLineColumn" id="timeLineColumn<%= column %>">
    <div class="timeLineColumnLabel" id="timeLineColumnLabel<%= column %>"><%= columnValue %></div>
    <div class="timeLineIcons" id="timeLineIcons<%= column %>">
      <table cellpadding="0" cellspacing="0" width="100%" height="100%">
  <%
end function

function renderTimeLineColumnEnding(row, visibleRows)
  dim i
  for i = row - 1 to uBound(visibleRows, 2)
    if visibleRows(0, i) then renderTimeLineNullIcon(i)
  next
  %></table></div></td><%
end function

function renderTimeLineNullIcon(row)
  %>
  <tr>
    <td align="center" valign="middle">
      <img class="invisible" src="front-end/resource/linea_tiempo_0.png" onload="this.style.visibility='visible';">
    </td>
  </tr>
  <%
end function

function aa
  %>
  
  <%
end function

function renderTimeLineIcon(id, row, rowQty, column, columnKeyValues)
  %>
  <tr>
    <td align="center" valign="middle">
      <img class="invisible anchor" id="<%= id %>" src="front-end/resource/linea_tiempo_<%= row %>.png" 
        onload="this.style.visibility='visible';"
        onclick="timeLineColumnSelectedChanged(<%= row %>,<%= rowQty %>,<%= column %>,'<%= columnKeyValues %>')">
    </td>
  </tr>
  <%
end function

function renderTimeLineGrid(timeLineData, rowQty)
  if uBound(timeLineData) < 0 then
    %>
    <table cellpadding="0" cellspacing="0" width="100%">
      <tr><td align="center" valign="middle" height="40"><%= eval("noDataMsg" & lang) %></td></tr>
    </table>
    <%
  else
    %>
    <table cellpadding="0" cellspacing="0">
      <tr>
      <%
        dim monthNames: monthNames = eval("monthNames" & lang)
        dim i, colKeyValues, categoryId
        dim currentColLabel: currentColLabel = ""
        dim row: row = 0
        dim column: column = 0
        dim imgId
        dim lastColumnTopImgId: lastColumnTopImgId = ""
        dim visibleRows: visibleRows = getVisibleRows
        for i = 0 to uBound(timeLineData, 2)
          colKeyValues = timeLineData(1, i) & "," & timeLineData(0, i)
          categoryId = timeLineData(2, i)
          if currentColLabel <> colKeyValues then
            if row > 0 then renderTimeLineColumnEnding row, visibleRows
            currentColLabel = colKeyValues
            row = 1
            column = column + 1
            renderTimeLineColumnBeginning column, monthNames(timeLineData(1, i) - 1) & "<br>" & timeLineData(0, i)
            lastColumnTopImgId = ""
          end if
          do while row < categoryId
            if visibleRows(0, row - 1) then renderTimeLineNullIcon(row)
            row = row + 1
          loop
          imgId = "timeLineIcon" & categoryId & "_" & column
          if len(lastColumnTopImgId) = 0 then lastColumnTopImgId = imgId
          renderTimeLineIcon imgId, row, rowQty, column, colKeyValues
          row = row + 1
        next
        if row > 0 then renderTimeLineColumnEnding row, visibleRows
        %>
      </tr>
    </table>  
    <%
  end if
  %>
  <div id="javascriptCode">
    timeLineColumnSelected = -1;
    timeLineRowSelected = -1;
    var e = document.getElementById("timeLineGridContainer");
    e.scrollLeft = e.scrollWidth - e.offsetWidth;
    var imgId = "<%= lastColumnTopImgId %>";
    if (imgId) 
      document.getElementById(imgId).onclick();
    else
      updateTimeLineDetail(-1, "-1,-1");
  </div>
  <%
end function

function renderTimeLine
  if not getUsrData then exit function
  %>
  <div id="dynPanelBg"></div>
  <div id="timeLineContainer">
    <div id="timeLineArrows">
      <div id="timeLineEndLeftArrow" onclick="window.scroll('timeLineGridContainer', 'right', 1000000)"></div>
      <div id="timeLineFastLeftArrow" onclick="window.scroll('timeLineGridContainer', 'right', 700)"></div>
      <div id="timeLineLeftArrow" onclick="window.scroll('timeLineGridContainer', 'right', 70)"></div>
      <div id="timeLineRightArrow" onclick="window.scroll('timeLineGridContainer', 'left', 70)"></div>
      <div id="timeLineFastRightArrow" onclick="window.scroll('timeLineGridContainer', 'left', 700)"></div>
      <div id="timeLineEndRightArrow" onclick="window.scroll('timeLineGridContainer', 'left', 1000000)"></div>
    </div>
    <div id="timeLineRowLabels"><% renderTimeLineRowLabels %></div>
    <div id="timeLine">
      <div id="timeLineScaleBg"></div>
      <div id="timeLineGridContainer"><% renderTimeLineGrid getTimeLineData, getCategoryQty %></div>
    </div>
    </div>
    <div id="timeLineDetail">
      <div id="timeLineDetailIconsBg"></div>
      <div id="timeLineDetailLeftArrow" onclick="window.scroll('timeLineDetailIcons', 'right', 750)"></div>
      <div id="timeLineDetailRightArrow" onclick="window.scroll('timeLineDetailIcons', 'left', 750)"></div>
      <div id="timeLineDetailIcons"></div>
    </div>
  </div>
  <%
  logActivity "Cartelera", ""
end function

' timeLineDatail =========================================================================================

function fileSizeStr(fileSize)
  if IsNull(fileSize) then
    fileSizeStr = ""
  else
    if filesize > 1024 then
      filesize = filesize / 1024
    		fileSizeStr = round(fileSize, 0) & " KB" 
    	if fileSize > 1024 then
  		  filesize = filesize / 1024
    		fileSizeStr = round(fileSize, 1) & " MB" 
    	end if
    else
      fileSizeStr = fileSize & " Bytes" 
    end if
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

dim imageFileTypes: imageFileTypes = array("jpg", "jpeg", "png", "gif", "bmp")

function isImageFile(fileType)
  isImageFile = false
  dim i
  for i = 0 to uBound(imageFileTypes)
    if imageFileTypes(i) = fileType then
      isImageFile = true
      exit for
    end if
  next
end function

dim recognisedFileTypes: recognisedFileTypes = array("dwg", "doc", "ppt", "xls", "pdf", "zip", "rar")

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
    iconFilename = "front-end/resource/file_icon_" & fileType & ".png"
  else
    iconFilename = "front-end/resource/file_icon_generic.png"
  end if    
end function

function getTimeLineDetailData(timeLineRow, keyValues)
  dim keyVals: keyVals = split(keyValues, ",")
  if dbGetData("SELECT ID, NOMBRE, ARCHIVO_FILENAME, ARCHIVO_FILESIZE, CONTENIDO_TEXTO FROM LINEA_TIEMPO WHERE APROBADO=1 AND " & _
      "ANIO=" & keyVals(1) & " AND MES=" & keyVals(0) & " AND ID_CATEGORIA=" & timeLineRow & " ORDER BY NOMBRE") then
    getTimeLineDetailData = rs.getRows
  else
    getTimeLineDetailData = array()
  end if
end function

function renderTimeLineDetail(timeLineRow, keyValues)
  if not getUsrData then exit function
  dim iconsData: iconsData = getTimeLineDetailData(timeLineRow, keyValues)
  if uBound(iconsData) < 0 then
    %>
    <table cellpadding="0" cellspacing="0" width="100%">
      <tr><td align="center" valign="middle"><%= eval("noDataMsg" & lang) %></td></tr>
    </table>
    <%
  else
    %>
    <table cellpadding="0" cellspacing="0">
      <tr>
      <%
        dim i, rowClass
        for i = 0 to uBound(iconsData, 2)
          if (i mod 2) = 1 then rowClass = "timeLineDetailIconDarkBg" else rowClass = ""
          %>
          <td class="<%= rowClass %>" valign="top" align="center">
            <% renderDownloadIcon "cartelera", iconsData(0, i), iconsData(1, i), iconsData(2, i), iconsData(3, i), iconsData(4, i), 49 %>
            <div class="timeLineDetailIconName"><%= iconsData(1, i) %></div>
          </td>
          <%
        next
      %>
      </tr>
    </table>
    <%
  end if
end function

function sendTimeLineData(recordId, forceDownload)
  if not getUsrData then 
    response.write("Session not valid.")
    exit function
  end if
  dbGetBinaryData "LINEA_TIEMPO", "", "", recordId, forceDownload
  if forceDownload then
    dbGetData("SELECT MES, ANIO, NOMBRE FROM LINEA_TIEMPO WHERE ID = " & recordId)
    logActivity "Descarga", rs("MES") & "-" & rs("ANIO") & ": " & rs("NOMBRE")
    dbReleaseData
  end if
end function

function sendTimeLineThumbnail(recordId, height)
  if request.ServerVariables("SERVER_NAME") = "localhost" then
    sendTimeLineData recordId, false
  else
    if not getUsrData then exit function
    dim s: s = getStringParam("dbFieldBaseName", 30) & "_CONTENTTYPE"
    dim t: t = getStringParam("dbFieldBaseName", 30) & "_BINARYDATA"
    dbGetData("SELECT " & s & "," & t & " FROM LINEA_TIEMPO WHERE ID=" & recordId)
    if not rs.EOF then
      if not isNull(rs(s)) then
        response.contentType = rs(s)
        dim Jpeg: set Jpeg = Server.CreateObject("Persits.Jpeg")
        Jpeg.OpenBinary rs(t).Value
        jpeg.Height = height
        jpeg.Width = jpeg.OriginalWidth * jpeg.Height / jpeg.OriginalHeight
        Jpeg.SendBinary
        set Jpeg = nothing
      end if
    end if
    dbReleaseData
  end if
end function

%>
