<%

function renderDownloads(categoryId)
  if not getUsrData then exit function
  %>
  <div id="dynPanelBg"></div>
  <div id="downloadsListingHeading">
    <div id="downloadControls">
      MOSTRAR: <select id="downloadsCategorySelector" onchange="downloadCategorySelected(this)">
      <%
        if isNull(categoryId) then
          %>
          <option selected="selected" value="">TODO</option>
          <%
        else
          %>
          <option value="">TODAS</option>
          <%
        end if
        dbGetData("SELECT C.ID, C.NOMBRE " & _
          "FROM CATEGORIAS_LINEA_TIEMPO C " & _
          "WHERE C.VIGENTE = 1 AND EXISTS (SELECT * FROM LINEA_TIEMPO L WHERE L.ID_CATEGORIA=C.ID AND L.APROBADO=1) ORDER BY ID")
        do while not rs.EOF
          if rs("ID") = categoryId then
            %>
            <option selected="selected" value="<%= rs("ID") %>"><%= uCase(rs("NOMBRE")) %></option>
            <%
          else
            %>
            <option value="<%= rs("ID") %>"><%= uCase(rs("NOMBRE")) %></option>
            <%
          end if
          rs.moveNext
        loop
        dbReleaseData
      %>
      </select>
    </div>
    <table cellpadding="3" cellspacing="0" width= "100%">
      <tr>
        <td width="80" align="center">VIGENCIA</td>
        <% if isNull(categoryId) then %>
          <td width="150" align="center">CATEGORIA</td>
        <% end if %>
        <td align="left">DESCRIPCION</td>
      </tr>
    </table>
  </div>
  <div id="downloadsListing">
  </div>
	<%
end function

function renderDownloadListing(categoryId)
  if not getUsrData then exit function
  %>
  <table cellpadding="3" cellspacing="0" width= "100%">
    <%
    dim s: s = ""
    if not isNull(categoryId) then s = " AND C.ID=" & categoryId
    if dbGetData("SELECT L.ID, L.NOMBRE, L.PERIODO, L.CONTENIDO_TEXTO, L.ARCHIVO_FILENAME, L.ARCHIVO_FILESIZE, C.NOMBRE AS CATEGORIA " & _
        "FROM LINEA_TIEMPO L JOIN CATEGORIAS_LINEA_TIEMPO C ON C.ID = L.ID_CATEGORIA " & _
        "WHERE C.VIGENTE = 1 AND L.APROBADO=1" & s & " ORDER BY L.ANIO DESC, L.MES DESC, L.ID DESC") then
      dim descrip, rowClass, oddRow
      do while not rs.EOF
        descrip = rs("NOMBRE")
        if isNull(categoryId) then descrip = "<b><i>" & rs("CATEGORIA") & "</i></b>: " & descrip
        if oddRow then rowClass = "oddRow" else rowClass = ""
        %>
        <tr class="<%= rowClass %>">
          <td width="80" align="center"><%= rs("PERIODO") %></td>
          <% if isNull(categoryId) then %>
            <td width="150" align="center"><%= rs("CATEGORIA") %></td>
            <td width="500"><%= rs("NOMBRE") %></td>
          <% else %>
            <td width="650"><%= rs("NOMBRE") %></td>
          <% end if %>
          <td align="center">
            <% renderDownloadIcon "Historico", rs("ID"), rs("NOMBRE"), rs("ARCHIVO_FILENAME"), rs("ARCHIVO_FILESIZE"), rs("CONTENIDO_TEXTO"), 30 %>
          </td>
        </tr>
        <%
        oddRow = not oddRow
        rs.MoveNext
      loop
    else
      %>
      <tr><td align="center" height="25">(No hay ítems disponibles en esta categoría.)</td></tr>
      <%
    end if
    dbReleaseData
    %>
  </table>
  <%
  if isNull(categoryId) then
    s = "Todo"
  else
    dbGetData("SELECT NOMBRE FROM CATEGORIAS_LINEA_TIEMPO WHERE ID=" & categoryId)
    s = rs(0)
    dbReleaseData
  end if
  logActivity "Histórico", s
end function

function renderDownloadIcon(sectionName, id, descrip, filename, filesize, textContent, iconHeight)
  if not isNull(textContent) then
    %>
    <img class="downloadIcon invisible" onload="this.style.visibility='visible';" src="<%= iconFilename("txt") %>" height="<%= iconHeight %>"
      onclick="track('<%= sectionName %>-texto', '<%= descrip %>', 'Familia ' + usrName); showDataViewer(<%= id %>)"
      title="Información adicional">
    <%
    exit function
  end if

  if isNull(filename) then
    response.write("&nbsp;")
    exit function
  end if

  dim ext: ext = getFileExt(filename)
  if isImageFile(ext) then
    dim url
    url = serverApp & "?sessionId=" & sessionId & "&content=timeLineThumbnail&recordId=" & id & "&dbFieldBaseName=ARCHIVO&height=49"
    %>
    <img class="downloadIcon invisible" onload="this.style.visibility='visible';" src="<%= url %>" height="<%= iconHeight %>"
      onclick="track('<%= sectionName %>-imagen', '<%= descrip %>', 'Familia ' + usrName); showImgViewer(this)"
      title="Imagen">
    <%
  else
    %>
    <img class="downloadIcon invisible" onload="this.style.visibility='visible';" src="<%= iconFilename(ext) %>" height="<%= iconHeight %>"
      onclick="track('<%= sectionName %>-descarga', '<%= descrip %>', 'Familia ' + usrName); document.downloadForm<%= id %>.submit()"
      title="<%= filename %> (<%= fileSizeStr(filesize) %>)">
    <div style="display:none">
      <form name="downloadForm<%= id %>" target="_self" method="get" action="<%= serverApp %>">
        <input type="text" name="content" value="timeLineData">
        <input type="text" name="sessionId" value="<%= sessionId %>">
        <input type="text" name="recordId" value="<%= id %>">
        <input type="text" name="dbFieldBaseName" value="ARCHIVO">
        <input type="text" name="forceDownload" value="1">
      </form>
    </div>
    <%
  end if
end function

function getTextContent(contentId)
  if not getUsrData then exit function
  dbGetData("SELECT CONTENIDO_TEXTO FROM LINEA_TIEMPO WHERE ID = " & contentId)
  response.write(replace(rs(0), vbLf, "<br>"))
  dbReleaseData
end function

%>
