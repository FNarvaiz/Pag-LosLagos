<%

function renderClassifieds(categoryId)
  if not getUsrData then exit function
  if not servicesAllowed then exit function
  %>
  <img src="contenidos/costos.jpg" id="costos">
  <%
end function

function renderClassifiedsListing(categoryId)
  if not getUsrData then exit function
  dim searchExpr: searchExpr = ""
  if not isNull(categoryId) then searchExpr = " AND ID_CATEGORIA=" & categoryId
  if dbGetData("SELECT ID, ID_CATEGORIA, dbo.NOMBRE_CATEGORIA_AVISO(ID_CATEGORIA) AS CATEGORIA, " & _
      "dbo.NOMBRE_VECINO(ID_VECINO) AS VECINO, FECHA_ALTA, NOMBRE, COALESCE(DESCRIPCION, '') AS DESCRIPCION, ARCHIVO_CONTENTTYPE " & _
      "FROM AVISOS WHERE VIGENTE=1 AND CADUCIDAD >= GETDATE() " & searchExpr & _
      "ORDER BY FECHA_ALTA DESC") then
    %>
    <div id="classifiedsLeftCol">
    <%
      dim i: i = 1
      do while not rs.EOF
        %>
        <div class="classifiedItem" id="classifiedsItem<%= i %>"><% renderClassifiedItem(i) %></div>
        <%
        i = i + 1
        rs.moveNext
      loop
    %>
    </div>
    <div id="classifiedsRightCol"></div>
    <%
  else
    %>
    <div id="classifiedsEmpty">En este momento no hay ningún aviso.</div>
    <%
  end if
  dbReleaseData
  dim s
  if isNull(categoryId) then
    s = "Todos"
  else
    dbGetData("SELECT NOMBRE FROM CATEGORIAS_AVISOS WHERE ID=" & categoryId)
    s = rs(0)
    dbReleaseData
  end if
  logActivity "Avisos", s
end function

function renderClassifiedItem(itemNumber)
  dim url
  url = serverApp & "?sessionId=" & sessionId & "&content=classifiedImage&recordId=" & rs("ID") & "&dbFieldBaseName=ARCHIVO"
  %>
  <div class="classifiedImage">
    <img class="invisible anchor" onload="this.style.visibility='visible';" src="<%= url %>" title="Clic para ampliar"
      onclick="track('aviso-imagen', '<%= rs("NOMBRE") %>', 'Familia ' + usrName); showImgViewer(this)">
  </div>
  <div class="classifiedCategory">
    <span class="classifiedCategory">Rubro <%= rs("ID_CATEGORIA") %> - <%= uCase(rs("CATEGORIA")) %></span>
    <span class="classifiedDate"><%= day(rs("FECHA_ALTA")) %>-<%= month(rs("FECHA_ALTA")) %>-<%= year(rs("FECHA_ALTA")) %></span> 
  </div>
  <div class="classifiedTitle"><%= rs("NOMBRE") %></div>
  <div class="classifiedDescription"><%= replace(rs("DESCRIPCION"), vbLf, "<br>") %></div>
  <%
end function

function sendClassifiedImage(recordId)
  if not getUsrData then exit function
  dbGetBinaryData "AVISOS", "", "", recordId, false
end function

%>
