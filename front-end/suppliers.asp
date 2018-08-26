<%

function renderSuppliers(categoryId)
  if not getUsrData then exit function
  if not servicesAllowed then exit function
  %>
  <div id="dynPanelBg"></div>
  <div id="suppliersTitle">
    DIRECTORIO DE PROVEEDORES
    <p>Aquí podrás consultar listados de los proveedores que colaboran con nuestro barrio.</p>
    <p>Para cada uno de ellos podrás expresar tu valoración personal y consultar cómo son valorados dentro del barrio.</p>
    <p>En caso de querer incluir un proveedor envianos un e-mail a 
      <a href="mailto:Vecinos de los Lagos <sistema@vecinosloslagos.com.ar>">sistema@vecinosloslagos.com.ar</a></p> 
  </div>
  <div id="suppliersCategories">
    RUBRO: 
    <select id="suppliersCategorySelector" onchange="suppliersCategorySelected(this)">
      <%
      if isNull(categoryId) then
        %>
        <option selected="selected" value="">TODOS</option>
        <%
      else
        %>
        <option value="">TODOS</option>
        <%
      end if
      dbGetData("SELECT C.ID, C.NOMBRE " & _
        "FROM CATEGORIAS_PROVEEDORES C WHERE EXISTS (SELECT * FROM PROVEEDORES P WHERE P.ID_CATEGORIA=C.ID AND VIGENTE=1) ORDER BY NOMBRE")
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
  <div id="suppliersPanel">
    <% renderSuppliersListing(categoryId) %>
  </div>
  <%
end function

function renderSuppliersListing(categoryId)
  if not getUsrData then exit function
  %>
  <table cellpadding="0" cellspacing="0" width= "100%">
    <%
    dim searchExpr: searchExpr = ""
    if not isNull(categoryId) then searchExpr = " AND ID_CATEGORIA=" & categoryId
    if dbGetData("SELECT ID, dbo.NOMBRE_CATEGORIA_PROVEEDOR(ID_CATEGORIA) AS CATEGORIA, NOMBRE, " & _
        "COALESCE(DOMICILIO, '') AS DOMICILIO, COALESCE(TELEFONOS, '') AS TELEFONOS, COALESCE(EMAIL, '') AS EMAIL, " & _
        "dbo.VALORACION_PROVEEDOR(ID) AS VALORACION, dbo.NOMBRE_VALORACION(dbo.VALORACION_PROVEEDOR(ID)) AS NOMBRE_VALORACION, " & _ 
        "dbo.CANTIDAD_VALORACIONES_PROVEEDOR(ID) AS CANTIDAD_VALORACIONES, " & _
        "dbo.NOMBRE_VALORACION(dbo.VALORACION_PROVEEDOR_VECINO(ID, " & usrId & ")) AS VALORACION_VECINO " & _
        "FROM PROVEEDORES WHERE VIGENTE=1 " & searchExpr & _
        "ORDER BY CATEGORIA, NOMBRE") then
      dim i: i = 0
      dim j: j = 1
      do while not rs.EOF
        if i = 0 then
          %>
          <tr>
          <%
        end if
        %>
        <td width="48%" valign="top" class="supplierItem">
          <div id="suppliersItem<%= j %>"><% renderSuppliersItem(j) %></div>
        </td>
        <%
        if i = 0 then
          %>
          <td width="20">&nbsp;</td>
          <%
          i = 1
        else
          %>
          </tr>
          <tr><td>&nbsp;</td></tr>
          <%
          i = 0
        end if
        j = j + 1
        rs.moveNext
      loop
    else
      %>
      <tr><td align="center" height="25">(No hay proveedores disponibles en este momento.)</td></tr>
      <%
    end if
    dbReleaseData
    %>
  </table>
  <%
  dim s
  if isNull(categoryId) then
    s = "Todos"
  else
    dbGetData("SELECT NOMBRE FROM CATEGORIAS_PROVEEDORES WHERE ID=" & categoryId)
    s = rs(0)
    dbReleaseData
  end if
  logActivity "Proveedores", s
end function

function renderSuppliersItem(itemNumber)
  %>
  <div class="supplierResult">
  <%
    if isNull(rs("VALORACION")) then
      %>
      <div class="voteQty">Sin votos</div>
      <%
    else
      %>
      <div class="voteResultBg"><img src="front-end/resource/votos_bg.png"></div>
      <div class="voteResultfg">
        <img src="front-end/resource/votos_fg.png"
          style="position: absolute; clip:rect(0px,<%= cInt(round(64 * cDbl(rs("VALORACION")) / 100))  %>px,10px,0px)">
      </div>
      <div class="voteQty">
        <%= rs("NOMBRE_VALORACION") %><br>Votos: <%= rs("CANTIDAD_VALORACIONES") %>
      </div>
      <%
    end if
  %>
  </div>
  <div class="supplierCategory"><%= rs("CATEGORIA") %></div>
  <div class="supplierTitle"><%= rs("NOMBRE") %></div>
  <div class="supplierDescription">
    <%
    dim s: s = rs("TELEFONOS")
    if not isNull(rs("EMAIL")) then
      if len(s) > 0 then s = s & " - "
      s = s & "<a href=" & dQuotes("mailto:" & rs("EMAIL")) & ">" & rs("EMAIL") & "</a>"
    end if
    %>
    <%= s %><br>
    <%= replace(rs("DOMICILIO"), vbLf, ", ") %>
  </div>
  <div class="supplierVote">
  <%
    s = rs("VALORACION_VECINO")
    dim t: t = ""
    dim confirmationQuestion: confirmationQuestion = ""
    if isNull(s) then
      s = ""
    else
      confirmationQuestion = "Para " & ucase(rs("NOMBRE")) & _
        ", ¿querés cambiar tu valoración de " & ucase(s) & " a "
    end if
    if s = "Malo" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto1.png" title="Malo"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Malo") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto1.png" title="Malo"
        onclick="saveSupplierVote('<%= t %>', <%= rs("ID") %>, 0, <%= itemNumber %>)"></div>
      <%
    end if
    if s = "Regular" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto2.png" title="Regular"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Regular") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto2.png" title="Regular"
        onclick="saveSupplierVote('<%= t %>', <%= rs("ID") %>, 33, <%= itemNumber %>)"></div>
      <%
    end if
    if s = "Bueno" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto3.png" title="Bueno"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Bueno") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto3.png" title="Bueno"
        onclick="saveSupplierVote('<%= t %>', <%= rs("ID") %>, 67, <%= itemNumber %>)"></div>
      <%
    end if
    if s = "Excelente" then
      %>
      <div class="voteMarked"><img src="front-end/resource/voto4.png" title="Excelente"></div>
      <%
    else
      if len(s) > 0 then t = confirmationQuestion & ucase("Excelente") & "?"
      %>
      <div class="vote"><img src="front-end/resource/voto4.png" title="Excelente"
        onclick="saveSupplierVote('<%= t %>', <%= rs("ID") %>, 100, <%= itemNumber %>)"></div>
      <%
    end if
    if isNull(rs("VALORACION_VECINO")) then
      %>
      <div class="vote">Aún no has expresado tu valoración.</div>
      <%
    else
      %>
      <div class="vote">Tu valoración: <%= rs("VALORACION_VECINO") %>.</div>
      <%
    end if
  %>
  </div>
  <%
end function

function saveSupplierVote(supplierId, voteValue, itemNumber)
  if not getUsrData then exit function
  dim voteId: voteId = null
  if dbGetData("SELECT ID FROM PROVEEDORES_VALORACIONES WHERE ID_PROVEEDOR=" & supplierId & " AND ID_VECINO=" & usrId) then
    voteId = rs(0)
  end if
  dbReleaseData
  if isNull(voteId) then
    dbExecute("INSERT INTO PROVEEDORES_VALORACIONES (ID_PROVEEDOR, ID_VECINO, VALORACION) VALUES (" & supplierId & ", " & usrId & ", " & voteValue & ")")
  else
    dbExecute("UPDATE PROVEEDORES_VALORACIONES SET VALORACION=" & voteValue & ", FECHA=GETDATE() WHERE ID_PROVEEDOR=" & supplierId & " AND ID_VECINO=" & usrId)
  end if
  dbGetData("SELECT ID, dbo.NOMBRE_CATEGORIA_PROVEEDOR(ID_CATEGORIA) AS CATEGORIA, NOMBRE, " & _
    "COALESCE(DOMICILIO, '') AS DOMICILIO, COALESCE(TELEFONOS, '') AS TELEFONOS, COALESCE(EMAIL, '') AS EMAIL, " & _
    "dbo.VALORACION_PROVEEDOR(ID) AS VALORACION, dbo.NOMBRE_VALORACION(dbo.VALORACION_PROVEEDOR(ID)) AS NOMBRE_VALORACION, " & _ 
    "dbo.CANTIDAD_VALORACIONES_PROVEEDOR(ID) AS CANTIDAD_VALORACIONES, " & _
    "dbo.NOMBRE_VALORACION(dbo.VALORACION_PROVEEDOR_VECINO(ID, " & usrId & ")) AS VALORACION_VECINO " & _
    "FROM PROVEEDORES WHERE ID=" & supplierId)
  renderSuppliersItem(itemNumber)
  logActivity "Voto proveedor", "Proveedor = " & rs("NOMBRE") & ", Valoración = " & rs("VALORACION_VECINO")
  dbReleaseData
end function

%>
