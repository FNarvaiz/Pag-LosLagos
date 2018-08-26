<%

function handleAppProcess
  handleAppProcess = true
  select case verb
    case "updatePriceListFromCSVFile": updatePriceListFromCSVFile
    case "updateProductsFromCSVFile": updateProductsFromCSVFile
    case "updateProjectSectionsFromCSVFile": updateProjectSectionsFromCSVFile
    case else handleAppProcess = false
  end select
end function

function updatePriceListFromCSVFile
  Server.ScriptTimeOut = 300
  dim priceListId: priceListId = objUpload("priceListId").value
  dim lines: lines = split(objUpload("csvFile").BinaryAsText, vbCrLf)
  dbConnect
  dbExecute("BEGIN TRANSACTION")
  on error resume next
  dim i, s, prodId, prodCode, prodDescrip, localPrice, exportPrice, currentLocalPrice, currentExportPrice
  dim updatedProdCount: updatedProdCount = 0
  dim newProdCount: newProdCount = 0
  dim updatedPrices
  set updatedPrices = Server.CreateObject("ADODB.Stream")
  initTextStream(updatedPrices)
  dim newPrices
  Set newPrices = Server.CreateObject("ADODB.Stream")
  initTextStream(newPrices)
  for i = 1 to uBound(lines)
    s = split(lines(i), ";")
    if uBound(s) >= 3 then
      prodCode = s(0)
      prodDescrip = s(1)
      localPrice = strToDecimal(s(2))
      if isNull(localPrice) then localPrice = 0
      exportPrice = strToDecimal(s(3))
      if isNull(exportPrice) then exportPrice = 0
      if dbGetData("SELECT ID FROM ARTICULOS WHERE CODIGO=" & sQuotes(prodCode)) then
        prodId = rs("ID")
      else
        prodId = -1
      end if
      dbReleaseData
      if prodId >= 0 then
        if dbGetData("SELECT PRECIO_LOCAL, PRECIO_EXPORTACION FROM LISTAS_PRECIOS_ITEMS " & _
            "WHERE ID_LISTA_PRECIOS=" & priceListId & " AND ID=" & prodId) then
          if isNull(rs("PRECIO_LOCAL")) then currentLocalPrice = 0 else currentLocalPrice = cDbl(rs("PRECIO_LOCAL"))
          if isNull(rs("PRECIO_EXPORTACION")) then currentExportPrice = 0 else currentExportPrice = cDbl(rs("PRECIO_EXPORTACION"))
        else
          currentLocalPrice = -1
          currentExportPrice = -1
        end if
        dbReleaseData
        if currentLocalPrice < 0 or currentExportPrice < 0 then
          dbExecute("INSERT INTO LISTAS_PRECIOS_ITEMS (REC_ID_USUARIO, ID_LISTA_PRECIOS, ID, PRECIO_LOCAL, PRECIO_EXPORTACION) VALUES (" & _
            usrId & ", " & priceListId & ", " & prodId & ", CAST(" & sQuotes(systemToSqlDecimal(localPrice)) & " AS MONEY), " & _
            "CAST(" & sQuotes(systemToSqlDecimal(exportPrice)) & " AS MONEY))")
          newPrices.WriteText("<tr><td>" & prodCode & "</td><td>" & prodDescrip & "</td><td align=right>" & localPrice & _
            "</td><td align=right>" & exportPrice & "</td></tr>" & vbCrLf)
          newProdCount = newProdCount + 1
        elseif currentLocalPrice <> localPrice or currentExportPrice <> exportPrice then
          dbExecute("UPDATE LISTAS_PRECIOS_ITEMS SET REC_ID_USUARIO=" & usrId & ", REC_FECHA=GETDATE(), " & _
            "PRECIO_LOCAL=CAST(" & sQuotes(systemToSqlDecimal(localPrice)) & " AS MONEY), " & _
            "PRECIO_EXPORTACION=CAST(" & sQuotes(systemToSqlDecimal(exportPrice)) & " AS MONEY) " & _
            "WHERE ID_LISTA_PRECIOS=" & priceListId & " AND ID=" & prodId)
          updatedPrices.WriteText("<tr><td>" & prodCode & "</td><td>" & prodDescrip & "</td><td align=right>" & localPrice & _
            "</td><td align=right>" & exportPrice & "</td></tr>" & vbCrLf)
          updatedProdCount = updatedProdCount + 1
        end if
      end if
    end if
    if err.number <> 0 then exit for
  next
  if failed then
    dbExecute("ROLLBACK TRANSACTION")
    %>
    <b>ERROR</b><br>
    <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
      <tr><td><b>El proceso terminó debido a un error. No se modificó ningún dato.<br><br>Mensaje: <%= errorMsg %></b></td></tr>
    </table>
    <%
  else
    dbExecute("COMMIT TRANSACTION")
    %>
    <b>RESULTADOS</b><br>
    <table cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
      <tr><td>Precios que se actualizaron:</td><td><%= updatedProdCount %></td></tr>
      <tr><td>Precios que se agregaron:</td><td><%= newProdCount %></td></tr> 
    </table>
    <br>
    <%
    if updatedProdCount > 0 then
      %>
      <b>PRECIOS QUE SE ACTUALIZARON</b><br>
      <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
        <thead>
          <tr>
            <th align="center"><b>Código</b></th>
            <th align="center"><b>Descripción</b></th>
            <th align="center"><b>Precio Local</b></th>
            <th align="center"><b>Precio Exportación</b></th>
          </tr>
        </thead>
        <tbody>
        <%
        sendTextStream(updatedPrices)
        %>
        </tbody>
      </table>
      <%
    end if
    if newProdCount > 0 then
      %>
      <b>PRECIOS QUE SE AGREGARON</b><br>
      <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
        <thead>
          <tr>
            <th align="center"><b>Código</b></th>
            <th align="center"><b>Descripción</b></th>
            <th align="center"><b>Precio Local</b></th>
            <th align="center"><b>Precio Exportación</b></th>
          </tr>
        </thead>
        <tbody>
        <%
        sendTextStream(newPrices)
        %>
        </tbody>
      <%
    end if
  end if
  newPrices.Close
	set newPrices = Nothing
  updatedPrices.Close
	set updatedPrices = Nothing
  dbDisconnect
end function

' updateProductsFromCSVFile =============================================================

function updateProductsGetTypeId(prodType)
  prodType = escapeQuotes(prodType)
  dim prodTypeId: prodTypeId = null
  if dbGetData("SELECT ID FROM TIPOS_ARTICULOS WHERE LOWER(NOMBRE)=LOWER(" & sQuotes(prodType) & ")") then
    prodTypeId = rs("ID")
  end if
  dbReleaseData
  if isNull(prodTypeId) then
    if dbGetData("SELECT MAX(ID) AS MAX_ID FROM TIPOS_ARTICULOS") then
      prodTypeId = ((rs("MAX_ID") \ 10) + 1) * 10
    end if
    dbReleaseData
    if isNull(prodTypeId) then prodTypeId = 10
    dbExecute("INSERT INTO TIPOS_ARTICULOS (REC_ID_USUARIO, ID, NOMBRE) VALUES (" & usrId & ", " & prodTypeId & ", " & _
      sQuotes(prodType) & ")")
  end if
  updateProductsGetTypeId = prodTypeId
end function

function updateProductsGetLineId(prodLine)
  prodLine = escapeQuotes(prodLine)
  dim prodLineId: prodLineId = null
  if dbGetData("SELECT ID FROM LINEAS_ARTICULOS WHERE LOWER(NOMBRE)=LOWER(" & sQuotes(prodLine) & ")") then
    prodLineId = rs("ID")
  end if
  dbReleaseData
  if isNull(prodLineId) then
    if dbGetData("SELECT MAX(ID) AS MAX_ID FROM LINEAS_ARTICULOS") then
      prodLineId = ((rs("MAX_ID") \ 10) + 1) * 10
    end if
    dbReleaseData
    if isNull(prodLineId) then prodLineId = 10
    dbExecute("INSERT INTO LINEAS_ARTICULOS (REC_ID_USUARIO, ID, NOMBRE) VALUES (" & usrId & ", " & prodLineId & ", " & _
      sQuotes(prodLine) & ")")
  end if
  updateProductsGetLineId = prodLineId
end function

function updateProductsGetFactorySectorId(prodFactorySector)
  prodFactorySector = escapeQuotes(prodFactorySector)
  dim prodFactorySectorId: prodFactorySectorId = null
  if dbGetData("SELECT ID FROM SECTORES_PRODUCCION WHERE LOWER(NOMBRE)=LOWER(" & sQuotes(prodFactorySector) & ")") then
    prodFactorySectorId = rs("ID")
  end if
  dbReleaseData
  if isNull(prodFactorySectorId) then
    if dbGetData("SELECT MAX(ID) AS MAX_ID FROM SECTORES_PRODUCCION") then
      prodFactorySectorId = ((rs("MAX_ID") \ 10) + 1) * 10
    end if
    dbReleaseData
    if isNull(prodFactorySectorId) then prodFactorySectorId = 10
    dbExecute("INSERT INTO SECTORES_PRODUCCION (REC_ID_USUARIO, ID, NOMBRE) VALUES (" & usrId & ", " & prodFactorySectorId & ", " & _
      sQuotes(prodFactorySector) & ")")
  end if
  updateProductsGetFactorySectorId = prodFactorySectorId
end function

function updateProductsFromCSVFile
  Server.ScriptTimeOut = 300
  dim lines: lines = split(objUpload("csvFile").BinaryAsText, vbCrLf)
  on error resume next
  dbConnect
  dbExecute("BEGIN TRANSACTION")
  dbExecute("UPDATE ARTICULOS SET ULTIMA_OPERACION=NULL")
  dim i, s, prodId, prodCode, prodDescrip, prodType, prodLine, prodFactorySector, prodTypeId, prodLineId, prodFactorySectorId
  dim prodOk, lastOperation
  dim prodCount: prodCount = 0
  dim updatedProducts
  Set updatedProducts = Server.CreateObject("ADODB.Stream")
  initTextStream(updatedProducts)
  dim newProducts
  Set newProducts = Server.CreateObject("ADODB.Stream")
  initTextStream(newProducts)
  dim updatedProdCount: updatedProdCount = 0
  dim newProdCount: newProdCount = 0
  dim dataError: dataError = ""
  for i = 1 to uBound(lines)
    s = split(lines(i), ";")
    if uBound(s) >= 4 then
      prodCount = prodCount + 1
      prodCode = trim(s(0))
      prodDescrip = trim(s(1))
      prodType = trim(s(2))
      prodLine = s(3)
      prodFactorySector = s(4)
      if len(prodDescrip) = 0 or len(prodType) = 0 or len(prodLine) = 0 or len(prodFactorySector) = 0 then
        dataError = "Los datos del artículo " & prodCode & " están incompletos."
        exit for
      end if
      prodTypeId = updateProductsGetTypeId(prodType)
      prodLineId = updateProductsGetLineId(prodLine)
      prodFactorySectorId = updateProductsGetFactorySectorId(prodFactorySector)
      prodId = null
      prodOk = true
      lastOperation = null
      if dbGetData("SELECT ID, COALESCE(NOMBRE, '') AS NOMBRE, COALESCE(ID_TIPO_ARTICULO, -1) AS ID_TIPO_ARTICULO, " & _
          "COALESCE(ID_LINEA_ARTICULO, -1) AS ID_LINEA_ARTICULO, COALESCE(ID_SECTOR_PRODUCCION, -1) AS ID_SECTOR_PRODUCCION, " & _
          "ULTIMA_OPERACION FROM ARTICULOS WHERE CODIGO=" & sQuotes(prodCode)) then
        prodId = rs("ID")
        lastOperation = rs("ULTIMA_OPERACION")
        prodOk = (prodDescrip = rs("NOMBRE")) and (prodTypeId = rs("ID_TIPO_ARTICULO")) and (prodLineId = rs("ID_LINEA_ARTICULO")) and _
          (prodFactorySectorId = rs("ID_SECTOR_PRODUCCION"))
      else
        prodOk = false
      end if
      dbReleaseData
      if not isNull(lastOperation) then
        dataError = "El artículo " & prodCode & " aparece más de una vez en el archivo enviado."
        exit for
      end if
      if prodOk then
        dbExecute("UPDATE ARTICULOS SET ULTIMA_OPERACION='N' WHERE ID=" & prodId)
      else
        prodCode = escapeQuotes(prodCode)
        prodDescrip = escapeQuotes(prodDescrip)
        if isNull(prodId) then
          dbExecute("INSERT INTO ARTICULOS (REC_ID_USUARIO, CODIGO, NOMBRE, ID_TIPO_ARTICULO, ID_LINEA_ARTICULO, " & _
            "ID_SECTOR_PRODUCCION, ULTIMA_OPERACION) VALUES (" & usrId & ", " & sQuotes(prodCode) & ", " & _
            sQuotes(prodDescrip) & ", " & prodTypeId & ", " & prodLineId & ", " & prodFactorySectorId & ", 'I')")
          newProducts.writeText("<tr><td>" & prodCode & "</td><td>" & prodDescrip & "</td><td align=center>" & prodType & _
            "</td><td align=center>" & prodLine & "</td><td align=center>" & prodFactorySector & "</td></tr>" & vbCrLf)
          newProdCount = newProdCount + 1
        else
          dbExecute("UPDATE ARTICULOS SET REC_ID_USUARIO=" & usrId & ", REC_FECHA=GETDATE(), " & _
            "NOMBRE=" & sQuotes(prodDescrip) & ", ID_TIPO_ARTICULO=" & prodTypeId & ", ID_LINEA_ARTICULO=" & prodLineId & _
            ", ID_SECTOR_PRODUCCION=" & prodFactorySectorId & ", ULTIMA_OPERACION='U' WHERE ID=" & prodId)
          updatedProducts.writeText("<tr><td>" & prodCode & "</td><td>" & prodDescrip & "</td><td align=center>" & prodType & _
            "</td><td align=center>" & prodLine & "</td><td align=center>" & prodFactorySector & "</td></tr>" & vbCrLf)
          updatedProdCount = updatedProdCount + 1
        end if
      end if
    end if
    if err.number <> 0 then exit for
  next
  if len(dataError) > 0 then
    dbExecute("ROLLBACK TRANSACTION")
    %>
    <b>PROBLEMA</b>
    <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
      <tr><td><b><%= dataError %><br>No se modificó ningún dato.</b></td></tr>
    </table>
    <%
  elseif failed then
    dbExecute("ROLLBACK TRANSACTION")
    %>
    <b>ERROR</b><br>
    <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
      <tr><td><b>El proceso terminó debido a un error.<br>No se modificó ningún dato.<br><br>Mensaje: <%= errorMsg %></b></td></tr>
    </table>
    <%
  else
    dbExecute("COMMIT TRANSACTION")
    %>
    <b>RESULTADOS</b><br>
    <table cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
      <tr><td>Artículos en el archivo recibido:</td><td><%= prodCount %></td></tr>
      <tr><td>Artículos existentes que se actualizaron:</td><td><%= updatedProdCount %></td></tr>
      <tr><td>Artículos nuevos que se agregaron:</td><td><%= newProdCount %></td></tr>
    </table>
    <br>
    <%
    if updatedProdCount > 0 then
      %>
      <b>ARTICULOS QUE SE ACTUALIZARON</b><br>
      <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
        <thead>
          <tr>
            <th align="center"><b>Código</b></th>
            <th align="center"><b>Descripción</b></th>
            <th align="center"><b>Clase</b></th>
            <th align="center"><b>Línea</b></th>
            <th align="center"><b>Sector</b></th>
          </tr>
        </thead>
        <tbody>
          <%
          sendTextStream(updatedProducts)
          %>
        </tbody>
      </table>
      <%
    end if
    if newProdCount > 0 then
      %>
      <br><b>ARTICULOS NUEVOS QUE SE AGREGARON</b>
      <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
        <thead>
          <tr>
            <th align="center"><b>Código</b></th>
            <th align="center"><b>Descripción</b></th>
            <th align="center"><b>Clase</b></th>
            <th align="center"><b>Línea</b></th>
            <th align="center"><b>Sector</b></th>
          </tr>
        </thead>
        <tbody>
          <%
          sendTextStream(newProducts)
          %>
        </tbody>
      </table>
      <%
    end if
  end if
  newProducts.Close
	set newProducts = Nothing
  updatedProducts.Close
	set updatedProducts = Nothing
  dbDisconnect
end function

' updateProjectSectionsFromCSVFile =============================================================

function updateProjectSectionsGetSectionId(projectId, variantId, versionId, sectionName)
  dim sectionId: sectionId = null
  dim searchExpr: searchExpr = "ID_PROYECTO=" & projectId & " AND ID_VARIANTE=" & variantId & " AND ID_VERSION=" & versionId
  if dbGetData("SELECT ID FROM PROYECTOS_SECCIONES WHERE " & searchExpr & " AND LOWER(NOMBRE)=LOWER(" & sQuotes(sectionName) & ")") then
    sectionId = rs("ID")
  end if
  dbReleaseData
  if isNull(sectionId) then
    if dbGetData("SELECT MAX(ID) AS MAX_ID FROM PROYECTOS_SECCIONES WHERE " & searchExpr) then
      sectionId = ((rs("MAX_ID") \ 10) + 1) * 10
    end if
    if isNull(sectionId) then sectionId = 10
    dbReleaseData
    dbExecute("INSERT INTO PROYECTOS_SECCIONES (REC_ID_USUARIO, ID_PROYECTO, ID_VARIANTE, ID_VERSION, ID, NOMBRE) VALUES (" & _
      usrId & ", " & projectId & ", " & variantId & ", " & versionId & ", " & sectionId & ", " & sQuotes(sectionName) & ")")
  end if
  updateProjectSectionsGetSectionId = sectionId
end function

function updateProjectSectionsFromCSVFile
  Server.ScriptTimeOut = 300
  dim projectId: projectId = objUpload("projectId").value
  dim variantId: variantId = objUpload("variantId").value
  dim versionId: versionId = objUpload("versionId").value
  dim lines: lines = split(objUpload("csvFile").BinaryAsText, vbCrLf)
  dim i, s
  dim sectionNameColumn: sectionNameColumn = -1
  dim prodCodeColumn: prodCodeColumn = -1
  s = split(lines(0), ";")
  for i = 0 to uBound(s)
    select case lcase(s(i)):
      case "block name": prodCodeColumn = i
      case "layer": sectionNameColumn = i
    end select
  next
  if sectionNameColumn < 0 or prodCodeColumn < 0 then
    response.write("El formato del archivo es incorrecto.<br><br>" & vbCrLf & join(lines, "<br>" & vbCrLf))
    exit function
  end if

  dbConnect
  dbExecute("BEGIN TRANSACTION")
  'on error resume next
  dim addedProducts
  Set addedProducts = Server.CreateObject("ADODB.Stream")
  initTextStream(addedProducts)
  dim unknownProducts
  Set unknownProducts = Server.CreateObject("ADODB.Stream")
  initTextStream(unknownProducts)
  dim unvaluedProducts
  Set unvaluedProducts = Server.CreateObject("ADODB.Stream")
  initTextStream(unvaluedProducts)
  dim addedProdCount: addedProdCount = 0
  dim unknownProductCount: unknownProductCount = 0
  dim unvaluedProductCount: unvaluedProductCount = 0
  dim sectionName, sectionId, prodCode, prodId, prodDescrip, prodType, prodLine, prodCount, prodTypeId, prodLineId, prodPrice
  dim searchExpr: searchExpr = "ID_PROYECTO=" & projectId & " AND ID_VARIANTE=" & variantId & " AND ID_VERSION=" & versionId
  dim itemId
  for i = 1 to uBound(lines)
    s = split(lines(i), ";")
    if not isArray(s) then exit for
    if (uBound(s) < sectionNameColumn) or (uBound(s) < prodCodeColumn) then exit for
    sectionName = trim(s(sectionNameColumn))
    prodCode = trim(s(prodCodeColumn))
    if len(sectionName) = 0 or len(prodCode) = 0 then exit for
    if dbGetData("SELECT ID, NOMBRE, dbo.NOMBRE_LINEA_ARTICULO(ID_LINEA_ARTICULO) AS LINEA, " & _
        "dbo.NOMBRE_TIPO_ARTICULO(ID_TIPO_ARTICULO) AS TIPO, dbo.PRECIO_ARTICULO_PROYECTO(" & projectId & ", " & variantId & ") AS PRECIO " & _
        "FROM ARTICULOS WHERE CODIGO=" & sQuotes(prodCode)) then
      prodId = rs("ID")
      prodDescrip = rs("NOMBRE")
      prodType = rs("TIPO")
      prodLine = rs("LINEA")
      prodPrice = rs("PRECIO")
    else
      prodId = -1
    end if
    dbReleaseData
    if prodId >= 0 then
      if isNull(prodPrice) then
        'if inStr(unvaluedProducts, prodCode) = 0 then
          unvaluedProducts.writeText("<tr><td>" & sectionName & "</td><td>" & prodCode & "</td></tr>" & vbCrLf)
          unvaluedProductCount = unvaluedProductCount + 1
        'end if
      else
        sectionId = updateProjectSectionsGetSectionId(projectId, variantId, versionId, sectionName)
        itemId = -1
        if dbGetData("SELECT ID FROM PROYECTOS_SECCIONES_ITEMS WHERE " & searchExpr & " AND ID_SECCION=" & sectionId & _
            " AND ID_ARTICULO=" & prodId) then
          itemId = rs("ID")
        end if
        dbReleaseData
        if cInt(itemId) >= 0 then
          dbExecute("UPDATE PROYECTOS_SECCIONES_ITEMS SET REC_ID_USUARIO=" & usrId & ", REC_FECHA=GETDATE(), " & _
            "CANTIDAD=CANTIDAD + 1 WHERE ID=" & itemId)
        else
          dbExecute("INSERT INTO PROYECTOS_SECCIONES_ITEMS " & _
            "(REC_ID_USUARIO, ID_PROYECTO, ID_VARIANTE, ID_VERSION, ID_SECCION, CANTIDAD, ID_ARTICULO) VALUES (" & usrId & ", " & _
            projectId & ", " & variantId & ", " & versionId & ", " & sectionId & ", 1, " & prodId & ")")
        end if
        addedProducts.writeText("<tr><td>" & sectionName & "</td><td>" & prodCode & "</td><td>" & prodDescrip & _
          "</td><td align=center>" & prodType & "</td><td align=center>" & prodType & "</td></tr>" & vbCrLf)
        addedProdCount = addedProdCount + 1
      end if
    else'if inStr(unknownProducts, prodCode) = 0 then
      unknownProducts.writeText("<tr><td>" & sectionName & "</td><td>" & prodCode & "</td></tr>" & vbCrLf)
      unknownProductCount = unknownProductCount + 1
    end if
    if err.number <> 0 then exit for
  next
  if failed then
    dbExecute("ROLLBACK TRANSACTION")
    %>
    <b>ERROR</b><br>
    <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
      <tr><td><b>El proceso terminó debido a un error. No se incorporaron datos.<br><br>Mensaje: <%= errorMsg %></b></td></tr>
    </table>
    <%
  else
    dbExecute("COMMIT TRANSACTION")
    %>
    <b>RESULTADOS</b><br>
    <table cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
      <tr><td>Artículos que se incorporaron al proyecto:</td><td><%= addedProdCount %></td></tr>
      <tr><td>Artículos desconocidos:</td><td><%= unknownProductCount %></td></tr>
      <tr><td>Artículos sin precio (no se incorporaron):</td><td><%= unvaluedProductCount %></td></tr> 
    </table>
    <%
    if addedProdCount > 0 then
      %>
      <b>ARTICULOS QUE SE INCORPORARON AL PROYECTO</b><br>
      <table width="100%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
        <thead>
          <tr>
            <th align="center"><b>Sección</b></th>
            <th align="center"><b>Código</b></th>
            <th align="center"><b>Descripción</b></th>
            <th align="center"><b>Clase</b></th>
            <th align="center"><b>Línea</b></th>
          </tr>
        </thead>
        <tbody>
          <%
          sendTextStream(addedProducts)
          %>
        </tbody>
      </table>
      <%
    end if
    if unknownProductCount > 0 then
      %>
      <b>ARTICULOS DESCONOCIDOS</b><br>
      <table width="50%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
        <thead>
          <tr>
            <th align="center"><b>Sección</b></th>
            <th align="center"><b>Código</b></th>
          </tr>
        </thead>
        <tbody>
          <%
          sendTextStream(unknownProducts)
          %>
        </tbody>
      </table>
      <%
    end if
    if unvaluedProductCount > 0 then
      %>
      <b>ARTICULOS SIN PRECIO</b><br>
      <table width="50%" cellpadding="0" cellspacing="4" style="font-family: arial; font-size: 11px">
        <thead>
          <tr>
            <th align="center"><b>Sección</b></th>
            <th align="center"><b>Código</b></th>
          </tr>
        </thead>
        <tbody>
          <%
          sendTextStream(unvaluedProducts)
          %>
        </tbody>
      </table>
      <%
    end if
  end if
  addedProducts.close
  set addedProducts = nothing
  unknownProducts.close
  set unknownProducts = nothing
  unvaluedProducts.close
  set unvaluedProducts = nothing
  dbDisconnect
end function

%>
