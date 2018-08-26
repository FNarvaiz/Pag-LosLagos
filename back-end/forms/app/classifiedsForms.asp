<%

function formAdminClassifieds
  usrAccessLevel = usrAccessAdminClassifieds
  formContainerCssClass = "ClassifiedsFormContainer"
  formContainerTitleCssClass = "ClassifiedsFormContainerTitle"
  formTitle = "Avisos clasificados"
  forms = array("formClassifieds")
  childFormIds = "formClassifieds"
end function

dim formClassifiedsQueryTypeNames
formClassifiedsQueryTypeNames = array("Vigentes", "No vigentes", "Todos")

dim formClassifiedsQueryTypeSearchConditions
formClassifiedsQueryTypeSearchConditions = array("VIGENTE=1 AND CADUCIDAD >= GETDATE()", "VIGENTE=0 OR CADUCIDAD < GETDATE()", "")

dim formClassifiedsQuerySearchFieldNames
formClassifiedsQuerySearchFieldNames = array("Por Nombre")

dim formClassifiedsQuerySearchFields
formClassifiedsQuerySearchFields = array("NOMBRE")

dim formClassifiedsQueryLimitNames
formClassifiedsQueryLimitNames = array("hasta 100 ítems", "hasta 1000 ítems", "Todo")

dim formClassifiedsQueryLimitClauses
formClassifiedsQueryLimitClauses = array("TOP 100", "TOP 1000", "")

function formClassifieds
  usrAccessLevel = usrAccessAdminClassifieds
  formTitle = ""
  formTable = "AVISOS"
  
  formGridViewRowCount = 25
  formGridColumns = array("FECHA_ALTA", "dbo.NOMBRE_CATEGORIA_AVISO(ID_CATEGORIA)", "NOMBRE", "ARCHIVO", "dbo.UNIDAD_VECINO(ID_VECINO)", _
    "dbo.NOMBRE_VECINO(ID_VECINO)")
  formGridColumnTypes = array(formGridColumnDate, formGridColumnGeneralCenter, formGridColumnGeneral, formGridColumnImage, _
    formGridColumnGeneralCenter, formGridColumnGeneral)
  formGridColumnLabels = array("Alta", "Rubro", "Título", "Foto", "U.", "Familia")
  formGridColumnWidths = array(60, 100, 200, 70, 40, 130)
  gridViewOrderBy = "1 DESC"

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 200
  recordViewIdFieldIsIdentity = true

  recordViewSeparators = array(null, null, null, null, null, null, null, "Descripción", "Foto")
  recordViewFields = array("ID", "ID_CATEGORIA", "ID_VECINO", "NOMBRE", "FECHA_ALTA", "VIGENTE", "CADUCIDAD", "DESCRIPCION", "ARCHIVO_CONTENTTYPE")
  recordViewDBFields = array(true, true, true, true, true, true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, false, true, false, false, false, true)
  recordViewFieldLabels = array("Código", "Categoría", "Vecino", "Título", "Fecha alta", "Vigente", "Caducidad", "Descripción", "Archivo")
  recordViewFieldDefaults = array(null, null, null, null, date, true, dateAdd("m", 1, date), null, null)
  recordViewNullableFields = array(true, false, false, false, true, false, false, false, true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", _
    "renderRecordViewLookupField(" & dQuotes("CATEGORIAS_AVISOS,ID,VIGENTE=1 OR ID=dbo.ID_CATEGORIA_AVISO(" & recordId & ")") & ")", _
    "renderRecordViewLookupField(" & dQuotes("VECINOS,ID") & ")", _
    "renderRecordViewLiteralField", "renderRecordViewDateField", "renderRecordViewBooleanField", "renderRecordViewDateField", _
    "renderRecordViewNotesField(10)", "renderRecordViewFileUploadField")

  dim i
  dbGetData("SELECT COUNT(*) AS QTY FROM CATEGORIAS_AVISOS")
  i = rs("QTY")
  dbReleaseData
  if i > 0 then
    redim preserve formClassifiedsQueryTypeNames(i + 2)
    redim preserve formClassifiedsQueryTypeSearchConditions(i + 2)
    dbGetData("SELECT * FROM CATEGORIAS_AVISOS ORDER BY ID")
    i = 3
    do while not rs.EOF
      formClassifiedsQueryTypeNames(i) = "Rubro: " & rs("NOMBRE")
      formClassifiedsQueryTypeSearchConditions(i) = "VIGENTE=1 AND ID_CATEGORIA=" & rs("ID")
      rs.moveNext
      i = i + 1
    loop
    dbReleaseData
  end if
end function

%>
