<%

function formAdminSuppliers
  usrAccessLevel = usrAccessAdminSuppliers
  formContainerCssClass = "suppliersFormContainer"
  formContainerTitleCssClass = "suppliersFormContainerTitle"
  formTitle = "Proveedores"
  forms = array("formSuppliers", "formSuppliersReviews")
  childFormIds = "formSuppliers,formSuppliersReviews"
end function

dim formSuppliersQueryTypeNames
formSuppliersQueryTypeNames = array("Vigentes", "No vigentes", "Todos")

dim formSuppliersQueryTypeSearchConditions
formSuppliersQueryTypeSearchConditions = array("VIGENTE=1", "VIGENTE=0", "")

dim formSuppliersQuerySearchFieldNames
formSuppliersQuerySearchFieldNames = array("Por Nombre")

dim formSuppliersQuerySearchFields
formSuppliersQuerySearchFields = array("NOMBRE")

function formSuppliersBeforeUpdate
  formSuppliersBeforeUpdate =  true
  if fieldChanged(getFieldIndex("VIGENTE")) then
    if cInt(fieldNewValues(getFieldIndex("VIGENTE"))) = 0 then
      setFieldNewValue "FECHA_BAJA", date
    else
      setFieldNewValue "FECHA_BAJA", null
    end if
  end if
end function

function formSuppliers
  usrAccessLevel = usrAccessAdminSuppliers
  formTitle = ""
  formTable = "PROVEEDORES"
  childFormIds = "formSuppliersReviews"
  
  formBeforeUpdateFunc = "formSuppliersBeforeUpdate"
  
  formGridViewRowCount = 27
  formGridColumns = array("NOMBRE", "TELEFONOS", "dbo.NOMBRE_VALORACION(dbo.VALORACION_PROVEEDOR(ID))", "dbo.CANTIDAD_VALORACIONES_PROVEEDOR(ID)")
  formGridColumnTypes = array(formGridColumnGeneral, formGridColumnGeneral, formGridColumnGeneralCenter, formGridColumnGeneralCenter)
  formGridColumnLabels = array("Nombre", "Teléfonos", "Valor.", "Votos")
  formGridColumnWidths = array(160, 130, 40, 40)
  defaultQueryLimit = -1

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 200
  recordViewIdFieldIsIdentity = true

  recordViewSeparators = array(null, null, null, null, null, null, _ 
    null, null, null, "Notas")
  recordViewFields = array("ID", "ID_CATEGORIA", "NOMBRE", "DOMICILIO", "TELEFONOS", "EMAIL", _
    "FECHA_ALTA", "VIGENTE", "FECHA_BAJA", "NOTAS")
  recordViewDBFields = array(true, true, true, true, true, true, _
    true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, false, false, false, _
    true, false, true, false)
  recordViewFieldLabels = array("Código", "Rubro", "Nombre", "Domicilio", "Teléfonos", "e-mail", _
    "Fecha alta", "Vigente", "Fecha baja", "")
  recordViewFieldDefaults = array(null, null, null, null, null, null, _
    null, true, null, null)
  recordViewNullableFields = array(true, false, false, true, true, true, _
    true, false, true, true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", _
    "renderRecordViewLookupField(" & dQuotes("CATEGORIAS_PROVEEDORES,ID,VIGENTE=1 OR ID=dbo.ID_CATEGORIA_PROVEEDOR(" & recordId & ")") & ")", _
    "renderRecordViewLiteralField", "renderRecordViewTextField(3)", "renderRecordViewLiteralField", "renderRecordViewLiteralField", _
    "renderRecordViewDateField", "renderRecordViewBooleanField", "renderRecordViewDateField", "renderRecordViewNotesField(8)")

  dim b: b = dbConnect
  dim i
  dbGetData("SELECT COUNT(*) AS QTY FROM CATEGORIAS_PROVEEDORES")
  i = rs("QTY")
  dbReleaseData
  if i > 0 then
    redim preserve formSuppliersQueryTypeNames(i + 2)
    redim preserve formSuppliersQueryTypeSearchConditions(i + 2)
    dbGetData("SELECT * FROM CATEGORIAS_PROVEEDORES ORDER BY ID")
    i = 3
    do while not rs.EOF
      formSuppliersQueryTypeNames(i) = "Rubro: " & rs("NOMBRE")
      formSuppliersQueryTypeSearchConditions(i) = "VIGENTE=1 AND ID_CATEGORIA=" & rs("ID")
      rs.moveNext
      i = i + 1
    loop
    dbReleaseData
  end if
  if b then dbDisconnect
end function

function formSuppliersReviews
  usrAccessLevel = usrAccessAdminSuppliers
  formTitle = "Valoraciones de los vecinos"
  formTable = "PROVEEDORES_VALORACIONES"
  parentFormId = "formSuppliers"
  keyFieldName = "ID_PROVEEDOR"
  useAuditData = false
  
  formGridViewRowCount = 28
  formGridColumns = array("dbo.UNIDAD_VECINO(ID_VECINO)", "dbo.NOMBRE_VECINO(ID_VECINO)", "dbo.NOMBRE_VALORACION(VALORACION)")
  formGridColumnTypes = array(formGridColumnGeneralCenter, formGridColumnGeneral, formGridColumnGeneralCenter)
  formGridColumnLabels = array("Unidad", "Familia", "Valoración")
  formGridColumnWidths = array(60, 120, 70)
  gridViewOrderBy = "1"
  defaultQueryLimit = -1

  formRecordViewRenderFunc = null
end function

%>
