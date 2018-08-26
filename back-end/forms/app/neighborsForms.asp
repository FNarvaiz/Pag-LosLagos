<%

function formAdminNeighbors
  usrAccessLevel = usrAccessAdminNeighbors
  formContainerCssClass = "neighborsFormContainer"
  formContainerTitleCssClass = "neighborsFormContainerTitle"
  formTitle = "Vecinos"
  forms = array("formNeighbors", "formNeighborsActivity")
  childFormIds = "formNeighbors"
end function

dim formNeighborsQueryTypeNames
formNeighborsQueryTypeNames = array("Habilitados", "Deshabilitados", "Habilitados con permiso p/servicios", _
  "Habilitados sin permiso p/servicios", "Todos")

dim formNeighborsQueryTypeSearchConditions
formNeighborsQueryTypeSearchConditions = array("HABILITADO=1", "HABILITADO=0", "HABILITADO=1 AND PERMISO_SERVICIOS=1", _
  "HABILITADO=1 AND PERMISO_SERVICIOS=0", "")

dim formNeighborsQuerySearchFieldNames
formNeighborsQuerySearchFieldNames = array("Por Nombre")

dim formNeighborsQuerySearchFields
formNeighborsQuerySearchFields = array("UNIDAD + ' ' + NOMBRE")

function formNeighbors
  usrAccessLevel = usrAccessAdminNeighbors
  formTitle = ""
  formTable = "VECINOS"
  childFormIds = "formNeighborsActivity"
  
  formGridViewRowCount = 27
  formGridColumns = array("dbo.SPACE_PAD(UNIDAD, 15)", "NOMBRE")
  formGridColumnTypes = array(formGridColumnGeneralCenter, formGridColumnGeneral)
  formGridColumnLabels = array("Unidad", "Familia")
  formGridColumnWidths = array(70, 200)
  gridViewReordering = false
  gridViewOrderBy = "1"
  defaultQueryLimit = -1

  recordViewFieldLeftPos = 110
  recordViewEditboxWidth = 180

  recordViewIdFieldIsIdentity = true
  recordViewSeparators = array("Datos básicos", null, null, null, "Datos de acceso", null, null, null, "Notas") 
  recordViewFields = array("ID", "UNIDAD", "NOMBRE", "TELEFONOS", "EMAIL", "CLAVE", "HABILITADO", "PERMISO_SERVICIOS", "NOTAS")
  recordViewDBFields = array(true, true, true, true, true, true, true, true, true)
  recordViewReadOnlyFields = array(true, false, false, false, false, false, false, false, false)
  recordViewFieldLabels = array("Código", "Unidad/Lote", "Familia", "Teléfonos", "e-mail", "Contraseña", "Acceso habilitado", "Permiso p/servicios", "")
  recordViewFieldDefaults = array(null, null, null, null, null, null, true, true, null)
  recordViewNullableFields = array(true, false, false, true, false, false, false, false, true)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", "renderRecordViewLiteralField", _
    "renderRecordViewNameField", "renderRecordViewLiteralField", "renderRecordViewLiteralField", _
    "renderRecordViewLiteralField", "renderRecordViewBooleanField", "renderRecordViewBooleanField", "renderRecordViewNotesField(10)")
end function

function formNeighborsActivity
  usrAccessLevel = usrAccessAdminNeighbors

  formTitle = ""
  formTable = "VECINOS_ACTIVIDADES"
  parentFormId = "formNeighbors"
  keyFieldName = "ID_VECINO"
  
  formGridViewRowCount = 30
  formGridColumns = array("FECHA", "ACTIVIDAD", "ACTIVIDAD_DETALLES")
  formGridColumnTypes = array(formGridColumnDateTime, formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnLabels = array("Fecha", "Actividad", "Detalles")
  formGridColumnWidths = array(90, 110, 140)
  gridViewReordering = true
  gridViewOrderBy = "1 DESC"
  defaultQueryLimit = -1
  
  formRecordViewRenderFunc = null
end function

%>
