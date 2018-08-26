<%

function formAdminParams
  usrAccessLevel = usrAccessAdminParams
  formContainerCssClass = "paramsFormContainer"
  formContainerTitleCssClass = "paramsFormContainerTitle"
  formTitle = "Parámetros del sistema"
  forms = array("formCommissions", "formTimelineCategories", "formSuppliersCategories", "formSurveysCategories", "formClassifiedsCategories")
  childFormIds = "formCommissions,formTimelineCategories,formSuppliersCategories,formSurveysCategories,formClassifiedsCategories"
end function

function formCommissions
  usrAccessLevel = usrAccessAdminParams
  formTitle = "Comisiones"
  formTable = "COMISIONES"
  
  formGridViewRowCount = 7
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnWidths = array(25, 140)
  formGridViewShowFooter = false

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 130

  recordViewFields = array("ID", "NOMBRE", "HABILITADA", "SIGLA", "COLOR","MAILS")
  recordViewDBFields = array(true, true, true, true, true,true)
  recordViewReadOnlyFields = array(false, false, false, false, false,false)
  recordViewFieldLabels = array("Código", "Nombre", "Habilitada", "Sigla", "Color","Mails")
  recordViewFieldDefaults = array(null, null, true, null, null,null)
  recordViewNullableFields = array(false, false, false, false, false,false)
  recordViewFieldRenderFuncs = array("renderRecordViewNumericField", "renderRecordViewLiteralField", "renderRecordViewBooleanField", _
    "renderRecordViewLiteralField", "renderRecordViewLiteralField","renderRecordViewLiteralField")
end function

function formTimelineCategoriesBeforeDelete
  formTimelineCategoriesBeforeDelete = true
  if dbGetData("SELECT TOP 1 * FROM LINEA_TIEMPO WHERE ID_CATEGORIA = " & recordId) then
    formTimelineCategoriesBeforeDelete = reportError("No es posible eliminar esta categoría porque existen datos dependientes de ella.")
  end if
  dbReleaseData
end function

function formTimelineCategories
  usrAccessLevel = usrAccessAdminParams
  formTitle = "Categorías del Archivo Histórico"
  formTable = "CATEGORIAS_LINEA_TIEMPO"
  
  formBeforeDeleteFunc = "formTimelineCategoriesBeforeDelete"
  
  formGridViewRowCount = 7
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnWidths = array(25, 140)
  formGridViewShowFooter = false

  recordViewFieldLeftPos = 110
  recordViewEditboxWidth = 130
  suggestIdBy = 1

  recordViewFields = array("ID", "NOMBRE", "VIGENTE", "MOSTRAR")
  recordViewDBFields = array(true, true, true, true)
  recordViewReadOnlyFields = array(not nullRecord and recordId <= 6, false, false, false)
  recordViewFieldLabels = array("Código", "Nombre", "Vigente", "Visible en Historico")
  recordViewFieldDefaults = array(null, null, true, true)
  recordViewNullableFields = array(false, false, true, true)
  recordViewFieldRenderFuncs = array("renderRecordViewNumericField", "renderRecordViewLiteralField", _
    "renderRecordViewBooleanField", "renderRecordViewBooleanField")
end function

function formSuppliersCategories
  usrAccessLevel = usrAccessAdminParams
  formTitle = "Rubros de proveedores"
  formTable = "CATEGORIAS_PROVEEDORES"
  
  formGridViewRowCount = 7
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnWidths = array(25, 140)
  formGridViewShowFooter = false

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 130

  recordViewFields = array("ID", "NOMBRE", "VIGENTE")
  recordViewDBFields = array(true, true, true)
  recordViewReadOnlyFields = array(false, false, false)
  recordViewFieldLabels = array("Código", "Nombre", "Vigente")
  recordViewFieldDefaults = array(null, null, true)
  recordViewNullableFields = array(false, false, false)
  recordViewFieldRenderFuncs = array("renderRecordViewNumericField", "renderRecordViewLiteralField", "renderRecordViewBooleanField")
end function

function formSurveysCategories
  usrAccessLevel = usrAccessAdminParams
  formTitle = "Categorías de encuestas"
  formTable = "CATEGORIAS_ENCUESTAS"
  
  formGridViewRowCount = 7
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnWidths = array(25, 140)
  formGridViewShowFooter = false

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 130

  recordViewFields = array("ID", "NOMBRE", "VIGENTE")
  recordViewDBFields = array(true, true, true)
  recordViewReadOnlyFields = array(false, false, false)
  recordViewFieldLabels = array("Código", "Nombre", "Vigente")
  recordViewFieldDefaults = array(null, null, true)
  recordViewNullableFields = array(false, false, false)
  recordViewFieldRenderFuncs = array("renderRecordViewNumericField", "renderRecordViewLiteralField", "renderRecordViewBooleanField")
end function

function formClassifiedsCategories
  usrAccessLevel = usrAccessAdminParams
  formTitle = "Rubros de avisos"
  formTable = "CATEGORIAS_AVISOS"
  
  formGridViewRowCount = 5
  formGridColumns = array("ID", "NOMBRE")
  formGridColumnTypes = array(formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnWidths = array(25, 140)
  formGridViewShowFooter = false

  recordViewFieldLeftPos = 80
  recordViewEditboxWidth = 130

  recordViewFields = array("ID", "NOMBRE", "VIGENTE")
  recordViewDBFields = array(true, true, true)
  recordViewReadOnlyFields = array(false, false, false)
  recordViewFieldLabels = array("Código", "Nombre", "Vigente")
  recordViewFieldDefaults = array(null, null, true)
  recordViewNullableFields = array(false, false, false)
  recordViewFieldRenderFuncs = array("renderRecordViewNumericField", "renderRecordViewLiteralField", "renderRecordViewBooleanField")
end function

%>

