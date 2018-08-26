<%

' Forms parametrization

function formAdminUsers
  if usrAccessAdminUsers then
    usrAccessLevel = usrPermissionDelete
  else
    usrAccessLevel = usrPermissionNone
  end if

  formContainerCssClass = "usersFormContainer"
  formContainerTitleCssClass = "usersFormContainerTitle"
  formTitle = "Usuarios de la administración"
  if usrAccessAdminMaster then
    forms = array("formUsers", "formUsersReports", "formUsersActivity")
  else
    forms = array("formUsers", "formUsersReports")
  end if
  childFormIds = "formUsers"
end function

function formUsersBeforeUpdateFunc
  formUsersBeforeUpdateFunc = true
  if cInt(fieldNewValues(getFieldIndex(usrProfileField))) = usrProfileCommission then
    if isNull(fieldNewValues(getFieldIndex("ID_COMISION"))) then
      formUsersBeforeUpdateFunc = reportError("Debe indicar la Comisión a la que pertenece el usuario.")
      exit function
    end if
    setFieldNewValue usrAccessMasterField, 0
    setFieldNewValue usrAccessAdminUsersField, 0
    setFieldNewValue usrAccessAdminNeighborsField, usrPermissionReadOnly
    setFieldNewValue usrAccessAdminTimeLineField, usrPermissionDelete
    setFieldNewValue usrAccessAdminBookingsField, usrPermissionNone
    setFieldNewValue usrAccessAdminSurveysField, usrPermissionNone
    setFieldNewValue usrAccessAdminClassifiedsField, usrPermissionNone
    setFieldNewValue usrAccessAdminSuppliersField, usrPermissionNone
    setFieldNewValue usrAccessAdminReportsField, usrPermissionReadOnly
    setFieldNewValue usrAccessAdminParamsField, usrPermissionNone
  else
    setFieldNewValue "ID_COMISION", null
  end if
end function

function formUsers
  if usrAccessAdminUsers then
    usrAccessLevel = usrPermissionDelete
  else
    usrAccessLevel = usrPermissionNone
  end if

  formTitle = ""
  formTable = usrTable
  if usrAccessAdminMaster then
    childFormIds = "formUsersReports,formUsersActivity"
  else
    childFormIds = "formUsersReports"
  end if
  formBeforeUpdateFunc = "formUsersBeforeUpdateFunc"
  
  formGridViewRowCount = 19
  if usrProfile = usrProfileIT then
    formGridColumns = array("ID", usrNameField, "dbo.NOMBRE_PERFIL_USUARIO(ID_PERFIL)", _
      "DATEDIFF(minute, " & usrDateLastAccessField & ", GETDATE())", _
      "CASE WHEN EXISTS(SELECT * FROM USUARIOS_SESIONES S WHERE S.ID_USUARIO=ID) THEN 'hot' ELSE '' END")
    formGridColumnTypes = array(formGridColumnGeneralCenter, formGridColumnGeneral, formGridColumnGeneral, _
      formGridColumnGeneralCenter, formGridColumnHidden)
    formGridColumnWidths = array(30, 110, 90, 60, 0)
    formGridColumnLabels = array("#", "Nombre", "Perfil", "Min.", "")
    gridViewOrderBy = "2"
    gridViewReordering = true
    formGridRowCssClassColumn = uBound(formGridColumns)
  else
    formGridColumns = array(usrNameField, "dbo.NOMBRE_PERFIL_USUARIO(ID_PERFIL)")
    formGridColumnTypes = array(formGridColumnGeneral, formGridColumnGeneral)
    formGridColumnWidths = array(180, 110)
    formGridColumnLabels = array("Nombre", "Perfil")
    gridViewOrderBy = "1"
    gridViewReordering = false
  end if
  
  recordViewFieldLeftPos = 95
  recordViewEditboxWidth = 170
  recordViewIdFieldIsIdentity = true
  recordViewLookupFieldNameField = "NOMBRE"
  recordViewButtons = array(true, recordId <> usrId, true)

  dim masterAccessFieldFunc, usrProfileFieldFunc
  if not usrAccessAdminMaster then
    addSearchExpr(usrAccessMasterField & "=0")
    masterAccessFieldFunc = "renderRecordViewHiddenField"
    usrProfileFieldFunc = "renderRecordViewLookupField(" & dQuotes("USUARIOS_PERFILES,ID,ID>1") & ")"
  else
    masterAccessFieldFunc = "renderRecordViewBooleanField"
    usrProfileFieldFunc = "renderRecordViewLookupField(" & dQuotes("USUARIOS_PERFILES,ID") & ")"
  end if

  recordViewSeparators = array("General", null, null, null, _
    "Log-in", null, null , null, "Perfil de usuario", null, _
    "Permisos", null, _
    null, null, _
    null, null, null, null, _
    null, null)
  recordViewFields = array("ID", usrNameField, usrEmailField, usrDateRegistration, _
    usrLoginNameField, usrLoginPwdField, usrDateLastAccessField, usrEnabledField, usrProfileField, "ID_COMISION", _
    usrAccessAdminUsersField, usrAccessMasterField, _
    usrAccessAdminNeighborsField, usrAccessAdminTimeLineField, _
    usrAccessAdminBookingsField, usrAccessAdminSurveysField, usrAccessAdminClassifiedsField, usrAccessAdminSuppliersField, _
    usrAccessAdminReportsField, usrAccessAdminParamsField)
  recordViewDBFields = array(true, true, true, true, _
    true, true, true, true, true, true, _
    true, true, _
    true, true, _
    true, true, true, true, _
    true, true)
  dim b: b = (recordId = usrId)
  recordViewReadOnlyFields = array(true, false, false, true, _
    false, false, true, false, b, b, _
    b, b, _
    b, b, _
    b, b, b, b, _
    b, b)
  recordViewFieldLabels = array("Código", "Nombre", "e-mail", "Registración", _
    "Usuario", "Contraseña", "Último acceso", "Habilitado", "Perfil", "Comisión", _
    "Usuarios admin.", "Master", _
    "Vecinos", "Línea de tiempo", _
    "Reservas", "Encuestas", "Avisos", "Proveedores", _
    "Informes", "Parámetros")
  recordViewFieldDefaults = array(0, null, null, null, _
    null, null, null, true, 20, null, _
    false, false, _
    usrPermissionNone, usrPermissionNone, _
    usrPermissionNone, usrPermissionNone, usrPermissionNone, usrPermissionNone, _
    usrPermissionReadOnly, usrPermissionNone)
  recordViewNullableFields = array(false, false, false, true, _
    false, false, true, false, false, true, _
    false, false, _
    false, false, _
    false, false, false, false, _
    false, false)
  dim s: s = "renderRecordViewLookupField(" & dQuotes("USUARIOS_NIVELES_PERMISOS,ID") & ")"
  recordViewFieldRenderFuncs = array("renderRecordViewHiddenField", "renderRecordViewLiteralField", _
    "renderRecordViewLiteralField", "renderRecordViewTimeStampField", _
    "renderRecordViewLiteralField", "renderRecordViewLiteralField", "renderRecordViewTimeStampField", "renderRecordViewBooleanField", _
    usrProfileFieldFunc, "renderRecordViewLookupField(" & dQuotes("COMISIONES,ID") & ")", _
    "renderRecordViewBooleanField", masterAccessFieldFunc, _
    s, s, s, s, s, s, s, s)
end function

function formUsersReports
  if usrAccessAdminUsers then
    usrAccessLevel = usrPermissionDelete
  else
    usrAccessLevel = usrPermissionNone
  end if

  formTitle = ""
  formTable = "USUARIOS_INFORMES"
  parentFormId = "formUsers"
  keyFieldName = "ID_USUARIO"

  formGridViewRowCount = 5
  formGridColumns = array("(SELECT I.NOMBRE FROM INFORMES I WHERE I.ID=ID_INFORME)")
  formGridColumnTypes = array(formGridColumnGeneral)
  formGridColumnLabels = array("Informes habilitados")
  formGridColumnWidths = array(295)
  gridViewReordering = false
  formGridViewShowFooter = false
  gridViewOrderBy = "ID_INFORME"

  recordViewFieldLeftPos = 60
  recordViewEditboxWidth = 190
  recordViewIdFieldIsIdentity = true

  recordViewFields = array("ID", "ID_INFORME")
  recordViewDBFields = array(true, true)
  recordViewReadOnlyFields = array(true, false)
  recordViewFieldLabels = array("Código", "Informe")
  recordViewFieldDefaults = array(null, null)
  recordViewNullableFields = array(true, false)
  recordViewFieldRenderFuncs = array("renderRecordViewIdentityField", _
    "renderRecordViewLookupField(" & dQuotes("INFORMES,ID,VISIBLE=1") & ")")
end function

function formUsersActivity
  if usrAccessAdminUsers then
    usrAccessLevel = usrPermissionDelete
  else
    usrAccessLevel = usrPermissionNone
  end if

  formTitle = ""
  formTable = "USUARIOS_ACTIVIDADES"
  parentFormId = "formUsers"
  keyFieldName = "ID_USUARIO"
  
  formGridViewRowCount = 30
  formGridColumns = array("FECHA", "RECURSO", "PARAMETROS")
  formGridColumnTypes = array(formGridColumnDateTime, formGridColumnGeneral, formGridColumnGeneral)
  formGridColumnLabels = array("Fecha", "Actividad", "Detalles")
  formGridColumnWidths = array(90, 130, 120)
  gridViewReordering = true
  gridViewOrderBy = "1 DESC"
  defaultQueryLimit = -1
  
  formRecordViewRenderFunc = null
end function

%>
