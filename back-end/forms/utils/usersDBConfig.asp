<%

const usrTable = "USUARIOS"

const usrNameField = "NOMBRE"
const usrEmailField = "EMAIL"
const usrLoginNameField = "NOMBRE_LOGIN"
const usrLoginPwdField = "CLAVE_LOGIN"
const usrEnabledField = "HABILITADO"
const usrDateRegistration = "FECHA_REGISTRACION"
const usrDateLastAccessField = "FECHA_ULTIMO_LOGIN"
const usrAccessMasterField = "PERMISO_MASTER"
const usrAccessAdminUsersField = "PERMISO_ADMIN_USUARIOS"

const usrSessionTable = "USUARIOS_SESIONES"

const usrSessionIdUsrField = "ID_USUARIO"
const usrSessionIdField = "ID_SESION"
const usrSessionLoginDateField = "FECHA_LOGIN"

const usrActivityTable = "USUARIOS_ACTIVIDADES"
const usrActivityIdUsrField = "ID_USUARIO"
const usrActivityDateField = "FECHA"
const usrActivityResourceField = "RECURSO"
const usrActivityParamsField = "PARAMETROS"

dim usrSessionIdleTimeout
dim usrSessionExpiration
if request.serverVariables("SERVER_NAME") = "localhost" then
  usrSessionIdleTimeout = 10000
  usrSessionExpiration = 10000
else
  usrSessionIdleTimeout = 720
  usrSessionExpiration = 720
end if

const usrProfileIT = 1  'Sistemas
const usrProfileAdministrator = 10  'Administrador
const usrProfileCommission = 20  'Comision
const usrProfileOperator = 30  'Varios
const usrProfileSecurity = 40  'Guardia

const usrProfileField = "ID_PERFIL"
const usrDefaultProfile = 30

dim usrProfile: usrProfile = usrDefaultProfile

' Application permission fieldnames

const usrAccessAdminNeighborsField = "PERMISO_VECINOS"
const usrAccessAdminTimeLineField = "PERMISO_LINEA_TIEMPO"
const usrAccessAdminBookingsField = "PERMISO_RESERVAS"
const usrAccessAdminSurveysField = "PERMISO_ENCUESTAS"
const usrAccessAdminClassifiedsField = "PERMISO_AVISOS"
const usrAccessAdminSuppliersField = "PERMISO_PROVEEDORES"
const usrAccessAdminContentsField = "PERMISO_CONTENIDOS"
const usrAccessAdminParamsField = "PERMISO_PARAMS"
const usrAccessAdminReportsField = "PERMISO_INFORMES"

' Application permissions

dim usrAccessAdminNeighbors: usrAccessAdminNeighbors = usrPermissionNone
dim usrAccessAdminTimeLine: usrAccessAdminTimeLine = usrPermissionNone
dim usrAccessAdminBookings: usrAccessAdminBookings = usrPermissionNone
dim usrAccessAdminSurveys: usrAccessAdminSurveys = usrPermissionNone
dim usrAccessAdminClassifieds: usrAccessAdminClassifieds = usrPermissionNone
dim usrAccessAdminSuppliers: usrAccessAdminSuppliers = usrPermissionNone
dim usrAccessAdminContents: usrAccessAdminContents = usrPermissionNone
dim usrAccessAdminParams: usrAccessAdminParams = usrPermissionNone
dim usrAccessAdminReports: usrAccessAdminReports = usrPermissionNone

function readUsrData
  usrProfile = rs(usrProfileField)
  if usrAccessAdminMaster then
    usrAccessAdminNeighbors = usrPermissionUnrestricted
    usrAccessAdminTimeLine = usrPermissionUnrestricted
    usrAccessAdminBookings = usrPermissionUnrestricted
    usrAccessAdminSurveys = usrPermissionUnrestricted
    usrAccessAdminClassifieds = usrPermissionUnrestricted
    usrAccessAdminSuppliers = usrPermissionUnrestricted
    usrAccessAdminContents = usrPermissionUnrestricted
    usrAccessAdminParams = usrPermissionUnrestricted
    usrAccessAdminReports = usrPermissionUnrestricted
  else
    usrAccessAdminNeighbors = rs(usrAccessAdminNeighborsField)
    usrAccessAdminTimeLine = rs(usrAccessAdminTimeLineField)
    usrAccessAdminBookings = rs(usrAccessAdminBookingsField)
    usrAccessAdminSurveys = rs(usrAccessAdminSurveysField)
    usrAccessAdminClassifieds = rs(usrAccessAdminClassifiedsField)
    usrAccessAdminSuppliers = rs(usrAccessAdminSuppliersField)
    usrAccessAdminContents = rs(usrAccessAdminParamsField) ' El permiso de parametros habilita Contenidos.
    usrAccessAdminParams = rs(usrAccessAdminParamsField)
    usrAccessAdminReports = rs(usrAccessAdminReportsField)
  end if
end function

function getUsrProfileInfo
  exit function
  
  dim b: b = dbConnect
  select case usrProfile
  end select
  if b then dbDisconnect
end function

%>
