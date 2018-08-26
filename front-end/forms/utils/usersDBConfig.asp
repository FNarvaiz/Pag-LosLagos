<%

const usrTable = "VECINOS"

const usrNameField = "NOMBRE"
const usrEmailField = "EMAIL"
const usrLoginNameField = "NOMBRE_LOGIN"
const usrLoginPwdField = "CLAVE_LOGIN"
const usrEnabledField = "HABILITADO"

const usrSessionTable = "VECINOS_SESIONES"
const usrSessionIdUsrField = "ID_VECINO"
const usrSessionIdField = "ID_SESION"
const usrSessionLoginDateField = "FECHA_LOGIN"

const usrActivityTable = "VECINOS_ACTIVIDADES"
const usrActivityIdUsrField = "ID_VECINO"
const usrActivityDateField = "FECHA"
const usrActivityResourceField = "ACTIVIDAD"
const usrActivityParamsField = "ACTIVIDAD_DETALLES"

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

const usrProfileField = "ID_PERFIL"
const usrDefaultProfile = 30

dim usrProfile: usrProfile = usrDefaultProfile
dim usrCommissionId: usrCommissionId = 0

' Application permission fieldnames

' Application permissions

function readUsrData
end function

function getUsrProfileInfo
  exit function
  
  dim b: b = dbConnect
  select case usrProfile
  end select
  if b then dbDisconnect
end function

%>
