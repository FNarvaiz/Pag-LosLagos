<%

dim lang: lang = getStringParam("lang", 2): if isNull(lang) then lang = "SP"

dim mainMenuLabelsSP
mainMenuLabelsSP = array("Inicio", "Quienes&nbsp;somos", "Hístorico", "Documentos", "Reservas", "Costos", "Actividades",  "Mi perfil", "Datos Utiles")
dim mainMenuLabelsEN
mainMenuLabelsEN = array("HOME", "ABOUT US", "TIME CHART", "LISTING", "CLASSIFIEDS", "RESERVATIONS", "SURVEYS", "CONTRACTORS", "MY PROFILE", "HELP")

dim mainMenuContents
mainMenuContents = array("home", "about", "timeLine", "downloads", "bookings", "classifieds", "surveys","profile", "help")
'mainMenuContents = array("home", "timeLine", "downloads", "bookings", "classifieds", "surveys", "suppliers", "help")

const noDataMsgSP = "(No hay datos disponibles en este momento.)"
const noDataMsgEN = "(There's no available data at this moment.)"

const noSearchResultsMsgSP = "No hay datos que cumplan con el criterio de búsqueda"
const noSearchResultsMsgEN = "There are no records matching the search criteria."

'dim timeLineRowLabelsSP: timeLineRowLabelsSP = array()
'dim timeLineRowLabelsEN: timeLineRowLabelsEN = array()

dim monthNamesSP: monthNamesSP = array("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SET", "OCT", "NOV", "DIC")
dim monthNamesEN: monthNamesEN = array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")

const logoutLabelSP = "Salir"
const logoutLabelEN = "Logout"
const changePasswordLabelSP = "Cambiar contraseña"
const changePasswordLabelEN = "Change password"

const dialogOKBtnLabelSP = "Aceptar"
const dialogOKBtnLabelEN = "OK"
const dialogCancelBtnLabelSP = "Cancelar"
const dialogCancelBtnLabelEN = "Cancel"

const loginMenuItemSP = "INGRESO AL SISTEMA"
const loginMenuItemEN = "SYSTEM LOGIN"
const registrationMenuItemSP = "REGISTRATE"
const registrationMenuItemEN = "REGISTER"
const faqMenuItemSP = "AYUDA/FAQ"
const faqMenuItemEN = "HELP/FAQ"

const loginDialogUsrFieldLabelSP = "E-MAIL"
const loginDialogUsrFieldLabelEN = "E-MAIL"
const loginDialogPwdFieldLabelSP = "CONTRASEÑA"
const loginDialogPwdFieldLabelEN = "PASSWORD"
const loginDialogSubmitBtnLabelSP = "INGRESAR"
const loginDialogSubmitBtnLabelEN = "LOGIN"
const loginDialogPasswordRecoveryBtnLabelSP = "Olvidé mi contraseña"
const loginDialogPasswordRecoveryBtnLabelEN = "Forgot my password"

const changePasswordDialogTitleSP = "CAMBIO DE CONTRASEÑA"
const changePasswordDialogTitleEN = "CHANGE PASSWORD"
const changePasswordDialogCurrentPwdFieldLabelSP = "Contraseña actual"
const changePasswordDialogCurrentPwdFieldLabelEN = "Current password"
const changePasswordDialogPwd1FieldLabelSP = "Nueva contraseña"
const changePasswordDialogPwd1FieldLabelEN = "New password"
const changePasswordDialogPwd2FieldLabelSP = "Repítala, por favor"
const changePasswordDialogPwd2FieldLabelEN = "Repeat new password"
const changePasswordDialogSubmitBtnLabelSP = "ENVIAR"
const changePasswordDialogSubmitBtnLabelEN = "SEND"
const changePasswordDialogCancelBtnLabelSP = "CANCELAR"
const changePasswordDialogCancelBtnLabelEN = "CANCEL"

const registrationDialogTitleSP = "Ingresá tus datos y enviá tu solicitud."
const registrationDialogTitleEN = "Please, complete this registration form."

const changePasswordResultMsgOKSP = "Se ha cambiado la contraseña."
const changePasswordResultMsgOKEN = "Your password has been changed."
const changePasswordResultMsgFailedSP = "La contraseña que indicó como actual no es correcta.<br>El cambio de contraseña no se realizó."
const changePasswordResultMsgFailedEN = "Current password is not correct.<br>Your password has not been changed."
const changePasswordResultBtnLabelSP = "► FINALIZAR"
const changePasswordResultBtnLabelEN = "► FINISH"

const passwordRecoveryDialogTitleSP = "Ingrese su e-mail y le enviaremos la contraseña."
const passwordRecoveryDialogTitleEN = "Enter your e-mail address to receive your password."
const passwordRecoveryDialogUsrFieldLabelSP = "E-MAIL"
const passwordRecoveryDialogUsrFieldLabelEN = "E-MAIL"
const passwordRecoveryDialogSubmitBtnLabelSP = "ENVIAR"
const passwordRecoveryDialogSubmitBtnLabelEN = "SEND"

const searchFieldLabelSP = "Buscar"
const searchFieldLabelEN = "Search"

%>