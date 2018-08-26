<%

dim usrId: usrId = -1
dim usrEnabled: usrEnabled = false
dim usrPwd: usrPwd = ""
dim usrName: usrName = ""
dim usrServicesPemission: usrServicesPemission = false
dim sessionId: sessionId = getStringParam("sessionId", 50): if isNull(sessionId) then sessionId = -1

const activityLogin = "LOGIN"
const activityLogout = "LOGOUT"
const activityPasswordRemainder = "Recupero de contraseña"
const activityChangePassword = "Cambio de contraseña"

function usrIP
  dim IP
  IP = request.serverVariables("HTTP_X_FORWARDED_FOR")
  if len(IP) > 0 then 
    IP = IP & " (proxy)"
  else
    IP = request.serverVariables("REMOTE_ADDR")
  end if
  usrIP = IP
end function

function logActivity(activityId, activityInfo)
  if usrId >= 0 then
    dbExecute("DELETE FROM VECINOS_ACTIVIDADES WHERE ID_VECINO=" & usrId & " AND FECHA < DATEADD(DAY, -60, GETDATE())") 
    dbExecute("INSERT INTO VECINOS_ACTIVIDADES (ID_VECINO, ACTIVIDAD, ACTIVIDAD_DETALLES, IP) VALUES (" & usrId & ", " & _
      "N" & sQuotes(activityId) & ", N" & sQuotes(activityInfo) & ", N" & sQuotes(usrIP) & ")")
  end if
end function

function login
  dim usr: usr = getStringParam("usr", 50)
	dim pwd: pwd = getStringParam("pwd", 20)
	if dbGetData("SELECT ID, NOMBRE, HABILITADO FROM VECINOS WHERE EMAIL=" & sQuotes(usr) & " AND CLAVE=" & sQuotes(pwd)) then
	  usrId = rs("ID")
	  usrName = rs("NOMBRE")
	  usrEnabled = rs("HABILITADO")
	end if
	dbReleaseData
	if usrId < 0 then
    JSONAddOpFailed
    JSONAddMessage "Los datos de acceso indicados no son correctos."
    JSONSend
	  exit function
	end if
	if not usrEnabled then
    JSONAddOpFailed
    JSONAddMessage "Su acceso se encuentra deshabilitado por la Administración."
    JSONSend
    logActivity activityLogin, "Denegado: usuario deshabilitado"
	  exit function
	end if
  dbExecute("DELETE FROM VECINOS_SESIONES WHERE ID_VECINO=" & usrId)
 	dbExecute("INSERT INTO VECINOS_SESIONES (ID_VECINO) VALUES (" & usrId & ")")
  dbGetData("SELECT CONVERT(VARCHAR(100), ID_SESION) FROM VECINOS_SESIONES WHERE ID_VECINO=" & usrId)
  sessionId = rs(0)
	dbReleaseData
  logActivity activityLogin, "OK. IP=" & usrIP
  JSONAddOpOK
  JSONAdd "sessionId", sessionId
  JSONAdd "usrName", usrName
  JSONSend
end function

function logout
  if getUsrData then
    dbExecute("DELETE FROM VECINOS_SESIONES WHERE ID_VECINO=" & usrId)
    logActivity activityLogout, "usr=" & usr & ", pwd=" & pwd & ", IP=" & usrIP
    JSONAddOpOK
  else
    JSONAddOpFailed
  end if
end function

function sendPasswordRemainder
  dim usr: usr = getStringParam("usr", 50)
	if dbGetData("SELECT ID FROM VECINOS WHERE EMAIL=" & sQuotes(usr)) then
    JSONAddOpOK
    JSONAddMessage "Se le ha enviado un e-mail con los datos de acceso."
    logActivity activityPasswordRemainder, "OK"
  else
    JSONAddOpFailed
    JSONAddMessage "La dirección de e-mail indicada no es correcta."
    logActivity activityPasswordRemainder, "El e-mail indicado es incorrecto"
	end if
	dbReleaseData
	dbDisconnect
	JSONSend
end function

function changePassword(currentPwd, newPwd)
  changePassword = false
  if getUsrData then
    response.write(usrPwd & " - " & currentPwd)
    if usrPwd = currentPwd then
      dbExecute("UPDATE VECINOS SET CLAVE=" & sQuotes(newPwd) & " WHERE ID=" & usrId)
      logActivity activityChangePassword, "OK"
      changePassword = true
    else
      logActivity activityChangePassword, "Contraseña actual incorrecta"
    end if
  end if
end function

function getUsrData
  getUsrData = false
  if sessionId = -1 then exit function
  if dbGetData("SELECT B.* FROM VECINOS_SESIONES A, VECINOS B WHERE A.ID_VECINO=B.ID AND B.HABILITADO=1 AND A.ID_SESION=" & sQuotes(sessionId)) then
	  usrId = rs("ID")
    usrPwd = rs("CLAVE")
  	usrName = rs("NOMBRE")
  	usrServicesPemission = rs("PERMISO_SERVICIOS")
		getUsrData = true
	end if
	dbReleaseData
end function

%>