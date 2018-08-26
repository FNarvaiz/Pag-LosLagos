
<!--#include file="usersDBConfig.asp"-->

<%

dim sessionChecked: sessionChecked = false
dim usr:            usr = getStringParam("usr", 50)
dim pwd:            pwd = getStringParam("pwd", 15)
dim sessionId:      sessionId = getStringParam("sessionId", 40)
dim usrId:          usrId = 0
dim usrName:        usrName = ""
dim usrEnabled:     usrEnabled = false

dim usrAccessAdminMaster:    usrAccessAdminMaster = false
dim usrAccessAdminUsers:     usrAccessAdminUsers = false

function usrExists
  dim b: b = dbConnect
  usrExists = dbGetData("SELECT ID FROM " & usrTable & " WHERE LOGIN_NAME=" & sQuotes(usr))
  dbReleaseData
  if b then dbDisconnect
end function

function registerNewUser
  dim b: b = dbConnect
  usrName = getStringParam("name", 100)
  on Error resume next
  dbExecute("INSERT INTO " & usrTable & " (" & join(array(usrNameField, usrEmailField, usrLoginNameField, usrLoginPwdField), ",") & _
    ") VALUES (" & join(array(sQuotes(usrName), sQuotes(usr), sQuotes(usr), sQuotes(pwd)), ",") & ")")
  registerNewUser = not failed
  if b then dbDisconnect
end function

function login
  if not isNull(usr) and not isNull(pwd) then
    dim b: b = dbConnect
    if dbGetData("SELECT ID, " & usrNameField & ", "& usrEnabledField & " FROM " & usrTable & " WHERE " & usrLoginNameField & "=" & _
      sQuotes(usr) & " AND " & usrLoginPwdField & "=" & sQuotes(pwd)) then
      usrId = rs("ID")
      usrName = rs(usrNameField)
      usrEnabled = rs(usrEnabledField)
    end if
    dbReleaseData
    if usrId > 0 then
      if usrEnabled then
        dbExecute("DELETE FROM " & usrSessionTable & " WHERE " & usrSessionIdUsrField & "=" & sQuotes(usrId))
        dbExecute("UPDATE " & usrTable & " SET " & usrDateLastAccessField & "=GETDATE() WHERE ID=" & usrId)
        dbExecute("INSERT INTO " & usrSessionTable & " (" & usrSessionIdUsrField & ") VALUES (" & sQuotes(usrId) & ")")
        'if usrId=28 then dbLog(usrName & " - LOGIN - " & request.serverVariables("remote_addr"))
        dbGetData("SELECT CONVERT(VARCHAR(100), " & usrSessionIdField & ") AS SESSION_ID FROM " & usrSessionTable & " WHERE " & _
          usrSessionIdUsrField & "=" & usrId)
        sessionId = rs("SESSION_ID")
        JSONAddOpOK
        JSONAddStr "sessionId", sessionId
        JSONAddStr "usrName", usrName
        logUserActivity "LOGIN", request.serverVariables("remote_addr")
      else
        JSONAddOpFailed
        JSONAddMessage "Su cuenta de usuario se encuentra deshabilitada."
      end if
    else
      JSONAddOpFailed
      JSONAddMessage "Los datos de acceso que ha ingresado son incorrectos."
    end if
  else
    JSONAddOpFailed
    JSONAddMessage "Los datos de acceso que ha ingresado son incorrectos."
  end if
  JSONSend
  if b then dbDisconnect
end function

function logout
  if getLoggedUsrData then
    dbExecute("DELETE FROM " & usrSessionTable & " WHERE " & usrSessionIdUsrField & "=" & sQuotes(usrId))
    dbExecute("DELETE FROM " & usrActivityTable & " WHERE " & usrActivityIdUsrField & "=" & sQuotes(usrId) & _
      " AND " & usrActivityDateField & " < DATEADD(MONTH, -1, GETDATE())")
    JSONAddOpOK
    logUserActivity "LOGOUT", ""
  else
    JSONAddOpFailed
  end if
  JSONSend
end function

function sendPasswordRemainder
  dim b: b = dbConnect
	if dbGetData("SELECT ID FROM " & usrTable & " WHERE " & usrLoginNameField & "=" & sQuotes(usr)) then
    JSONAddOpOK
    JSONAddMessage "Se envió un e-mail con los datos de acceso\na la dirección vinculada a la cuenta de usuario " & dQuotes(usr) & "."
  else
    JSONAddOpFailed
    JSONAddMessage "No existe una cuenta de usuario identificada como " & dQuotes(usr) & "."
	end if
  JSONSend
	dbReleaseData
	if b then dbDisconnect
end function

function getLoggedUsrData
  if sessionChecked then
    getLoggedUsrData = usrId > 0
  else
    dim b: b = dbConnect
    getLoggedUsrData = false
    if not isNull(sessionId) then
      dim dbNow, dateLastLogin, dateLastAccess
      usrId = -1
      if dbGetData("SELECT GETDATE() AS NOW, A." & usrSessionLoginDateField & ", B.ID, B." & usrDateLastAccessField & _
          " FROM " & usrSessionTable & " A, " & usrTable & " B" & _
          " WHERE A." & usrSessionIdUsrField & "=B.ID AND B." & usrEnabledField & "=1 AND A." & usrSessionIdField & "=" & _
          sQuotes(sessionId)) then
        usrId = rs("ID")
        dbNow = rs("NOW")
        dateLastLogin = rs(usrSessionLoginDateField)
        dateLastAccess = rs(usrDateLastAccessField)
      end if
      dbReleaseData
      if usrId > 0 then
        if dateDiff("n", dateLastAccess, dbNow) > usrSessionIdleTimeout or dateDiff("n", dateLastLogin, dbNow) > usrSessionExpiration then
          dbExecute("DELETE FROM " & usrSessionTable & " WHERE " & usrSessionIdUsrField & "=" & usrId)
          logUserActivity "SESSION EXPIRED", ""
        else
          dbGetData("SELECT * FROM " & usrTable & " WHERE ID=" & usrId)
          usrName = rs(usrNameField)
          usrAccessAdminMaster = rs(usrAccessMasterField)
          if usrAccessAdminMaster then
            usrAccessAdminUsers = true
          else
            usrAccessAdminUsers = rs(usrAccessAdminUsersField)
          end if
          readUsrData
          dbReleaseData
          if dateDiff("s", dateLastAccess, dbNow) > 15 then
            dbExecute("UPDATE " & usrTable & " SET " & usrDateLastAccessField & "=GETDATE() WHERE ID=" & usrId)
'            dbGetData("SELECT COUNT(*) FROM " & usrSessionTable & _
'              " WHERE (DATEDIFF(MINUTE, " & usrSessionLoginDateField & ", GETDATE()) <= " & usrSessionExpiration & ") AND " & _
'                usrSessionIdUsrField & " NOT IN (SELECT ID FROM USUARIOS WHERE DATEDIFF(MINUTE, " & usrDateLastAccessField & _
'                  ", GETDATE()) > " & usrSessionIdleTimeout & ")")
'            dim i: i = rs(0)
'            if i = 1 then
'            end if
            dbExecute("INSERT INTO " & usrActivityTable & " (" & usrActivityIdUsrField & ", " & usrActivityResourceField & ", " & _
              usrActivityParamsField & ") " & _
              "SELECT " & usrSessionIdUsrField & ", 'SESSION EXPIRED', 'MAX DURATION' FROM " & usrSessionTable & _
              " WHERE DATEDIFF(MINUTE, " & usrSessionLoginDateField & ", GETDATE()) > " & usrSessionExpiration)
            dbExecute("INSERT INTO " & usrActivityTable & " (" & usrActivityIdUsrField & ", " & usrActivityResourceField & ", " & _
              usrActivityParamsField & ") " & _
              "SELECT ID, 'SESSION EXPIRED', 'IDLE TIMEOUT' FROM " & usrTable & _
              " WHERE DATEDIFF(MINUTE, " & usrDateLastAccessField & ", GETDATE()) > " & usrSessionIdleTimeout & " AND " & _
                 "ID IN (SELECT " & usrSessionIdUsrField & " FROM " & usrSessionTable & ")") 
            dbExecute("DELETE FROM " & usrSessionTable & _
              " WHERE (DATEDIFF(MINUTE, " & usrSessionLoginDateField & ", GETDATE()) > " & usrSessionExpiration & ") OR " & _
                usrSessionIdUsrField & " IN (SELECT ID FROM " & usrTable & " WHERE DATEDIFF(MINUTE, " & usrDateLastAccessField & _
                  ", GETDATE()) > " & usrSessionIdleTimeout & ")")
          end if
          getLoggedUsrData = true
        end if
      end if
    end if
    if b then dbDisconnect
    sessionChecked = true
  end if
end function

function getSessionInfo
  if getLoggedUsrData then
    JSONAddOpOK
    JSONAddStr "sessionId", sessionId
    JSONAddStr "usrName", usrName
  else
    JSONAddOpFailed
  end if
  JSONSend
end function

function logUserActivity(resourceName, parameters)
  if usrId < 0 then exit function ' or usrProfile = usrProfileIT 
  dim b: b = dbConnect
  dbExecute("INSERT INTO " & usrActivityTable & " (" & join(array(usrActivityIdUsrField, usrActivityResourceField, usrActivityParamsField), ",") & _
    ") VALUES (" & join(array(usrId, sqlValue(left(resourceName, 100)), sQuotes(left(parameters, 100))), ",") & ")")
  if b then dbDisconnect
end function

%>

