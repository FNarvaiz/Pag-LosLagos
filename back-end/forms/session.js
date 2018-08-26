
var sessionId = "";
var usrName = "";

function login()
{
  if (!document.loginForm.usr.value)
    document.loginForm.usr.focus();
  else if (!document.loginForm.pwd.value)
    document.loginForm.pwd.focus();
  else
    ajaxLoadingSignalOn();
    ajaxPostText(server + "?verb=login", "", serializeForm(document.loginForm), "", 
      function() {
      	document.loginForm.reset();
        if (ajaxResponseText)
        {
    		  eval("var obj = " + ajaxResponseText);
          if (obj.result == "OK")
          {
            sessionId = obj.sessionId;
            usrName = obj.usrName;
      		  document.loginForm.usr.blur(); // needed by mozilla!!
      		  document.loginForm.pwd.blur();
            hideDialog();
            enter();
          }
          else
          {
            ajaxLoadingSignalOff();
            alert(obj.message);
          }
        }
        else
          ajaxLoadingSignalOff();
      }
    );
}

function getExistingSession(currentSessionId)
{
  ajaxLoadingSignalOn();
  ajaxGetText(server + "?verb=getSessionInfo&sessionId=" + currentSessionId, "", "",
    function() {
      sessionId = "";
      if (ajaxResponseText)
      {
        eval("var obj = " + ajaxResponseText);
        if (obj.result == "OK")
        {
          sessionId = obj.sessionId;
          usrName = obj.usrName;
          enter();
        }
        else
        {
          ajaxLoadingSignalOff();
          start();
        }
      }
      else
      {
        ajaxLoadingSignalOff();
        start();
      }
    }
  );
}

function enter()
{
  ajaxGetText(server + "?sessionId=" + sessionId + "&verb=menu", "menu", "",
    function() {
      if (typeof(loginOK) == "function")
        loginOK();
      ajaxLoadingSignalOff();
      document.getElementById("initialMessage").style.visibility = "hidden";
    }
  );
}

function logout()
{
  if (sessionId)
    ajaxSynchronousGetText(server + "?sessionId=" + sessionId + "&verb=logout", "");
  sessionId = "";
}

function sendPasswordRemainder()
{
  if (document.loginForm.usr.value)
    ajaxGetText(server + "?verb=sendPasswordRemainder&usr=" + encodeURIComponent(document.loginForm.usr.value), "", loadingSignalId,
      function() {
        if (ajaxResponseText)
        {
    		  eval("var obj = " + ajaxResponseText);
          alert(obj.message);
        }
      }
    );
  else
  {
    alert("Por favor, ingrese su usuario.");
    document.loginForm.usr.focus();
  }
}


