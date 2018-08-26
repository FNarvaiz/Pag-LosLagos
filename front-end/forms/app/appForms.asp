
<!--#include file="profileForms.asp"-->

<%

' App messages

const noPermissionForOperation = "No tiene autorización para realizar esta operación."

const mainEmailName = "Vecinos de los Lagos"
const mainEmailAddress = "sistema@vecinosloslagos.com.ar"

' Custom verbs

function handleCustomVerbs
  handleCustomVerbs = false
  exit function
  handleCustomVerbs = true
  select case verb
    case else
      if getLoggedUsrData then
        select case verb
          case else handleCustomVerbs = false
        end select
      else
        handleCustomVerbs = false
      end if
  end select
end function

%>
