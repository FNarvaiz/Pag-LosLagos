
<%

const clubHouseBookingResourceId = 100

function renderDefaultReport
  %>
  <center>
  EN PREPARACIÓN...
  </center>
  <%
end function
function renderListadoMascotas
  
  if dbGetData("SELECT V.UNIDAD,V.NOMBRE AS FAMILIA,M.NOMBRE AS MASCOTA,M.RAZA " & _
      "FROM VECINOS V INNER JOIN VECINOS_MASCOTAS M ON V.ID=M.ID_VECINO " & _
      "GROUP BY V.UNIDAD,V.NOMBRE,M.NOMBRE,M.RAZA " & _
      "ORDER BY V.NOMBRE") then
    %>
        
        <center><table class="reportLevel1" cellpadding="3" cellspacing="0" style="width: 40%">
        <thead>
        <tr><th colspan="3"></tr>
        <tr>
          <th><b>Lote</b></th>
          <th><b>Familia</b></th>
          <th><b>Mascota</b></th>
          <th><b>Raza</b></th>
        </tr>
        </thead>
        <tbody>
        <%
    do while not rs.EOF
      
      %>
      <tr>
        <td valign="top" align="center"><%= rs("UNIDAD") %></td>
        <td valign="top" align="center"><%= rs("FAMILIA") %></td>
        <td valign="top" align="center"><%= rs("MASCOTA") %></td>
        <td valign="top" align="center"><%= rs("RAZA") %></td>
      </tr>
      <%
      rs.moveNext
    loop
  
      %></tbody></table><br></center><%
  
  
  end if
  dbReleaseData
end function

function renderBookingListing
  'create the excel object
  Set objExcel = CreateObject("Excel.Application") 

'view the excel program and file, set to false to hide the whole process
  objExcel.Visible = True 

'add a new workbook
  Set objWorkbook = objExcel.Workbooks.Add 

'set a cell value at row 3 column 5
  objExcel.Cells(3,5).Value = "new value"

'change a cell value
  objExcel.Cells(3,5).Value = "something different"
  
'delete a cell value
  objExcel.Cells(3,5).Value = ""

'get a cell value and set it to a variable
  r3c5 = objExcel.Cells(3,5).Value

'save the new excel file (make sure to change the location) 'xls for 2003 or earlier
  objWorkbook.SaveAs "C:\vbsTest.xls" 

'close the workbook
  objWorkbook.Close 

'exit the excel program
  objExcel.Quit

'release objects
  Set objExcel = Nothing
  Set objWorkbook = Nothing
end function

%>