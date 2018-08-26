function renderStandardListingHeader

end function

function renderStandardListingRow

end function

function renderStandardListingFooter

end function

listingHeaderRenderFunc = "renderStandardListingHeader"
listingRowRenderFunc = "renderStandardListingRow"
listingFooterRenderFunc = "renderStandardListingFooter"

function renderListing(tableName, searchCondition, order, headings, columns, columnWidths, id, itemsPerPage, pageNum, onPgUp, onPgDn)
  dbConnect
	if searchCondition = "" then searchCondition = "1=1"
	if order = "" then order = "ID"
 	dbGetData("SELECT COUNT(*) AS ITEMCOUNT FROM " & tableName & " WHERE " & searchCondition)
	itemCount = rs("ITEMCOUNT")
	dbReleaseData
	pageCount = itemCount \ itemsPerPage
	if (itemCount mod itemsPerPage) > 0 then pageCount = pageCount + 1
	if pageNum > 1 then
	  %>
	  <div class="PgUp" id="<%= id %>PgUpBtn" onclick="<%= onPgUp %>('<%= pageNum - 1 %>')"></div>
		<%
	end if

	dbGetData("SELECT TOP " & itemsPerPage & " " & join(columns, ",") & " FROM " & tableName & _
	  " WHERE NOT ID IN (SELECT TOP " & itemsPerPage * (pageNum - 1) & " ID FROM " & tableName & ") AND " & searchCondition & _
		" ORDER BY " & order)
  eval(listingHeaderRenderFunc)
	i = 0
	do while not rs.EOF
	  %>
	  <div class="listingRow" id="<%= id %>Row">
  		<%
		  eval(listingRowRenderFunc & "(" & i & "," & itemCount & ")")
  		%>
		</div>
		<%
	  rs.MoveNext
		i = i + 1
	loop
  eval(listingFooterRenderFunc)
  dbReleaseData
	dbDisconnect

	if pageCount > pageNum then
	  %>
	  <div class="PgDn" id="<%= id %>PgDnBtn" onclick="<%= onPgDn %>('<%= pageNum + 1 %>')"></div>
		<%
	end if
end function

