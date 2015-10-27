<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conneshop.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="userlogin.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<SCRIPT LANGUAGE=JavaScript RUNAT=Server NAME="UC_CART">
//
// UltraDev UCart include file Version 1.0
//
function UC_ShoppingCart(Name, cookieLifetime, colNames, colComputed)  // Cart constructor
{
	// Name is the name of this cart. This is not really used in this implementation.
	// cookieLifeTime is in days. A value of 0 means do not use cookies.
	// colNames is a list of column names (must contain: ProductID, Quantity, Price, Total)
	// colComputed is a list of computed columns (zero length string means don't compute col.)

  // Public methods or UC_Cart API
  this.AddItem           = UCaddItem;        // Add an item to the cart
  this.GetColumnValue    = GetColumnValue;   // Get a value from the cart
  this.Destroy           = UCDestroy;        // remove all items, delete session, delete client cookie (if any)
	this.SaveToDatabase    = SaveToDatabase;   // persist cart to database.
	this.GetItemCount      = GetItemCount;     // the number of items in the cart.
	this.Update            = Update;           // Update the cart quantities.
	this.GetColumnTotal    = GetColumnTotal;   // Get the sum of a cart column for all items (e.g. price or shipping wt.).
  this.GetContentsSerial = UCGetContentsSerial// Get the contents of the cart as a single delimited string
  this.SetContentsSerial = UCSetContentsSerial// Set the contents of the cart from a serial string (obtained from GetContentsSerial)
  this.GetColNamesSerial = UCGetColNamesSerial// Get the list of column names as a delimited string.
  
	// PROPERTIES
	this.SC				= null;			// Cart data array
	this.numCols		= colNames.length;
	this.colComputed	= colComputed;
	this.colNames		= colNames;
	this.Name			= Name;
	this.cookieLifetime = cookieLifetime;
	this.bStoreCookie	= (cookieLifetime != 0);

	// *CONVENIENCE* PROPERTIES
	// (not used internally, but added to provide a place to store this data)
	this.CustomerID			= null;
	this.OrderID				= null;
	this.Tax						= null;
	this.ShippingCost		= null;

  // CONSTANTS
  this.PRODUCTID	= "ProductID";  // Required SKU cart column
  this.QUANTITY		= "Quantity";   // Required Quantity cart column
	this.PRICE			= "Price";			// Required Price cart column
	this.TOTAL			= "Total";			// Required Total column
  this.cookieColDel = "#UC_C#"
  this.cookieRowDel = "#UC_R#"

	// METHODS
	this.AssertCartValid = AssertCartValid

  // Private methods - don't call these unless you understand the internals.
  this.GetIndexOfColName = UCgetIndexOfColName;
  this.GetDataFromBindings = UCgetDataFromBindings;
	this.FindItem = UCfindItem;
	this.ComputeItemTotals = ComputeItemTotals;
  this.persist = UCpersist;
  
	this.BuildInsertColumnList = BuildInsertColumnList;
	this.BuildInsertValueList = BuildInsertValueList;
	this.UpdateQuantities = UpdateQuantities;
	this.UpdateTotals = UpdateTotals;
	this.DeleteItemsWithNoQuantity = DeleteItemsWithNoQuantity;
	this.CheckAddItemConfig = CheckAddItemConfig;
	this.ColumnExistsInRS = ColumnExistsInRS;
	this.DeleteLineItem = DeleteLineItem;
	this.GetCookieName = GetCookieName;
	this.SetCookie = SetCookie;
	this.PopulateFromCookie = PopulateFromCookie;
	this.DestroyCookie = UCDestroyCookie;

// Cart "internals" documentation:
// The this.SC datastructure is a single variable of type array.
// Each array element corresponds to a cart column. For example: 
//    Array element 1: ProductID
//    Array element 2: Quantity
//    Array element 3: Price
//    Array elemetn 4: Total
//
// Each of these is an array. Each array index corresponds to a line item.
// As such, each array should always be exactly the same length.
  this.AssertCartValid(colNames, "Cart Initialization: ");
	if (Session(this.Name) != null) {
		this.SC = Session(this.Name).SC;
	} else {
		this.SC = new Array(this.numCols);
		for (var i = 0; i < this.numCols; i++) this.SC[i] = new Array();

		// Since the cart doesn't exist in session, check for cookie from previous session
		if (this.bStoreCookie){
			cookieName = this.GetCookieName();
			cookieStr = Request.Cookies(cookieName);
			if (cookieStr != null && String(cookieStr) != "undefined" && cookieStr != "")
				this.PopulateFromCookie(cookieStr);
		}
		// Create a reference in the Session, pass the whole object (methods are not copied)
    this.persist();
	}  
}

// convert vb style arrays to js style arrays.
function UC_VbToJsArray(a) {
	if (a!=null && a.length==null) {
		a = new VBArray(a);
		a = a.toArray();
	}
	return a;
}

function UCpersist() {
  Session(this.Name) = this;
  if (this.bStoreCookie) this.SetCookie();
}

function UCDestroy(){
	this.SC = new Array(this.numCols);  // empty the "in-memory" cart.
	for (var i = 0; i < this.numCols; i++) this.SC[i] = new Array();
  this.persist();
	if (this.bStoreCookie) this.DestroyCookie() // remove the cookie
}

function UCgetDataFromBindings(adoRS, bindingTypes, bindingValues) {
	var values = new Array(bindingTypes.length)
	for (i=0; i<bindingTypes.length; i++) {
		var bindVal = bindingValues[i];
		if (bindingTypes[i] == "RS"){
			values[i] = String(adoRS(bindVal).Value)
			if (values[i] == "undefined") values[i] = "";
		}
		else if (bindingTypes[i] == "FORM"){
			values[i] = String(Request(bindVal))
			if (values[i] == "undefined") values[i] = "";
		} 
		else if (bindingTypes[i] == "LITERAL") values[i] = bindVal;
		else if (bindingTypes[i] == "NONE") values[i] = "";						// no binding
		else assert(false,"Unrecognized binding type: " + bindingTypes[i]);		// Unrecognized binding type
	}
	return values;
}

function UCfindItem(bindingTypes, values){
  // A product is a duplicate if it has the same unique ID
  // AND all values from form bindings (except quantity) are the same
  var indexProductID = this.GetIndexOfColName(this.PRODUCTID);
  var indexQuantity  = this.GetIndexOfColName(this.QUANTITY);
  assert(indexProductID >=0, "UC_Cart.js: Internal error 143");
  assert(indexQuantity >=0, "UC_Cart.js: Internal error 144");
	var newRow = -1
  for (var iRow=0; iRow<this.GetItemCount(); iRow++) {
    found = true;  // assume found
    for (var iCol=0; iCol<this.numCols; iCol++) {
      if (iCol != indexQuantity) {
        if ((iCol==indexProductID) || (bindingTypes[iCol]=="FORM")) {
          if (this.SC[iCol][iRow] != values[iCol]) {
            found = false;
            break;
        } }
    } }
    if (found) {
      newRow = iRow;
      break;
    }
  }
	return newRow
}

function UCaddItem(adoRS, bindingTypes, bindingValues, alreadyInCart){
  // alreadyInCart can be "increment" or "replace" to handle duplicate items in cart.
	bindingTypes = UC_VbToJsArray(bindingTypes);
	bindingValues = UC_VbToJsArray(bindingValues);

	// Check that length of binding types/values arrays is consistent with cart configuration
	assert(bindingTypes.length  == this.numCols, "UCaddItem: Array length mismatch (internal error 403)");
	assert(bindingValues.length == this.numCols, "UCaddItem: Array length mismatch (internal error 404)");

  // debug call
	//this.CheckAddItemConfig(adoRS, bindingTypes, bindingValues);

	var values = this.GetDataFromBindings(adoRS, bindingTypes, bindingValues) // get the actual values based on bindings
  var newRow = this.FindItem(bindingTypes, values);							// Check if this item is already in cart
  if (newRow == -1) {													// append a new item
		newRow = this.GetItemCount();    
    for (var iCol=0; iCol<this.numCols; iCol++) { // add data
      this.SC[iCol][newRow] = values[iCol];
    }
		this.ComputeItemTotals(newRow);						// add computed columns (defined in colsComputed)		
    this.persist();
	} else if (alreadyInCart == "increment") {
    var indexQuantity  = this.GetIndexOfColName(this.QUANTITY);
    this.SC[indexQuantity][newRow] = parseInt(this.SC[indexQuantity][newRow]) + parseInt(values[indexQuantity])
    if (isNaN(this.SC[indexQuantity][newRow])) this.SC[indexQuantity][newRow] = 1;
		this.ComputeItemTotals(newRow);
    this.persist();
	}
}

function UCgetIndexOfColName(colName) {
  var retIndex = -1;
  for (var i=0; i<this.numCols; i++) {
    if (this.colNames[i] == colName) {
      retIndex = i;
      break;
		} 
	}
  return retIndex;
}

function ComputeItemTotals(row){
	var indexQuantity = this.GetIndexOfColName(this.QUANTITY);
  var qty = parseInt(this.SC[indexQuantity][row])
	for (var iCol=0; iCol<this.numCols; iCol++) {
		var colToCompute = this.colComputed[iCol];
		if (colToCompute != "") {
		  indexColToCompute = this.GetIndexOfColName(colToCompute);
		  this.SC[iCol][row] = parseFloat(this.SC[indexColToCompute][row]) * qty;
		}
	}
}

function CheckAddItemConfig(adoRS, bindingTypes, bindingValues) {
	var ERR_SOURCE = "CheckAddItemConfig: "
	var ERR_RS_BINDING_VALUE = "Column for Recordset binding does not exist in recordset";
	// Check that all rs column names exist for rs binding types
	for (var i = 0; i < bindingTypes.length; i++) {
		if (bindingTypes[i] == "RS"){
			assert(this.ColumnExistsInRS(adoRS, bindingValues[i]), ERR_SOURCE + bindingValues[i] + ": " + ERR_RS_BINDING_VALUE);	
		}
	}  
}

function ColumnExistsInRS(adoRS, colName) {
	var bColExists = false;
	var items = new Enumerator(adoRS.Fields);
	while (!items.atEnd()) {
		if (items.item().Name == colName){
			bColExists = true;
			break;
		}
		items.moveNext();
	}
	return bColExists;
}

function GetColumnValue(colName, row){
	var retValue = "&nbsp;";
  var indexCol = this.GetIndexOfColName(colName);
	assert(!isNaN(row), "cart.GetColumnValue: row is not a number - row = " + row);
  assert(indexCol >=0, "cart.GetColumnValue: Could not find column \"" + colName + "\" in the cart");
  assert(row>=0, "cart.GetColumnValue: Bad row number input to cart - row = " + row);
  assert(this.GetItemCount()>0, "cart.GetColumnValue: The cart is empty - the requested data is unavailable");
  assert(row<this.GetItemCount(), "cart.GetColumnValue: The line item number is greater than the number of items in the cart - row = " + row + "; GetItemCount = " + this.GetItemCount());
  if (this.GetItemCount()>0) {
	  retValue = this.SC[indexCol][row];
	}
	return retValue;
}

function UpdateQuantities(formElementName) {
	var items = new Enumerator(Request.Form(formElementName))
	var j = 0;
  indexQuantity = this.GetIndexOfColName(this.QUANTITY);
	while(!items.atEnd()){
		var qty = parseInt(items.item());
		if (isNaN(qty) || qty < 0) {
		  this.SC[indexQuantity][j++] = 0
		} else {
		  this.SC[indexQuantity][j++] = qty;
		}
		items.moveNext();
	}
}

function UpdateTotals() {
  // this would be a little more efficient by making the outer loop over cols rather than rows.
	for (var iRow=0; iRow<this.GetItemCount(); iRow++) {
		this.ComputeItemTotals(iRow);
	}
}

function DeleteItemsWithNoQuantity() {
	var tmpSC= new Array(this.numCols);
  for (var iCol=0; iCol<this.numCols; iCol++) tmpSC[iCol] = new Array();

  var indexQuantity = this.GetIndexOfColName(this.QUANTITY);
  var iDest = 0;
	for (var iRow=0; iRow<this.GetItemCount(); iRow++) {    
    if (this.SC[indexQuantity][iRow] != 0) {
      for (iCol=0; iCol<this.numCols; iCol++) {
        tmpSC[iCol][iDest] = this.SC[iCol][iRow];
      }
      iDest++;
		}
	}
  this.SC = tmpSC;
}

function Update(formElementName){
	// Get new quantity values from Request object.
	// Assume they are all named the same, so you will get 
	// an array. The array length should be the same as the number
	// of line items and in the same order.
	this.UpdateQuantities(formElementName);
	this.DeleteItemsWithNoQuantity();
	this.UpdateTotals();
	this.persist();
}

function BuildInsertColumnList(orderIDCol, mappings){
	var colList = orderIDCol;
	for (var i = 0; i < mappings.length; i++) {
		if (mappings[i] != ""){
			colList += ", " + mappings[i];
		}
	}
	colList = "(" + colList + ")";
	return colList;
}

function BuildInsertValueList(orderIDColType, orderIDVal, destCols, destColTypes, row){
  var values = "";
  if (orderIDColType == "num") {
    values += orderIDVal;
  } else {
    values += "'" + orderIDVal.toString().replace(/'/g, "''") + "'";
  }

	for (var iCol=0; iCol<this.numCols; iCol++){
		if (destCols[iCol] != "") {
			if (destColTypes[iCol] == "num") {
        assert(this.SC[iCol][row] != "", "SaveToDatabase: A numeric value is missing in the SQL statement in column " + this.colNames[iCol]);
			  values += ", " + this.SC[iCol][row];
			} else {
			  values += ", '" + (this.SC[iCol][row]).toString().replace(/'/g, "''") + "'";  
			} 
		}	
	}
	values = "(" + values + ")";
	return values;
}

function SaveToDatabase(adoConn, dbTable, orderIDCol, orderIDColType, orderIDVal, destCols, destColTypes){
	// we are going to build SQL INSERT statements and 
	// throw it at the connection / table
	// Similar to existing UD insert to database behavior
	var ERR_MAPPINGS_LENGTH = "Array length must match the number of cart columns<BR>";
	var ERR_TRANS = "An error occured when inserting cart items in the database.  The transaction was rolled back<BR>";
	destCols = UC_VbToJsArray(destCols);
	destColTypes = UC_VbToJsArray(destColTypes);
	assert (destCols.length == this.numCols, "SaveToDatabase: " + "destCols - " + ERR_MAPPINGS_LENGTH);
	assert (destColTypes.length == this.numCols, "SaveToDatabase: " + "destColTypes - " + ERR_MAPPINGS_LENGTH);

	var insertColList = this.BuildInsertColumnList(orderIDCol, destCols);

	if (insertColList != "") { //proceed only if we have a column list
		var insertClause = "INSERT INTO " + dbTable + " " + insertColList + " VALUES ";
		var recs;
		adoConn.BeginTrans();
		for (var iRow=0; iRow<this.GetItemCount(); iRow++){
			var valList = this.BuildInsertValueList(orderIDColType, orderIDVal, destCols, destColTypes, iRow);
			var sql = insertClause + valList;
			adoConn.Execute(sql, recs, 1 /*adCmdText*/); 
		}
		if (adoConn.Errors.Count == 0){ 
			adoConn.CommitTrans();
			this.Destroy();	// All items saved to database, we can trash the cart
		}	else {
			adoConn.RollbackTrans();
			//assert(false, "SaveToDatabase: " + ERR_TRANS); Don't assert here - let ASP display the database error.
		}
	}
}

function GetItemCount(){
	return this.SC[0].length
}

function GetColumnTotal(colName){
	// Generic column Total function
	var colTotal = 0.0;
	index = this.GetIndexOfColName(colName);
	for (var i=0; i<this.SC[index].length; i++)
		colTotal += parseFloat(this.SC[index][i]);
    
	return colTotal
}


function DeleteLineItem(row){
	assert(!isNaN(row), "Failure in call to DeleteLineItem - row is not a number");
  assert(row>=0 && row <this.GetItemCount(), "failure in call to DeleteLineItem (internal error 121)");

	var tmpSC= new Array(this.numCols);
  var iDest = 0;
  for (var iCol=0; iCol<this.numCols; iCol++) tmpSC[iCol] = new Array();
  for (var iRow=0; iRow<this.GetItemCount(); iRow++) {
    if (iRow != row) {
      for (iCol=0; iCol<this.numCols; iCol++) {
        tmpSC[iCol][iDest] = this.SC[iCol][iRow];
      }
      iDest++;
		}
	}
  this.SC = tmpSC;
  this.persist();
}

function UCGetColNamesSerial(colDelim) {
  var serialCols = "";
  for (var iCol=0; iCol<this.numCols; iCol++) {
    if (iCol != 0) serialCols += colDelim;
    serialCols += this.colNames[iCol];
  }
  return serialCols;
}

function UCGetContentsSerial(colDelim, rowDelim) {
  var serialCart = "";
  for (var iRow=0; iRow<this.GetItemCount(); iRow++) {
    if (iRow != 0) serialCart += rowDelim
    for (var iCol=0; iCol<this.numCols; iCol++) {
      if (iCol != 0) serialCart += colDelim;
      serialCart += this.SC[iCol][iRow];
    }
  }
  return serialCart;
}

function UCSetContentsSerial(serialCart, colDelim, rowDelim) {
	var Rows = String(serialCart).split(rowDelim)
	for (iRow = 0; iRow < Rows.length; iRow++) {
		if (Rows[iRow] != "undefined" && Rows[iRow] != "") {
			Cols = Rows[iRow].split(colDelim)
			iCol = 0
			for (iCol = 0; iCol<Cols.length; iCol++) {
				this.SC[iCol][iRow] = Cols[iCol]
			}
		}
	}
	this.persist();
}

function SetCookie(){
	var cookieName = this.GetCookieName()
	var cookieStr = this.GetContentsSerial(this.cookieColDel, this.cookieRowDel)
	var cookieExp = GetCookieExp(this.cookieLifetime)
	Response.Cookies(cookieName) = cookieStr
	Response.Cookies(cookieName).expires = cookieExp
}

function GetCookieName(){
	var server = Request.ServerVariables("SERVER_NAME");
	return  server + this.Name;
}

function UCDestroyCookie(){
	cookieName = this.GetCookieName();
	Response.Cookies(cookieName) = ""
	Response.Cookies(cookieName).expires = "1/1/90"
}

function PopulateFromCookie(cookieStr){
  this.SetContentsSerial(cookieStr, this.cookieColDel, this.cookieRowDel)
}

// ***************** debug code ********************
function assert(bool, msg) {
	if (!bool) {
		Response.Write("<BR><BR>An error occured in the UltraDev shopping cart:<BR>" + msg + "<BR>");
		//Response.End();
	}
}

function AssertCartValid(colNames, msg) {
	// go through all cart data structures and insure consistency.
	// For example all column arrays should be the same length.
	// this function should be called often, especially just after
	// makeing changes to the data structures (adding, deleting, etc.)
	// also verify we always have the required columns:
	// ProductID, Quantity, Price, Total

	// the input arg is some I add as I code this package like
	// "Prior to return from AddToCart"
	//
	var ERR_COOKIE_SETTINGS = "Cookie settings on this page are inconsistent with those stored in the session cart<BR>"; 
	var ERR_BAD_NAME = "Cart name defined on this page is inconsistent with the cart name stored in the session<BR>";
	var ERR_COLUMN_COUNT = "The number of cart columns defined on this page is inconsistent with the cart stored in the session<BR>";
	var ERR_REQUIRED_COLUMNS = "Too few columns; minimum number of columns is 4<BR>";
	var ERR_REQUIRED_COLUMN_NAME = "Required Column is missing or at the wrong offset: ";
	var ERR_COLUMN_NAMES = "Cart column names defined on this page are inconsistent with the cart stored in the session";
	var ERR_INCONSISTENT_ARRAY_LENGTH = "Length of the arrays passed to cart constructor are inconsistent<BR>"
	var errMsg = "";
	var sessCart = Session(this.Name);

	if (sessCart != null) { // Validate inputs against session cart if it exists
		if (sessCart.Name != this.Name) errMsg += ERR_BAD_NAME;
		if (this.numCols < 4) errMsg += ERR_REQUIRED_COLUMNS;
		if (sessCart.numCols != this.numCols) errMsg += "Column Name Array: " + ERR_COLUMN_COUNT;
		if (sessCart.numCols != this.colComputed.length) errMsg += "Computed Column Array: " + ERR_COLUMN_COUNT;
		if (sessCart.bStoreCookie != this.bStoreCookie) errMsg += "Using Cookies: " + ERR_COOKIE_SETTINGS;
		if (sessCart.cookieLifetime != this.cookieLifetime) errMsg += "Cookie Lifetime: " + ERR_COOKIE_SETTINGS;

		// check that required columns are in the same place
		var productIndex = this.GetIndexOfColName(this.PRODUCTID);
		var quantityIndex = this.GetIndexOfColName(this.QUANTITY);
		var priceIndex = this.GetIndexOfColName(this.PRICE);
		var totalIndex = this.GetIndexOfColName(this.TOTAL);

		if (colNames[productIndex] != "ProductID") errMsg += ERR_REQUIRED_COLUMN_NAME + "ProductID<BR>";
		if (colNames[quantityIndex] != "Quantity") errMsg += ERR_REQUIRED_COLUMN_NAME + "Quantity<BR>";
		if (colNames[priceIndex] != "Price") errMsg += ERR_REQUIRED_COLUMN_NAME + "Price<BR>";
		if (colNames[totalIndex] != "Total") errMsg += ERR_REQUIRED_COLUMN_NAME + "Total<BR>";
	}
	else { // if cart doesn't exist in session, validate input array lengths and presence of reqiured columns
		if (this.numCols != this.colComputed.length) errMsg += ERR_INCONSISTENT_ARRAY_LENGTH;
		
		var bProductID = false, bQuantity = false, bPrice = false, bTotal = false;

		for (var j = 0; j < colNames.length; j++) {
			if (colNames[j] == "ProductID") bProductID = true;
			if (colNames[j] == "Quantity") bQuantity= true;
			if (colNames[j] == "Price") bPrice = true;
			if (colNames[j] == "Total") bTotal = true;
		}
		if (!bProductID) errMsg += ERR_REQUIRED_COLUMN_NAME + "ProductID<BR>";
		if (!bQuantity) errMsg += ERR_REQUIRED_COLUMN_NAME + "Quantity<BR>";
		if (!bPrice) errMsg += ERR_REQUIRED_COLUMN_NAME + "Price<BR>";
		if (!bTotal) errMsg += ERR_REQUIRED_COLUMN_NAME + "Total<BR>";
	}
	
	if (errMsg != "") {
		Response.Write(msg + "<BR>");
		Response.Write(errMsg + "<BR>");
		Response.End();
	}
}

function VBConstuctCart(Name, cookieLifetime, vbArrColNames, vbArrColComputed){
	var myObj;
	var a = new VBArray(vbArrColNames);
	var b = new VBArray(vbArrColComputed);
	eval("myObj = new UC_ShoppingCart(Name, cookieLifetime, a.toArray(), b.toArray())");
	return myObj;
}
</SCRIPT>
<SCRIPT LANGUAGE=vbscript runat=server NAME="UC_CART">
Function GetCookieExp(expDays) 
 	vDate = DateAdd("d", CInt(expDays), Now())
 	GetCookieExp = CStr(vDate)
End Function
</SCRIPT>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT NAME="UC_CART">
function DoNumber(str, nDigitsAfterDecimal, nLeadingDigit, nUseParensForNeg, nGroupDigits)
	DoNumber = FormatNumber(str, nDigitsAfterDecimal, nLeadingDigit, nUseParensForNeg, nGroupDigits)
End Function

function DoCurrency(str, nDigitsAfterDecimal, nLeadingDigit, nUseParensForNeg, nGroupDigits)
	DoCurrency = FormatCurrency(str, nDigitsAfterDecimal, nLeadingDigit, nUseParensForNeg, nGroupDigits)
End Function

function DoDateTime(str, nNamedFormat, nLCID)
	dim strRet
	dim nOldLCID

	strRet = str
	If (nLCID > -1) Then
		oldLCID = Session.LCID
	End If

	On Error Resume Next

	If (nLCID > -1) Then
		Session.LCID = nLCID
	End If

	If ((nLCID < 0) Or (Session.LCID = nLCID)) Then
		strRet = FormatDateTime(str, nNamedFormat)
	End If
										
	If (nLCID > -1) Then
		Session.LCID = oldLCID
	End If			
										
	DoDateTime = strRet
End Function

function DoPercent(str, nDigitsAfterDecimal, nLeadingDigit, nUseParensForNeg, nGroupDigits)
	DoPercent = FormatPercent(str, nDigitsAfterDecimal, nLeadingDigit, nUseParensForNeg, nGroupDigits)
End Function							

function DoTrim(str, side)
	dim strRet
	strRet = str

	If (side = "left") Then
		strRet = LTrim(str)
	ElseIf (side = "right") Then
		strRet = RTrim(str)
	Else
		strRet = Trim(str)
	End If
	DoTrim = strRet
End Function
</SCRIPT>
<%
UC_CartColNames=Array("ProductID","Quantity","ProductName","Price","TotalWeight","UnitWeight","Total")
UC_ComputedCols=Array("","","","","UnitWeight","","Price")
set UCCart1=VBConstuctCart("UCCart",2,UC_CartColNames,UC_ComputedCols)
UCCart1__i=0
%>
<%
'
'Update Stock and Sell  for each ProductID in the cart
'
set UpdateStock = Server.CreateObject("ADODB.Command")
UpdateStock.ActiveConnection = MM_conneshop_STRING
For UCCart1__i=0 To UCCart1.GetItemCount()-1
if (UCCart1.GetColumnValue("ProductID",UCCart1__i)) <> "" then
UpdateStock.CommandText = "UPDATE Products  SET Stock = Stock - " &(UCCart1.GetColumnValue("Quantity",UCCart1__i))& " ,Sell = Sell + " &(UCCart1.GetColumnValue("Quantity",UCCart1__i))& " WHERE ProductID = '" & (UCCart1.GetColumnValue("ProductID",UCCart1__i))& "'"
UpdateStock.CommandType = 1
UpdateStock.CommandTimeout = 0
UpdateStock.Prepared = true
UpdateStock.Execute()
end if
next
UpdateStock.ActiveConnection.Close
UCCart1.Destroy()
%>
<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="2" >
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" width="130"><img src="images/sitelogo.gif" height="88"></td>
    <td background="images/topback.gif" width="500" align="center" valign="middle"><a href="http://www.fans8.com" target="_blank"><img src="images/fans8.gif" width="468" height="60" border="0"></a> 
    </td>
    <td background="images/topback.gif" width="130"> <!-- #BeginLibraryItem "/Library/custm.lbi" --><table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td valign="middle" align="center"><a href="cart.asp"><img src="images/button_cart.gif" width="87" height="18" border="0"></a></td>
        </tr>
        <tr> 
          <td valign="middle" align="center"><a href="checkorder_login.asp"><img src="images/button_ddcx.gif" width="87" height="18" border="0"></a></td>
        </tr>
        <tr> 
          
    <td valign="middle" align="center"><a href="customer_register.asp"><img src="images/button_regist.gif" width="87" height="18" border="0"></a></td>
        </tr>
      </table><!-- #EndLibraryItem --></td>
  </tr>
</table>
<table width="760" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#000000">
  <tr> 
    <td bgcolor="#FF9900" height="22" valign="middle" align="center"> <!-- #BeginLibraryItem "/Library/nav.lbi" --><table width="80%" border="0" cellspacing="2" cellpadding="2">
          <tr align="center" valign="middle"> 
            <td><a href="newproduct.asp" class="white">新品快递</a></td>
            
    <td><a href="commend.asp" class="white">重点推荐</a></td>
            
    <td><a href="bestsell.asp" class="white">销售排行</a></td>
            
    <td><a href="bestprice.asp" class="white">特价商品</a></td>
          </tr>
        </table><!-- #EndLibraryItem --></td>
  </tr>
  <tr> 
    <td bgcolor="#FFCC66" height="22"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td>　<a href="default.asp" class="red">首页</a> &gt; 您的定单已成功提交</td>
          <td>&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td class="productName">
      <div align="center">
        <p>&nbsp;</p>
        <p>感谢您的惠顾！</p>
        <p>您的定单号为<%= Session("OrderID") %>，请牢记以供查询。</p>
        <p><a href="default.asp">返回商城</a></p>
        <p>&nbsp;</p>
      </div>
    </td>
  </tr>
</table>
<br>
<!-- #BeginLibraryItem "/Library/bottm.lbi" --><table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td background="images/topback.gif" align="center" height="16"><font color="#FFFFFF">copyright 
      2001 Powered by Peter.HJ</font></td>
  </tr>
</table><!-- #EndLibraryItem --></body>
</html>
<% Session.Abandon %>