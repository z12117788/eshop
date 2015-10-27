<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/conneshop.asp" -->
<%
' *** Edit Operations: declare variables

MM_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) <> "") Then

  MM_editConnection = MM_conneshop_STRING
  MM_editTable = "rating"
  MM_editRedirectUrl = ""
  MM_fieldsStr  = "radiobutton|value|hiddenField|value"
  MM_columnsStr = "RateLevel|none,none,NULL|ProductID|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(i+1) = CStr(Request.Form(MM_fields(i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: construct a sql insert statement and execute it

If (CStr(Request("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For i = LBound(MM_fields) To UBound(MM_fields) Step 2
    FormVal = MM_fields(i+1)
    MM_typeArray = Split(MM_columns(i+1),",")
    Delim = MM_typeArray(0)
    If (Delim = "none") Then Delim = ""
    AltVal = MM_typeArray(1)
    If (AltVal = "none") Then AltVal = ""
    EmptyVal = MM_typeArray(2)
    If (EmptyVal = "none") Then EmptyVal = ""
    If (FormVal = "") Then
      FormVal = EmptyVal
    Else
      If (AltVal <> "") Then
        FormVal = AltVal
      ElseIf (Delim = "'") Then  ' escape quotes
        FormVal = "'" & Replace(FormVal,"'","''") & "'"
      Else
        FormVal = Delim + FormVal + Delim
      End If
    End If
    If (i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End if
    MM_tableValues = MM_tableValues & MM_columns(i)
    MM_dbValues = MM_dbValues & FormVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

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

if(Request.QueryString("ProductID") <> "") then UpdateVisits__MMColParam = Request.QueryString("ProductID")

%>
<%
Dim CateName__MMColParam
CateName__MMColParam = "1"
if (Request.QueryString("ProductID") <> "") then CateName__MMColParam = Request.QueryString("ProductID")
%>
<%
set CateName = Server.CreateObject("ADODB.Recordset")
CateName.ActiveConnection = MM_conneshop_STRING
CateName.Source = "SELECT *  FROM Products INNER JOIN Categories ON  Products.CategoryID= Categories.CategoryID  WHERE ProductID = '" + Replace(CateName__MMColParam, "'", "''") + "'"
CateName.CursorType = 0
CateName.CursorLocation = 2
CateName.LockType = 3
CateName.Open()
CateName_numRows = 0
strproductid=CateName.Fields.Item("ProductID").Value
Session("ProductID") = strproductid
%>
<%
Dim SubCateName__MMColParam
SubCateName__MMColParam = "1"
if (Request.QueryString("ProductID") <> "") then SubCateName__MMColParam = Request.QueryString("ProductID")
%>
<%
set SubCateName = Server.CreateObject("ADODB.Recordset")
SubCateName.ActiveConnection = MM_conneshop_STRING
SubCateName.Source = "SELECT SubCategID ,SubCategoryName  FROM Products INNER JOIN SubCategories ON   Products.SubCategID= SubCategories.SubCategoryID  WHERE ProductID = '" + Replace(SubCateName__MMColParam, "'", "''") + "'"
SubCateName.CursorType = 0
SubCateName.CursorLocation = 2
SubCateName.LockType = 3
SubCateName.Open()
SubCateName_numRows = 0
%>
<%
Dim review__MMColParam
review__MMColParam = "1"
if (Request.QueryString("ProductID") <> "") then review__MMColParam = Request.QueryString("ProductID")
%>
<%
set review = Server.CreateObject("ADODB.Recordset")
review.ActiveConnection = MM_conneshop_STRING
review.Source = "SELECT * FROM review WHERE ProductID = '" + Replace(review__MMColParam, "'", "''") + "' ORDER BY ReviewTime DESC"
review.CursorType = 0
review.CursorLocation = 2
review.LockType = 3
review.Open()
review_numRows = 0
%>
<%
Dim rating__MMColParam
rating__MMColParam = "1"
if (Request.QueryString("ProductID") <> "") then rating__MMColParam = Request.QueryString("ProductID")
%>
<%
set rating = Server.CreateObject("ADODB.Recordset")
rating.ActiveConnection = MM_conneshop_STRING
rating.Source = "SELECT COUNT(ProductID) AS NUM_RATES, SUM(RateLevel) AS TOTAL_RATES, TOTAL_RATES/NUM_RATES AS AVERAGE_RATES,AVERAGE_RATES*14 AS GRAGHWIDTH,MAX(RateLevel) AS HIGH_RATE,MIN(RateLevel) AS LOW_RATE  FROM rating   WHERE ProductID = '" + Replace(rating__MMColParam, "'", "''") + "'"
rating.CursorType = 0
rating.CursorLocation = 2
rating.LockType = 3
rating.Open()
rating_numRows = 0
%>
<%

set UpdateVisits = Server.CreateObject("ADODB.Command")
UpdateVisits.ActiveConnection = MM_conneshop_STRING
UpdateVisits.CommandText = "UPDATE Products  SET Visits=Visits+1  WHERE ProductID='" + Replace(UpdateVisits__MMColParam, "'", "''") + "' "
UpdateVisits.CommandType = 1
UpdateVisits.CommandTimeout = 0
UpdateVisits.Prepared = true
UpdateVisits.Execute()

%>
<%
set ListValues = Server.CreateObject("ADODB.Recordset")
ListValues.ActiveConnection = MM_conneshop_STRING
ListValues.Source = "SELECT CategoryID, CategoryName FROM Categories"
ListValues.CursorType = 0
ListValues.CursorLocation = 2
ListValues.LockType = 3
ListValues.Open()
ListValues_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 5
Dim Repeat1__index
Repeat1__index = 0
review_numRows = review_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
review_total = review.RecordCount

' set the number of rows displayed on this page
If (review_numRows < 0) Then
  review_numRows = review_total
Elseif (review_numRows = 0) Then
  review_numRows = 1
End If

' set the first and last displayed record
review_first = 1
review_last  = review_first + review_numRows - 1

' if we have the correct record count, check the other stats
If (review_total <> -1) Then
  If (review_first > review_total) Then review_first = review_total
  If (review_last > review_total) Then review_last = review_total
  If (review_numRows > review_total) Then review_numRows = review_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (review_total = -1) Then

  ' count the total records by iterating through the recordset
  review_total=0
  While (Not review.EOF)
    review_total = review_total + 1
    review.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (review.CursorType > 0) Then
    review.MoveFirst
  Else
    review.Requery
  End If

  ' set the number of rows displayed on this page
  If (review_numRows < 0 Or review_numRows > review_total) Then
    review_numRows = review_total
  End If

  ' set the first and last displayed record
  review_first = 1
  review_last = review_first + review_numRows - 1
  If (review_first > review_total) Then review_first = review_total
  If (review_last > review_total) Then review_last = review_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = review
MM_rsCount   = review_total
MM_size      = review_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Add item to UC Shopping cart
UC_editAction = CStr(Request("URL"))
If (Request.QueryString <> "") Then
  UC_editAction = UC_editAction & "?" & Request.QueryString
End If
UC_recordId = CStr(Request.Form("UC_recordId"))
If (Request.Form("UC_recordId").Count = 1) Then
  set UC_rs=CateName
  UC_uniqueCol="ProductID"
  UC_redirectPage = ""
  If (NOT (UC_rs is Nothing)) Then
    ' Position recordset to correct location
    If (UC_rs.Fields.Item(UC_uniqueCol).Value <> UC_recordId) Then
      ' reset the cursor to the beginning
      If (UC_rs.CursorType > 0) Then
        If (Not UC_rs.BOF) Then UC_rs.MoveFirst
      Else
        UC_rs.Close
        UC_rs.Open
      End If
      Do While (Not UC_rs.EOF)
        If (Cstr(UC_rs.Fields.Item(UC_uniqueCol).Value) = UC_recordId) Then
          Exit Do
        End If
        UC_rs.MoveNext
      Loop
    End If
  End If
  UC_BindingTypes=Array("RS","LITERAL","RS","RS","NONE","RS","NONE")
  UC_BindingValues=Array("ProductID","1","ProductName","ListPrice","","UnitWeight","")
  UCCart1.AddItem UC_rs,UC_BindingTypes,UC_BindingValues,"increment"
  ' redirect with URL parameters
  If (UC_redirectPage <> "") Then
    If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      UC_redirectPage = UC_redirectPage & "?" & Request.QueryString
    End If
    Call Response.Redirect(UC_redirectPage)
  End If
End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
review_first = MM_offset + 1
review_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (review_first > MM_rsCount) Then review_first = MM_rsCount
  If (review_last > MM_rsCount) Then review_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>	
function DoTrimProperly(str, nNamedFormat, properly, pointed, points)
  dim strRet
  strRet = Server.HTMLEncode(str)
  strRet = replace(strRet, vbcrlf,"")
  strRet = replace(strRet, vbtab,"")
  If (LEN(strRet) > nNamedFormat) Then
    strRet = LEFT(strRet, nNamedFormat)			
    If (properly = 1) Then					
      Dim TempArray								
      TempArray = split(strRet, " ")	
      Dim n
      strRet = ""
      for n = 0 to Ubound(TempArray) - 1
        strRet = strRet & " " & TempArray(n)
      next
    End If
    If (pointed = 1) Then
      strRet = strRet & points
    End If
  End If
  DoTrimProperly = strRet
End Function
</SCRIPT>
<SCRIPT RUNAT=SERVER LANGUAGE=VBSCRIPT>										
function DoWhiteSpace(str)												
	DoWhiteSpace = Replace((Replace(str, vbCrlf, "<br>")), chr(32)&chr(32), "&nbsp;&nbsp;")			
End Function														
</SCRIPT>
<html>
<head>
<title>网上商城</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style/subcategory1.css" type="text/css">
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="2">
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
<form name="search_form" method="get" action="quick_search.asp">
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
            <td width="50%">　<a href="default.asp" class="red">首页</a> &gt; <A HREF="category.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "CategoryID=" & CateName.Fields.Item("CategoryID").Value %>" class="red"><%=(CateName.Fields.Item("CategoryName").Value)%></A> &gt; <A HREF="subcategory.asp?<%= MM_keepNone & MM_joinChar(MM_keepNone) & "SubCategoryID=" & SubCateName.Fields.Item("SubCategID").Value %>" class="red"><%=(SubCateName.Fields.Item("SubCategoryName").Value)%></A></td>
            <td width="50%" valign="middle" align="center"> 
              <select name="mnuCategory">
                <option value="<%=(CateName.Fields.Item("CategoryID").Value)%>" selected>在本类商城中</option>
                <%
While (NOT ListValues.EOF)
%>
                <option value="<%=(ListValues.Fields.Item("CategoryID").Value)%>" >在<%=(ListValues.Fields.Item("CategoryName").Value)%>中</option>
                <%
  ListValues.MoveNext()
Wend
If (ListValues.CursorType > 0) Then
  ListValues.MoveFirst
Else
  ListValues.Requery
End If
%>
              </select>
              <input type="text" name="textPname" size="20" maxlength="50">
              <input type="submit" name="Submit2" value="搜索">
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</form>
<table width="760" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="75%" valign="top" align="left"> 
      <table width="98%" border="0" cellspacing="2" cellpadding="2">
        <tr> 
          <td rowspan="4" align="left" valign="top" width="18%"><img src="images/product/<%=(CateName.Fields.Item("bImgUrl").Value)%>"></td>
          <td width="82%" class="productName"> 
            <form name="cartform" method="post" action="<%=UC_editAction%>">
              <%=(CateName.Fields.Item("ProductName").Value)%> 
              <% If CateName.Fields.Item("Price").Value <> (0) Then 'script %>
              <img src="images/hotprice.gif" width="24" height="24"> 
              <% End If ' end If CateName.Fields.Item("Price").Value <> (0) script %>
              <input type="image" border="0" name="imageField" src="images/addtocart.gif" width="30" height="18" alt="加入购物车">
              <input type="hidden" name="UC_recordId" value="<%= CateName.Fields.Item("ProductID").Value %>">
            </form>
          </td>
        </tr>
        <tr> 
          <td width="82%"><%=(CateName.Fields.Item("Supplier").Value)%>　<%=(CateName.Fields.Item("Author").Value)%></td>
        </tr>
        <tr> 
          <td width="82%">出版日期：<%=(CateName.Fields.Item("PubDate").Value)%>　<%=(CateName.Fields.Item("ProductID").Value)%></td>
        </tr>
        <tr> 
          <td width="82%"> 
            <% If CateName.Fields.Item("Price").Value <> (0) Then 'script %>
            原价：<span class="hotPrice"><%=(CateName.Fields.Item("Price").Value)%></span>元　现价： 
            <% End If ' end If CateName.Fields.Item("Price").Value <> (0) script %>
            <% If CateName.Fields.Item("Price").Value = (0) Then 'script %>
            价格： 
            <% End If ' end If CateName.Fields.Item("Price").Value = (0) script %>
            <%=(CateName.Fields.Item("ListPrice").Value)%>元</td>
        </tr>
        <tr> 
          <td colspan="2"><%= DoWhiteSpace(CateName.Fields.Item("Description").Value)%></td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="31%" valign="middle" align="center" height="22" class="bestselltitle">读者评论</td>
          <td width="69%">&nbsp;</td>
        </tr>
        <tr align="center" valign="top"> 
          <td colspan="2" class="bestsellbox"> 
            <table width="98%" border="0" cellspacing="2" cellpadding="2">
              <tr> 
                <td width="34%" height="17">&nbsp;</td>
                <td width="40%" height="17">&nbsp;</td>
                <td width="26%" height="17">&gt;&gt; <a href="#" onClick="MM_openBrWindow('review.asp','','scrollbars=yes,width=500,height=300')">我要评论</a> 
                  &lt;&lt;</td>
              </tr>
              <% 
While ((Repeat1__numRows <> 0) AND (NOT review.EOF)) 
%>
              <tr> 
                <td bgcolor="#E4E4E4"><a href="mailto:<%=(review.Fields.Item("ReviewEmail").Value)%>"><%=(review.Fields.Item("ReviewName").Value)%></a></td>
                <td colspan="2" align="center" bgcolor="#E4E4E4"><%=(review.Fields.Item("ReviewTime").Value)%></td>
              </tr>
              <tr> 
                <td colspan="3"><%= DoWhiteSpace(review.Fields.Item("ReviewContent").Value)%></td>
              </tr>
              <tr> 
                <td colspan="3" height="1" bgcolor="#E1E1E1"></td>
              </tr>
              <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  review.MoveNext()
Wend
%>
              <tr align="center"> 
                <td colspan="3"> 
                  <table border="0" width="50%">
                    <tr> 
                      <td width="23%" align="center"> 
                        <% If MM_offset <> 0 Then %>
                        <a href="<%=MM_moveFirst%>" class="navi">第一页</a> 
                        <% End If ' end MM_offset <> 0 %>
                      </td>
                      <td width="31%" align="center"> 
                        <% If MM_offset <> 0 Then %>
                        <a href="<%=MM_movePrev%>" class="navi">前一页</a> 
                        <% End If ' end MM_offset <> 0 %>
                      </td>
                      <td width="23%" align="center"> 
                        <% If Not MM_atTotal Then %>
                        <a href="<%=MM_moveNext%>" class="navi">下一页</a> 
                        <% End If ' end Not MM_atTotal %>
                      </td>
                      <td width="23%" align="center"> 
                        <% If Not MM_atTotal Then %>
                        <a href="<%=MM_moveLast%>" class="navi">最末页</a> 
                        <% End If ' end Not MM_atTotal %>
                      </td>
                    </tr>
                  </table>
                  <% If Not review.EOF Or Not review.BOF Then %>
                  第<%=(review_first)%> 到<%=(review_last)%>条评论，共<%=(review_total)%>条 
                  <% End If ' end Not review.EOF Or NOT review.BOF %>
                  <% If review.EOF And review.BOF Then %>
                  该商品目前没有任何评论。 
                  <% End If ' end review.EOF And review.BOF %>
                </td>
              </tr>
              <tr> 
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&gt;&gt; <a href="#" onClick="MM_openBrWindow('review.asp','','scrollbars=yes,width=500,height=300')">我要评论</a> 
                  &lt;&lt;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td width="25%" align="right" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="4" class="commendbox">
        <tr> 
          <td height="22" align="center" class="commendtitle" colspan="2">商品评分</td>
        </tr>
        <% If rating.Fields.Item("NUM_RATES").Value <> (0) Then 'script %>
        <tr> 
          <td colspan="2">以下是本商品得分情况：</td>
        </tr>
        <tr> 
          <td width="65%"> 
            <table width="70" border="0" cellspacing="0" cellpadding="0" background="images/rating/rate_back.gif" height="15">
              <tr> 
                <td> 
                  <table width="<%=(rating.Fields.Item("GRAGHWIDTH").Value)%>" border="0" cellspacing="0" cellpadding="0" height="15" background="images/rating/5.gif">
                    <tr> 
                      <td>&nbsp;</td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
          </td>
          <td width="35%"> 
            <% =(DoTrimProperly((rating.Fields.Item("AVERAGE_RATES").Value), 4, 0, 0, "")) %>
            分</td>
        </tr>
        <tr> 
          <td colspan="2">共有<%=(rating.Fields.Item("NUM_RATES").Value)%>位顾客为本商品评分</td>
        </tr>
        <tr> 
          <td colspan="2">最高分：<%=(rating.Fields.Item("HIGH_RATE").Value)%>分</td>
        </tr>
        <tr> 
          <td colspan="2">最低分：<%=(rating.Fields.Item("LOW_RATE").Value)%>分</td>
        </tr>
        <% End If ' end If rating.Fields.Item("NUM_RATES").Value <> (0) script %>
        <tr> 
          <% If rating.Fields.Item("NUM_RATES").Value = (0) Then 'script %>
          <td colspan="2">目前还没有顾客对此商品评分</td>
          <% End If ' end If rating.Fields.Item("NUM_RATES").Value = (0) script %>
        </tr>
        <tr align="center" bgcolor="#CCCCFF"> 
          <td colspan="2">&gt;&gt;&gt; 请您评分 &lt;&lt;&lt;</td>
        </tr>
        <tr align="center"> 
          <td colspan="2"> 
            <form name="rating_form" method="POST" action="<%=MM_editAction%>">
              <table width="90%" border="0" cellspacing="2" cellpadding="2">
                <tr> 
                  <td width="24%"> 
                    <input type="radio" name="radiobutton" value="5" checked>
                  </td>
                  <td width="76%"><img src="images/rating/5.gif" width="70" height="15"></td>
                </tr>
                <tr> 
                  <td width="24%"> 
                    <input type="radio" name="radiobutton" value="4">
                  </td>
                  <td width="76%"><img src="images/rating/4.gif" width="70" height="15"></td>
                </tr>
                <tr> 
                  <td width="24%"> 
                    <input type="radio" name="radiobutton" value="3">
                  </td>
                  <td width="76%"><img src="images/rating/3.gif" width="70" height="15"></td>
                </tr>
                <tr> 
                  <td width="24%"> 
                    <input type="radio" name="radiobutton" value="2">
                  </td>
                  <td width="76%"><img src="images/rating/2.gif" width="70" height="15"></td>
                </tr>
                <tr> 
                  <td width="24%"> 
                    <input type="radio" name="radiobutton" value="1">
                  </td>
                  <td width="76%"><img src="images/rating/1.gif" width="70" height="15"></td>
                </tr>
                <tr align="center"> 
                  <td colspan="2"> 
                    <input type="submit" name="Submit" value="评 分">
                    <input type="hidden" name="hiddenField" value="<%=(CateName.Fields.Item("ProductID").Value)%>">
                  </td>
                </tr>
              </table>
              <input type="hidden" name="MM_insert" value="true">
            </form>
          </td>
        </tr>
      </table>
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
<%
CateName.Close()
%>
<%
SubCateName.Close()
%>
<%
review.Close()
%>
<%
rating.Close()
%>
<%
ListValues.Close()
%>
