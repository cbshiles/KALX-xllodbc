// xllconnect.cpp - connect to an ODBC driver
#include "xllodbc.h"

using namespace xll;

static AddInX xai_odbc_connect(
	FunctionX(XLL_HANDLEX, _T("?xll_odbc_connect"), _T("ODBC.CONNECT"))
	.Arg(XLL_CSTRINGX, _T("DSN"), _T("is the data source name."))
	.Arg(XLL_CSTRINGX, _T("_User"), _T("is the optional user name."))
	.Arg(XLL_CSTRINGX, _T("_Pass"), _T("is the optional password."))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Return a handle to an ODBC connection."))
);
HANDLEX WINAPI xll_odbc_connect(SQLCHAR* dsn, SQLCHAR* user, SQLCHAR* pass)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<ODBC::Dbc> hdbc(new ODBC::Dbc());

		ensure (SQL_SUCCEEDED(SQLConnect(*hdbc, dsn, SQL_NTS, user, SQL_NTS, pass, SQL_NTS)));

		h = hdbc.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}
/* not used
static AddInX xai_odbc_browse_connect(
	FunctionX(XLL_HANDLEX, _T("?xll_odbc_browse_connect"), _T("ODBC.CONNECT.BROWSE"))
	.Arg(XLL_CSTRINGX, _T("\"DRIVER={Microsoft Excel Driver (*.xls)}\""), _T("is an odbc data source name."))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Display a list of available drivers and return a handle to an ODBC connection."))
);
HANDLEX WINAPI xll_odbc_browse_connect(sqlcstr conn)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<ODBC::Dbc> ph(new ODBC::Dbc());

		while (SQL_NEED_DATA == ph->BrowseConnect(conn)) {
			// get networks
			// get servers
			// get databases
			// get tables
		}

		h = ph.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
		// call ODBC::DiagField::Get for more info
	}

	return h;
}
*/
static AddInX xai_odbc_driver_connect(
	FunctionX(XLL_HANDLEX, _T("?xll_odbc_driver_connect"), _T("ODBC.CONNECT.DRIVER"))
	.Arg(XLL_CSTRINGX, _T("{DSN=\"DEFAULT\"}"), _T("is an odbc connection string."))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Display a list of available drivers and return a handle to an ODBC connection."))
);
HANDLEX WINAPI xll_odbc_driver_connect(const SQLCHAR* conn)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<ODBC::Dbc> ph(new ODBC::Dbc());

		ensure (SQL_SUCCEEDED(ph->DriverConnect(conn)));

		h = ph.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

static AddInX xai_odbc_connection_string(
	FunctionX(XLL_CSTRING, _T("?xll_odbc_connection_string"), _T("ODBC.CONNECTION.STRING"))
	.Handle(_T("Dbc"), _T("is a handle returned by ODBC.CONNECT.*."))
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Returns the full connection string."))
);
const SQLCHAR* WINAPI xll_odbc_connection_string(HANDLEX dbc)
{
#pragma XLLEXPORT
	handle<ODBC::Dbc> hdbc(dbc);

	return hdbc ? hdbc->connectionString() : 0;
}
