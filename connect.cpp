// connect.cpp - Connect to a ODBC data source.
#include "xllodbc.h"

using namespace xll;

// SQLConnect

// SQLBrowseConnect

static AddInX xai_sql_driver_connect(
	FunctionX(XLL_HANDLEX, _T("?xll_sql_driver_connect"), _T("ODBC.DRIVER.CONNECT"))
	.Arg(XLL_CSTRINGX, _T("DSN"), _T("is a Data Source Name. "))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a database connection from SQLDriveConnect."))
	.Documentation(
		_T("<codeInline>SQLDriverConnect</codeInline> is an alternative to <codeInline>SQLConnect</codeInline>. ")
		_T("It supports data sources that require more connection information ")
		_T("than the three arguments in <codeInline>SQLConnect</codeInline>, dialog boxes to prompt the ")
		_T("user for all connection information, and data sources that are not ")
		_T("defined in the system information.")
	)
);
HANDLEX WINAPI xll_sql_driver_connect(const SQLCHAR* dsn)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<SQL::Dbc> hdbc(new SQL::Dbc());

		ensure (SQL_SUCCEEDED(hdbc->DriverConnect((dsn))));

		h = hdbc.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}