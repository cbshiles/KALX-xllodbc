// stmt.cpp - Create a statement handle on a database connection.
#include "xllodbc.h"

using namespace xll;

static AddInX xai_sql_stmt(
	FunctionX(XLL_HANDLEX, _T("?xll_sql_stmt"), _T("ODBC.STATEMENT"))
	.Arg(XLL_HANDLEX, _T("Dbc"), _T("is a handle to a database connection. "))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to an ODBC statement object."))
	.Documentation(_T(""))
);
HANDLEX WINAPI xll_sql_stmt(HANDLEX dbc)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<SQL::Dbc> hdbc(dbc);
		ensure (hdbc);

		handle<SQL::Stmt> hstmt(new SQL::Stmt(*hdbc));

		h = hstmt.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}