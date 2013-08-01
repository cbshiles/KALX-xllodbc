// exec.cpp - Execute statements on a Stmt handle
#include "xllodbc.h"

using namespace xll;

static AddInX xai_sql_exec_direct(
	FunctionX(XLL_HANDLEX, _T("?xll_sql_exec_direct"), _T("ODBC.EXEC.DIRECT"))
	.Arg(XLL_HANDLEX, _T("Stmt"), _T("is a handle to a statement."))
	.Arg(XLL_CSTRINGX, _T("Query"), _T("is the query to execute. "))
	.Category(CATEGORY)
	.FunctionHelp(_T("Submit a SQL query for one-time execution."))
	.Documentation(_T(""))
);
HANDLEX WINAPI xll_sql_exec_direct(HANDLEX stmt, SQLCHAR* query)
{
#pragma XLLEXPORT
	handlex h;

	try {
		handle<SQL::Stmt> hstmt(stmt);
		ensure (hstmt);

		ensure (SQL_SUCCEEDED(SQLExecDirect(*hstmt, query, SQL_NTS)));

		h = stmt;
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}