// function.cpp - Rename this file and replace this description.
#include "xllodbc.h"

using namespace xll;

#ifdef _DEBUG

int xll_test_odbc(void)
{
	try {
		SQL::Dbc dbc;
		dbc.DriverConnect(0);

		SQL::Stmt stmt(dbc);
//		SQLExecDirect(stmt, txt, len);

	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_odbc(xll_test_odbc);

#endif // _DEBUG