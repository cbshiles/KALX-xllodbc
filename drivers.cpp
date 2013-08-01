// drivers.cpp - SQLDrivers
#include "xllodbc.h"

using namespace xll;
using SQL::sqlstring;

static AddInX xai_sql_drivers(
	FunctionX(XLL_LPOPERX, _T("?xll_sql_drivers"), _T("ODBC.DRIVERS"))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a two column array of available SQL drivers and their attributes."))
	.Documentation()
);
LPOPERX WINAPI xll_sql_drivers()
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o.resize(0,0);

		SQL::Driver d;
		net::enumerator<OPERX>& e(d.get());

		while (e.next()) {
			o.push_back(e.current());
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = ErrX(xlerrNA);
	}

	return &o;
}