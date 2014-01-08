// xlldrivers.cpp - list of ODBC drivers
#include "xllodbc.h"

using namespace xll;

static AddInX xai_odbc_drivers(
	Function(XLL_LPOPERX, _T("?xll_odbc_drivers"), _T("ODBC.DRIVERS"))
	.Uncalced()
	.Category(_T("ODBC"))
	.FunctionHelp(_T("Retrieve a range of ODBC drivers."))
);
LPOPERX WINAPI xll_odbc_drivers(void)
{
#pragma XLLEXPORT
	static OPERX o;

	try {
		o.resize(0,0);
		OPERX r(1,2);
		r[0] = OPERX(_T(""), 255);
		r[1] = OPERX(_T(""), 255);

		while (SQL_NO_DATA != SQLDrivers(ODBC::Env(), SQL_FETCH_NEXT, ODBC_BUFS(r[0]), ODBC_BUFS(r[1]))) {
			OPERX kv(split0(r[1].val.str + 1));
			handle<OPERX> hkv(new OPERX());
			for (xword i = 0; i < kv.size(); ++i) {
				hkv->push_back(split(kv[i].val.str + 1, kv[i].val.str[0], _T("=")).resize(1,2));
			}

			o.push_back(r[0]);
			o.push_back(OPERX(hkv.get()));

			r[0].val.str[0] = 255;
			r[1].val.str[0] = 255;
		}

		o.resize(o.size()/2, 2);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		o = OPERX(xlerr::NA);
	}

	return &o;
}

