#include "../xll8/xll/xll.h"

#ifndef CATEGORY
#define CATEGORY _T("XLL")
#endif

using namespace xll;

static AddInX xai_range_set(
	FunctionX(XLL_HANDLEX, _T("?xll_range_set"), _T("RANGE.SET"))
	.Range(_T("Range"), _T("is a range."))
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(_T("Return a handle to Range."))
);
HANDLEX WINAPI xll_range_set(LPOPERX px)
{
#pragma XLLEXPORT
	return handle<OPERX>(new OPERX(*px)).get();
}

static AddInX xai_range_get(
	FunctionX(XLL_LPOPERX, _T("?xll_range_get"), _T("RANGE.GET"))
	.Handle(_T("Handle"), _T("is a handle to a range."))
	.Category(CATEGORY)
	.FunctionHelp(_T("Return the range correponding to Handle."))
);
LPOPERX WINAPI xll_range_get(HANDLEX h)
{
#pragma XLLEXPORT
	return handle<OPERX>(h, false).ptr();
}