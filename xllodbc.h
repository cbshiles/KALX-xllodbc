// header.h - header file for project
// Uncomment the following line to use features for Excel2007 and above.
//#define EXCEL12
#pragma once
#include <memory>
#include "xll/xll.h"
#include "odbc.h"

#ifndef CATEGORY
#define CATEGORY _T("ODBC")
#endif

typedef xll::traits<XLOPERX>::xcstr xcstr; // pointer to const string
typedef xll::traits<XLOPERX>::xword xword; // use for OPER and FP indices

namespace SQL {
	// ab0cde00 -> ab;cde00
	inline void null2char(LPSTR s, SQLCHAR c = ';')
	{
		for (LPSTR p = s; !(p[0] == 0 && p[1] == 0); ++p)
			if (!*p)
				*p = c;
	}

	// 1x2 range of description and attributes
	class Drivers : public net::enumerator<OPERX> {
		static const SQLSMALLINT DESC_MAX = 1024, ATTR_MAX = 1024;
		OPERX da_;
		SQLSMALLINT ndesc_, nattr_;
		SQLUSMALLINT dir_;
		SQLRETURN Drivers_(SQLUSMALLINT direction)
		{
			SQLRETURN rc;
			SQLSMALLINT ndesc, nattr;

			rc = SQLDrivers(Env(), SQL_FETCH_FIRST,
				(SQLCHAR*)(da_[0].val.str + 1), ndesc_, &ndesc, 
				(SQLCHAR*)(da_[1].val.str + 1), nattr_, &nattr);

			da_[0].val.str[0] = static_cast<SQLCHAR>(ndesc);
			da_[1].val.str[0] = static_cast<SQLCHAR>(nattr);

			null2char(da_[1].val.str + 1);

			LPSTR s;
			s = da_[0].val.str+1;
			s = da_[1].val.str+1;

			return rc;
		}
	public:
		Drivers(SQLSMALLINT ndesc = DESC_MAX, SQLSMALLINT nattr = ATTR_MAX)
			: da_(1,2), ndesc_(ndesc), nattr_(nattr), dir_(SQL_FETCH_FIRST)
		{
			da_[0] = OPERX(_T(""), static_cast<SQLCHAR>(ndesc));
			da_[1] = OPERX(_T(""), static_cast<SQLCHAR>(nattr));
		}
		~Drivers()
		{ }
		void _reset()
		{
			Drivers_(SQL_FETCH_FIRST);
		}
		bool _next()
		{
			SQLRETURN rc = Drivers_(dir_);
			dir_ = SQL_FETCH_NEXT;

			return SQL_NO_DATA != rc;
			
		}
		const OPERX& _current() const
		{
			return da_;
		}
	};
	class Driver : public net::enumerable<OPERX> {
		Drivers d_;
	public:
		Drivers& _get()
		{
			return d_;
		}
	};

	class Results : public net::enumerator<OPERX> {
		SQL::Stmt& stmt_;
		OPERX row_; // struct { char buf; OPERX row; } pass &buf as last arg to SQLBindCol
		void Bind(xword i, OPERX& o)
		{
			SQLSMALLINT type;

			ensure (SQL_SUCCEEDED(SQLColAttribute(stmt_, i + 1, SQL_DESC_TYPE, 0, 0, 0, &type)));

			switch (type) {
			// convert DATE and TIME to Excel Julian at database query level
			case SQL_DOUBLE:
				SQLBindCol(stmt_, i + 1, type, &o.val.num, 0, 0);
				break;
			case SQL_VARCHAR: case SQL_WVARCHAR:
				o = OPERX(_T(""), 255);
				SQLBindCol(stmt_, i + 1, type, o.val.str + 1, 255, (SQLINTEGER*)(o.val.str - 1));
				break;
			}
		}
		Results(const Results&);
		Results& operator=(const Results&);
	public:
		Results(SQL::Stmt& stmt)
			: stmt_(stmt)
		{
			SQLSMALLINT cols;

			ensure (SQL_SUCCEEDED(SQLNumResultCols(stmt, &cols)));
			row_.resize(1, static_cast<xword>(cols));

			for (xword i = 0; i < row_.size(); ++i)
				Bind(i, row_[i]);
		}
		~Results()
		{
			SQLCancel(stmt_);
		}
		void _reset()
		{
			SQLSetPos(stmt_, 0, SQL_REFRESH, SQL_LOCK_NO_CHANGE); // ???
		}
		bool _next()
		{
			SQLRETURN rc = SQLFetch(stmt_);
			// fix up row?
			return SQL_SUCCEEDED(rc);
		}
		const OPERX& _current()
		{
			return row_;
		}
	};
}