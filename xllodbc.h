// xllodbc.h
#define EXCEL12
#include "../xll8/xll/xll.h"
#include "odbc.h"

typedef xll::traits<XLOPERX>::xchar xchar;
typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xword xword;
typedef xll::traits<XLOPERX>::xstring xstring;

#define ODBC_STR(o) reinterpret_cast<SQLTCHAR*>(o.val.str + 1), o.val.str[0]
#define ODBC_BUF0(o) ODBC_STR(o), 0
#define ODBC_BUFS(o) ODBC_STR(o), ODBC::lenptr<SQLSMALLINT>(o)
#define ODBC_BUFI(o) ODBC_STR(o), ODBC::lenptr<SQLINTEGER>(o)

template<SQLSMALLINT T>
inline bool ODBC_ERROR(ODBC::Handle<T>& h)
{
	ODBC::DiagRec<T> dr(h);
	std::basic_string<SQLTCHAR> rec;

	for (SQLSMALLINT i = 1; dr.Get(i) != SQL_NO_DATA; ++i) {
		rec.append(dr.state);
		rec.append((SQLTCHAR*)_T(": "));
		rec.append(dr.message);
		rec.append((SQLTCHAR*)_T("\n"));
	}

	if (rec.size())
		XLL_ERROR((char*)rec.c_str());

	return false;
}

namespace ODBC {

	template<class T>
	class lenptr {
		OPERX& o_;
		T len;
	public:
		lenptr(OPERX& o)
			: o_(o), len(0)
		{ }
		~lenptr()
		{
			if (len)
				o_.val.str[0] = static_cast<xchar>(len);
		}
		operator T*()
		{
			return &len;
		}
	};

	inline SQLSMALLINT NumResultsCols(ODBC::Stmt& stmt)
	{
		SQLSMALLINT n;

		ensure (SQL_SUCCEEDED(SQLNumResultCols(stmt, &n)) || ODBC_ERROR(stmt));

		return n;
	}
	inline SQLRETURN GetData(ODBC::Stmt& stmt, SQLUSMALLINT n, OPERX& o)
	{
		SQLRETURN rc(SQL_ERROR);

		switch (o.xltype) {
		case xltypeNum:
			rc = SQLGetData(stmt, n + 1, SQL_C_DOUBLE, &o.val.num, sizeof(double), 0);
			break;
		case xltypeStr:
			rc = SQLGetData(stmt, n + 1, SQL_C_TCHAR, ODBC_BUF0(o));
//		default:
//			throw std::runtime_error("ODBC::GetData: unknown type");
		}

		return rc;
	}


	struct Bind : public OPERX {
		typedef xll::traits<XLOPERX>::xword xword;
		Bind(ODBC::Stmt& stmt)
		{
			SQLSMALLINT n;
			ensure (SQL_SUCCEEDED(SQLNumResultCols(stmt, &n)));
			resize(1, n);

			for (xword i = 0; i < n; ++i) {
				OPERX& ri = operator[](i);
				OPERX name(_T(""), 255);
				SQLSMALLINT type(0), nullable(0), digits(0);
				SQLULEN len(0);
//				SQLColAttribute(stmt, n + 1, SQL_DESC_TYPE, ODBC_BUFS(name), &type);
				ensure (SQL_SUCCEEDED(SQLDescribeCol(stmt, i + 1, ODBC_BUFS(name), &type, &len, &digits, &nullable)) || ODBC_ERROR(stmt));
				switch (type) {
				case SQL_CHAR: case SQL_VARCHAR: case SQL_WCHAR: case SQL_WVARCHAR:
					ri = OPERX(_T(""), 255);
//					ensure (SQL_SUCCEEDED(SQLBindCol(stmt, n + 1, SQL_C_CHAR, ODBC_BUF0(ri))) || ODBC_ERROR(stmt));
					
					break;
				case SQL_SMALLINT: case SQL_INTEGER: case SQL_DOUBLE: case SQL_REAL:
				case SQL_BIT: case SQL_TINYINT: case SQL_BIGINT: case SQL_FLOAT:
					ri.xltype = xltypeNum;
//					ensure (SQL_SUCCEEDED(SQLBindCol(stmt, n + 1, SQL_C_DOUBLE, &ri.val.num, sizeof(double), 0)) || ODBC_ERROR(stmt));
					
					break;
				case SQL_DATE: case SQL_TIME: case SQL_TIMESTAMP:
				case SQL_TYPE_DATE: case SQL_TYPE_TIME: case SQL_TYPE_TIMESTAMP:
					ri = OPERX(_T(""), 255);
//					ensure (SQL_SUCCEEDED(SQLBindCol(stmt, n + 1, SQL_C_CHAR, ODBC_BUF0(ri))) || ODBC_ERROR(stmt));
					
					break;
//				default:
//					throw std::runtime_error("ODBC::Bind: unrecognized data type");
				}
			}
		}
		~Bind()
		{ }
	};

}