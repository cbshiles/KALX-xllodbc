// xllodbc.h
//#define EXCEL12
#include "../xll8/xll/xll.h"
#include "odbc.h"

typedef xll::traits<XLOPERX>::xchar xchar;
typedef xll::traits<XLOPERX>::xcstr xcstr;
typedef xll::traits<XLOPERX>::xword xword;
typedef const SQLCHAR* sqlcstr; 
typedef SQLCHAR* sqlstr; 

// "parse\0null\0terminated\0sequence\0\0
inline OPERX split0(xcstr s)
{
	OPERX o;

	for (xcstr b = s, e = _tcschr(s, 0); e[1]; b = e + 1, e = _tcschr(b, 0)) {
		o.push_back(OPERX(b, e - b));
	}

	return o;
}
// "split,strings;;on,sep" -> ["split","strings","","on","sep"]
inline OPERX split(xcstr s, xword ns = 0, xcstr sep = _T(",;"))
{
	OPERX o;

	if (!ns)
		ns = _tcslen(s);

	xcstr b = s, e;
	do {
		e = _tcspbrk(b, sep);
		xchar n = e ? static_cast<xchar>(e - b) : ns;
		o.push_back(OPERX(b, n));
		ns -= n + 1;
		b = e + 1;
	} while (e);

	return o;
}

#define ODBC_STR(o) reinterpret_cast<SQLCHAR*>(o.val.str + 1), o.val.str[0]
#define ODBC_BUF0(o) ODBC_STR(o), 0
#define ODBC_BUFS(o) ODBC_STR(o), ODBC::lenptr<SQLSMALLINT>(o)
#define ODBC_BUFI(o) ODBC_STR(o), ODBC::lenptr<SQLINTEGER>(o)

template<SQLSMALLINT T>
inline void ODBC_ERROR(ODBC::Handle<T>& h)
{
	ODBC::DiagRec<T> dr(h);
	for (SQLSMALLINT i = 1; dr.Get(i) != SQL_NO_DATA; ++i) {
		std::basic_string<SQLCHAR> rec(dr.state);
		rec.append((SQLCHAR*)_T(": "));
		rec.append(dr.message);

		XLL_INFO((char*)rec.c_str());
	}
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

	struct Row : public OPERX {
		typedef xll::traits<XLOPERX>::xword xword;
		Row(ODBC::Stmt& stmt)
		{
			SQLSMALLINT n;
			ensure (SQL_SUCCEEDED(SQLNumResultCols(stmt, &n)));
			resize(1, n);

			SQLSMALLINT type, nullable, digits;
			SQLULEN len;
			for (xword i = 0; i < n; ++i) {
				OPERX& val = operator[](i);
				OPERX name(_T(""), 255);
				ensure (SQL_SUCCEEDED(SQLDescribeCol(stmt, i + 1, ODBC_BUFS(name), &type, &len, &digits, &nullable)));
				switch (type) {
				case SQL_CHAR: case SQL_VARCHAR: {
					val = OPERX(_T(""), 255);
					ensure (SQLBindCol(stmt, n + 1, SQL_C_CHAR, 
						reinterpret_cast<SQLCHAR*>(val.val.str + 1), val.val.str[0], 0));
					break;
				}
				case SQL_SMALLINT: case SQL_INTEGER: case SQL_REAL: case SQL_FLOAT:
				case SQL_BIT: case SQL_TINYINT: case SQL_BIGINT:
					val.xltype = xltypeNum;
					ensure (SQLBindCol(stmt, n + 1, SQL_C_DOUBLE, &val.val.num, 0, 0));
					break;
				default:
					throw std::runtime_error("ODBC::Row: unrecognized data type");
				}
			}
		}
		~Row()
		{ }
	};

}