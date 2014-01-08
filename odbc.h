// odbc.h - platform independent ODBC code
#include "../xll8/xll/ensure.h"
#include <Windows.h>
#include <sqlext.h>

// Adapt OPER's to ODBC buffers.
//#define ODBC_BUF(o) reinterpret_cast<SQLCHAR*>(o.val.str + 1), o.val.str[0], ODBC::buf<SQLSMALLINT>(o).ptr()
//#define ODBC_BUF2(o) reinterpret_cast<SQLCHAR*>(o.val.str + 1), o.val.str[0], ODBC::buf<SQLLEN>(o).ptr()

//#define ODBC_STR(o) reinterpret_cast<SQLCHAR*>(o.val.str + 1), o.val.str[0], ODBC::buf<SQLSMALLINT>(o).ptr()
//#define ODBC_NUM(o) reinterpret_cast<SQLPOINTER>(o.val.num), 0, 0

//#define ODBC_BUF(o) ODBC::ptr(o), ODBC::len(o), ODBC::lenptr(o)

namespace ODBC {
/*
	template<class T>
	struct ptrlen_traits {
		typedef typename T type;
	};

	template<class T>
	inline SQLPOINTER ptr(T& t) { reutrn &t; }
	template<class T>
	inline T len(T&) { return sizeof(T); }
	template<class T>
	class ptrlen {
		ptrlen_traits<T>::type len;
	public:
		ptrlen(T&) { }
		~ptrlen() { }
		operator T*() { return 0; }
	};
	template<>
	class ptrlen<SQLCHAR*> {
		SQLCHAR* t_;
	public:
		ptrlen(SQLCHAR* t)
			: t_(t)
		{ }
		~ptrlen()
		{
			t_[0] = static_cast<SQLCHAR>(len_);
		}
		
	}

	// 
	template<class T, class L>
	class Buf {
		T& type_;
		L len_;
	public:
		Buf(T& type)
			: type_(type), len_(0)
		{ }
		~Buf()
		{ }
		T* ptr()
		{
			return &type;
		}
		L len()
		{
			return sizeof(T);
		}
		L* lenptr()
		{
			return 0;
		}
	};
	// int classes
	template<class T>
	class Buf<T,T> { };
	
	// preallocated counted strings
	template<class L>
	class Buf<SQLCHAR*,L> {
		~Buf()
		{
			if (len_ < type[0])
				type[0] = static_cast<SQLCHAR>(len_);
		}
		operator SQLCHAR*()
		{
			return type + 1;
		}
		L len()
		{
			return type[0];
		}
		L* lenptr()
		{
			return &len_;
		}
	};
*/
	template<SQLSMALLINT T>
	struct Handle {
		static const SQLSMALLINT type = T;
		SQLRETURN rc;
		SQLHANDLE h;
		Handle(const SQLHANDLE& _h = SQL_NULL_HANDLE)
			: rc(SQLAllocHandle(T, _h, &h))
		{ 
		}
		Handle(const Handle&) = delete;
		Handle& operator=(const Handle&) = delete;
		~Handle()
		{
			if (h)
				SQLFreeHandle(T, h);
		}
		operator SQLHANDLE&()
		{
			return h;
		}
		operator const SQLHANDLE&() const
		{
			return h;
		}
	};

	template<SQLSMALLINT T>
	class DiagRec {
		const Handle<T>& h_;
	public:
		SQLCHAR state[6], message[SQL_MAX_MESSAGE_LENGTH]; 
		SQLINTEGER error;
		SQLRETURN rc;
		DiagRec(const Handle<T>& h)
			: h_(h)
		{ }
		DiagRec(const DiagRec&) = delete;
		DiagRec& operator=(const DiagRec&) = delete;
		~DiagRec()
		{ }
		SQLRETURN Get(SQLSMALLINT n)
		{
			SQLSMALLINT len;

			return rc = SQLGetDiagRec(T, h_, n, state, &error, message, SQL_MAX_MESSAGE_LENGTH, &len);
		}
	};

/*	// proxy class for L*
	template<class L, class T>
	struct LenPtr {
		L& len;
		LenPtr(L& l, T& t)
			: len(l), type(t)
		{ }
		~LenPtr()
		operator L*()
		{
			return &len;
		}
	};
	template<class L>
	struct LenPtr<L,SQLCHAR*> {
		~LenPtr()
		{
			type[0] = static_cast<SQLCHAR>(len);
		}
	};
*/
	template<class R>
	struct field_traits {
		static const SQLSMALLINT len;
	};
#define ODBC_FIELD_TRAITS(t) template<> struct field_traits<SQL ## t > { static const SQL ## t len = SQL_IS_ ## t ; };
	ODBC_FIELD_TRAITS(SMALLINT)
	ODBC_FIELD_TRAITS(USMALLINT)
	ODBC_FIELD_TRAITS(INTEGER)
	ODBC_FIELD_TRAITS(UINTEGER)

	template<SQLSMALLINT T>
	class DiagField {
		const Handle<T>& h_;
	public:
		DiagField(const Handle<T>& h)
			: h_(h)
		{ }
		~DiagField()
		{ }
		template<class R>
		R Get(SQLSMALLINT n, SQLSMALLINT id)
		{
			R r;

			ensure (SQL_SUCCEEDED(SQLGetDiagField(Handle<T>::type, h_, n, id, &r, field_traits<R>::len(r), 0)));

			return r;
		}
		std::basic_string<SQLCHAR> Get(SQLSMALLINT n, SQLSMALLINT id)
		{
			std::basic_string<SQLCHAR> r(255);

			SQLSMALLINT len;
			ensure (SQL_SUCCEEDED(SQLGetDiagField(Handle<T>::type, h_, n, id, &r[0], r.size(), &len)));
			if (len < r.size())
				r.resize(len);

			return r;
		}
		void Get(SQLSMALLINT n, SQLSMALLINT id, SQLCHAR* str)
		{
			SQLSMALLINT len;
			ensure (SQL_SUCCEEDED(SQLGetDiagField(Handle<T>::type, h_, n, id, str + 1, str[0], &len)));
			if (len < r.size())
				str[0] = static_cast<SQLCHAR>(len);
		}
	};

	// Environment singleton
	class Env {
		static SQLHANDLE env_(void)
		{
			static bool first(true);
			static Handle<SQL_HANDLE_ENV> h;

			if (first) {
				ensure (SQL_SUCCEEDED(SQLSetEnvAttr(h.h, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0)));
				first = false;
			}

			return h.h;
		}
	public:
		operator SQLHANDLE()
		{
			return env_();
		}
	};

	class Dbc : public Handle<SQL_HANDLE_DBC> {
		SQLCHAR connect_[1024];
		Dbc(const Dbc&);
		Dbc& operator=(const Dbc&);
	public:
		Dbc()
			: Handle<SQL_HANDLE_DBC>(Env())
		{ }
#ifdef _WINDOWS
		SQLRETURN DriverConnect(const SQLCHAR* connect, SQLUSMALLINT complete = SQL_DRIVER_COMPLETE)
		{
			return rc = SQLDriverConnect(*this, GetDesktopWindow(), const_cast<SQLCHAR*>(connect), SQL_NTS, connect_, 1024, 0, complete);
		}
#endif
		SQLRETURN BrowseConnect(const SQLCHAR* connect)
		{
			return rc = SQLBrowseConnect(*this, const_cast<SQLCHAR*>(connect), SQL_NTS, connect_, 1024, 0);
		}
		const SQLCHAR* connectionString() const
		{
			return connect_;
		}
		SQLSMALLINT GetInfo(SQLSMALLINT type)
		{
			SQLSMALLINT si;

			ensure (SQL_SUCCEEDED(SQLGetInfo(*this, type, &si, SQL_IS_SMALLINT, 0)));

			return si;
		}

	};

	class Stmt : public Handle<SQL_HANDLE_STMT> {
		Stmt(const Stmt&);
		Stmt& operator=(const Stmt&);
	public:
		Stmt(const Dbc& dbc)
			: Handle<SQL_HANDLE_STMT>(dbc)
		{ }
		~Stmt()
		{
			SQLFreeStmt(*this, SQL_CLOSE);
		}
	};

}