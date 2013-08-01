// odbc.h - platform independent ODBC wrapper
// char needs to be unsigned (/J)
#pragma once
#include <windows.h>
#include <sqlext.h>
#include "xll/utility/enumerable.h"

namespace SQL {

	typedef std::basic_string<SQLCHAR> sqlstring;

	template<SQLSMALLINT type>
	class Handle {
		SQLRETURN rc_;
		SQLHANDLE h_;
		Handle(const Handle&);
		Handle& operator=(const Handle&);
	public:
		Handle(const SQLHANDLE& h)
			: rc_(SQLAllocHandle(type, h, &h_))
		{
			ensure (SQL_SUCCEEDED(rc_));
		}
		~Handle(void)
		{
			SQLFreeHandle(type, h_);
		}
		operator const SQLHANDLE() const
		{
			return h_;
		}
		SQLHANDLE* operator&(void)
		{
			return &h_;
		}
	};

	// Environment singleton
	class Env {
		static SQLHANDLE env_(void)
		{
			static SQLHANDLE h_(SQL_NULL_HANDLE);

			if (h_ == SQL_NULL_HANDLE) {
				ensure (SQL_SUCCEEDED(SQLAllocHandle(SQL_HANDLE_ENV, h_, &h_)));
				ensure (SQL_SUCCEEDED(SQLSetEnvAttr(h_, SQL_ATTR_ODBC_VERSION, (SQLPOINTER)SQL_OV_ODBC3, 0)));
			}

			return h_;
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
		SQLRETURN DriverConnect(const SQLCHAR* connect, SQLUSMALLINT complete = SQL_DRIVER_COMPLETE)
		{
			return SQLDriverConnect(*this, GetDesktopWindow(), const_cast<SQLCHAR*>(connect), SQL_NTS, connect_, 1024, 0, complete);
		}
		const SQLCHAR* connectionString() const
		{
			return connect_;
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

