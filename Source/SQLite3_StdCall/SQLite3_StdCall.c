// SQLite3_StdCall.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "SQLite3_StdCall.h"

SQLITE3_STDCALL_API const char* __stdcall sqlite3_stdcall_libversion(void)
{
	return sqlite3_libversion();
}

SQLITE3_STDCALL_API const char * __stdcall sqlite3_stdcall_errmsg(sqlite3 *pDb)
{
	return sqlite3_errmsg(pDb);
}

SQLITE3_STDCALL_API const void * __stdcall sqlite3_stdcall_errmsg16(sqlite3 *pDb)
{
	return sqlite3_errmsg16(pDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_errcode(sqlite3 *pDb)
{
	return sqlite3_errcode(pDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_extended_errcode(sqlite3 *pDb)
{
	return sqlite3_extended_errcode(pDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_open(
  const char *filename,   /* Database filename (UTF-8) */
  sqlite3 **ppDb          /* OUT: SQLite db handle */
)
{
	return sqlite3_open(filename, ppDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_open16(const void *filename, sqlite3 **ppDb)
{
	return sqlite3_open16(filename, ppDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_open_v2(const char *filename, sqlite3 **ppDb, int flags, const char *zVfs)
{
	return sqlite3_open_v2(filename, ppDb, flags, zVfs);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_close(sqlite3 *pDb)
{
	return sqlite3_close(pDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_changes(sqlite3 *pDb)
{
	return sqlite3_changes(pDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_total_changes(sqlite3 *pDb)
{
	return sqlite3_total_changes(pDb);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_prepare_v2(
  sqlite3 *db,            /* Database handle */
  const char *zSql,       /* SQL statement, UTF-8 encoded */
  int nByte,              /* Maximum length of zSql in bytes. */
  sqlite3_stmt **ppStmt,  /* OUT: Statement handle */
  const char **pzTail     /* OUT: Pointer to unused portion of zSql */
)
{
	return sqlite3_prepare_v2(db, zSql, nByte, ppStmt, pzTail);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_prepare16_v2(
  sqlite3 *db,            /* Database handle */
  const void *zSql,       /* SQL statement, UTF-16 encoded */
  int nByte,              /* Maximum length of zSql in bytes. */
  sqlite3_stmt **ppStmt,  /* OUT: Statement handle */
  const void **pzTail     /* OUT: Pointer to unused portion of zSql */
)
{
	return sqlite3_prepare16_v2(db, zSql, nByte, ppStmt, pzTail);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_finalize(sqlite3_stmt *pStmt)
{
	return sqlite3_finalize(pStmt);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_step(sqlite3_stmt *pStmt)
{
	return sqlite3_step(pStmt);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_reset(sqlite3_stmt *pStmt)
{
	return sqlite3_reset(pStmt);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_clear_bindings(sqlite3_stmt *pStmt)
{
	return sqlite3_clear_bindings(pStmt);
}

SQLITE3_STDCALL_API const void * __stdcall sqlite3_stdcall_column_blob(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_blob(pStmt, iCol);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_column_bytes(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_bytes(pStmt, iCol);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_column_bytes16(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_bytes16(pStmt, iCol);
}

SQLITE3_STDCALL_API double __stdcall sqlite3_stdcall_column_double(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_double(pStmt, iCol);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_column_int(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_int(pStmt, iCol);
}

SQLITE3_STDCALL_API sqlite3_int64 __stdcall sqlite3_stdcall_column_int64(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_int64(pStmt, iCol);
}

SQLITE3_STDCALL_API const unsigned char * __stdcall sqlite3_stdcall_column_text(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_text(pStmt, iCol);
}

SQLITE3_STDCALL_API const void * __stdcall sqlite3_stdcall_column_text16(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_text16(pStmt, iCol);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_column_type(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_type(pStmt, iCol);
}

SQLITE3_STDCALL_API sqlite3_value * __stdcall sqlite3_stdcall_column_value(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_value(pStmt, iCol);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_column_count(sqlite3_stmt* pStmt)
{
	return sqlite3_column_count(pStmt);
}

SQLITE3_STDCALL_API const unsigned char * __stdcall sqlite3_stdcall_column_name(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_name(pStmt, iCol);
}

SQLITE3_STDCALL_API const void * __stdcall sqlite3_stdcall_column_name16(sqlite3_stmt* pStmt, int iCol)
{
	return sqlite3_column_name16(pStmt, iCol);
}


SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_blob(sqlite3_stmt* pStmt, int paramIndex, const void* pValue, int nBytes, void(*pfDelete)(void*))
{
	return sqlite3_bind_blob(pStmt, paramIndex, pValue, nBytes, pfDelete);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_double(sqlite3_stmt* pStmt, int paramIndex, double value)
{
	return sqlite3_bind_double(pStmt, paramIndex, value);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_int(sqlite3_stmt* pStmt, int paramIndex, int value)
{
	return sqlite3_bind_int(pStmt, paramIndex, value);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_int64(sqlite3_stmt* pStmt, int paramIndex, sqlite3_int64 value)
{
	return sqlite3_bind_int64(pStmt, paramIndex, value);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_null(sqlite3_stmt* pStmt, int paramIndex)
{
	return sqlite3_bind_null(pStmt, paramIndex);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_text(sqlite3_stmt* pStmt, int paramIndex, const char* zValue, int nBytes, void(*pfDelete)(void*))
{
	return sqlite3_bind_text(pStmt, paramIndex, zValue, nBytes, pfDelete);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_text16(sqlite3_stmt* pStmt, int paramIndex, const void* zValue, int nBytes, void(*pfDelete)(void*))
{
	return sqlite3_bind_text16(pStmt, paramIndex, zValue, nBytes, pfDelete);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_value(sqlite3_stmt* pStmt, int paramIndex, const sqlite3_value* pValue)
{
	return sqlite3_bind_value(pStmt, paramIndex, pValue);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_zeroblob(sqlite3_stmt* pStmt, int paramIndex, int nBytes)
{
	return sqlite3_bind_zeroblob(pStmt, paramIndex, nBytes);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_parameter_count(sqlite3_stmt* pStmt)
{
	return sqlite3_bind_parameter_count(pStmt);
}

SQLITE3_STDCALL_API const char * __stdcall sqlite3_stdcall_bind_parameter_name(sqlite3_stmt* pStmt, int paramIndex)
{
	return sqlite3_bind_parameter_name(pStmt, paramIndex);
}

SQLITE3_STDCALL_API int __stdcall sqlite3_stdcall_bind_parameter_index(sqlite3_stmt* pStmt, const char *zName)
{
	return sqlite3_bind_parameter_index(pStmt, zName);
}

