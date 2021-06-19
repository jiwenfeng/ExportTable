#ifndef __EXCEL_H__
#define __EXCEL_H__

#include <map>
#include <string>
#include <functional>
#include <book.hpp>
#include <exceptions.hpp>
#include "CProgressBar.h"
#include <compoundfile/compoundfile_exceptions.hpp>

class CExcel
{
public:
	CExcel();

	~CExcel();

private:
	BOOL CheckSheet(Excel::Sheet* sheet);

	BOOL CheckHeader(Excel::Sheet* sheet, int row, BOOL isServerHeader);

	BOOL CheckData(Excel::Sheet* sheet, int row);

	BOOL CheckDataFormat(Excel::Sheet* sheet, const CString& type, const CString &data);
	
private:
	BOOL IsInt(const CString& data);

	BOOL IsFloat(const CString& data);

	BOOL IsString(const CString& data);

	BOOL IsMap(const CString& data);

	BOOL IsMacros(const CString& data);

	BOOL MapIsValid(const CString& data);

	BOOL MapDataIsValid(const CString& data);

	BOOL MapValueIsValid(const CString& data);

	BOOL CheckMapDataFormat(const CString& data);

	BOOL MapKeyIsValid(const CString& data);

	BOOL IsArray(const CString& data);

	BOOL ArrayIsValid(const CString& data);

	BOOL IsDigit(const CString& data);


private:
	std::string CString2String(CString from);
public:
	BOOL Check(const CString &file, int id);

	CString ErrorMsg();

private:
	CString m_err;
	std::map<int, CString> m_colInfo;
	CString m_types;
	CString m_file;
	CString m_sheet;
	std::map <CString, std::function<BOOL (const CString& data)> > m_checkFuncList;
};

#endif