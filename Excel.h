#ifndef __EXCEL_H__
#define __EXCEL_H__

#include <map>
#include <string>
#include <vector>
#include <functional>
#include <book.hpp>
#include <exceptions.hpp>
#include <compoundfile/compoundfile_exceptions.hpp>

using Row = std::vector<CString>;
using Table = std::vector<Row>;

class CExcel
{
public:
	CExcel(std::function<int(const CString&, const CString&)> check_cb);

	~CExcel();

private:
	BOOL CheckSheet(Excel::Sheet* sheet, CString& strClientExportFile, CString &strServerExportFile);

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

	CString GetCellData(Excel::Sheet* sheet, int row, int col);

public:
	const Table &Read(const CString& file, int sheet);

	BOOL Check(const CString &file, std::vector<std::pair<CString, CString> > &vClientExportFile, std::vector<std::pair<CString, CString> > &vServerExportFile, std::function<int(int)> SetProgressBarCallBack, std::function<int(int)> UpdateProgressBarCallBack);

public:
	static std::string CString2String(CString from);

	CString DoubleToString(double n);

private:
	CString m_err;
	CString m_types;
	CString m_file;
	CString m_sheet;
	std::map<int, CString> m_colInfo;
	std::function<int(const CString&, const CString&)> m_callback;
	std::map <CString, std::function<BOOL (const CString& data)> > m_checkFuncList;
	
	std::map<int, Table> m_data;
};

#endif