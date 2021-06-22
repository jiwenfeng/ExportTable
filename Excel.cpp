#include "pch.h"
#include "Excel.h"
#include <memory>
#include <vector>

CExcel::CExcel(std::function<int(const CString&, const CString&)> check_cb)
{
	m_checkFuncList[_T("int")] = std::bind(&CExcel::IsInt, this, std::placeholders::_1);
	m_checkFuncList[_T("macros")] = std::bind(&CExcel::IsMacros, this, std::placeholders::_1);
	m_checkFuncList[_T("str")] = std::bind(&CExcel::IsString, this, std::placeholders::_1);
	m_checkFuncList[_T("string")] = std::bind(&CExcel::IsString, this, std::placeholders::_1);
	m_checkFuncList[_T("base64")] = std::bind(&CExcel::IsString, this, std::placeholders::_1);
	m_checkFuncList[_T("map")] = std::bind(&CExcel::IsMap, this, std::placeholders::_1);
	m_checkFuncList[_T("array")] = std::bind(&CExcel::IsArray, this, std::placeholders::_1);
	m_checkFuncList[_T("float")] = std::bind(&CExcel::IsFloat, this, std::placeholders::_1);
	m_checkFuncList[_T("float_base64")] = std::bind(&CExcel::IsFloat, this, std::placeholders::_1);

	for (auto i : m_checkFuncList)
	{
		if (!m_types.IsEmpty())
		{
			m_types.Append(_T(","));
		}
		m_types.Append(i.first);
	}
	m_callback = check_cb;
}

CExcel::~CExcel()
{
}

BOOL CExcel::CheckSheet(Excel::Sheet* sheet, CString &strClientExportFile, CString &strServerExportFile)
{
	m_colInfo.clear();
	BOOL bCheckData = FALSE;
	if (sheet->cell(0, 1).getString() == _T("TRUE")) // 导出客户端
	{
		if (strClientExportFile.IsEmpty())
		{
			m_callback(m_sheet, _T("[2, B]需要填写客户端导出文件"));
		}
		else
		{
			int nHeaderRow = sheet->cell(2, 1).getDouble();
			if (CheckHeader(sheet, nHeaderRow - 1, FALSE))
			{
				bCheckData = TRUE;
			}
		}
	}
	if (sheet->cell(0, 2).getString() == _T("TRUE")) // 导出服务端
	{
		strServerExportFile = sheet->cell(1, 2).getString().c_str();
		if (strServerExportFile.IsEmpty())
		{
			m_callback(m_sheet, _T("[2, C]需要填写服务端导出文件"));
		}
		else
		{
			int nHeaderRow = sheet->cell(2, 2).getDouble();
			if (CheckHeader(sheet, nHeaderRow - 1, FALSE))
			{
				bCheckData = TRUE;
			}
		}
	}
	if (!bCheckData)
	{
		return FALSE;
	}
	int nDataRow = sheet->cell(3, 1).getDouble();
	return CheckData(sheet, nDataRow - 1);
}

BOOL CExcel::CheckHeader(Excel::Sheet* sheet, int row, BOOL isServerHeader)
{
	BOOL ret = TRUE;
	std::map<std::wstring, int> names;
	for (int col = 0; col < sheet->columnsCount(); ++col)
	{
		const std::wstring& data = sheet->cell(row, col).getString();
		if (data.empty())
		{
			continue;
		}
		size_t s1 = data.find('(');
		if (std::wstring::npos == s1)
		{
			continue;
		}
		int s2 = data.find(')');
		if (-1 == s2)
		{
			continue;
		}
		std::wstring strName = data.substr(0, s1);
		std::wstring strType = data.substr(s1 + 1, s2 - s1 - 1);
		if (strName.empty() || strType.empty())
		{
			continue;
		}

		if (m_checkFuncList.find(strType.c_str()) == m_checkFuncList.end())
		{
			m_err.Format(_T("[%d,%c]错误,列类型只支持[%s], 你填写的是%s"), row + 1, col + 65, m_types.GetBuffer(), strType.c_str());
			m_callback(m_sheet, m_err);
			ret = FALSE;
			continue;
		}
		std::map<std::wstring, int>::iterator itr = names.find(strName);
		if (itr != names.end())
		{
			m_err.Format(_T("[%d,%c]列和[%d,%c]列重名"), row + 1, col + 65, row + 1, itr->second + 65);
			m_callback(m_sheet, m_err);
			ret = FALSE;
			continue;
		}
		names[strName] = col;
		m_colInfo[col] = strType.c_str();
	}
	return ret;
}

BOOL CExcel::CheckData(Excel::Sheet* sheet, int row)
{
	std::map<CString, int> m_idList;
	BOOL ret = TRUE;
	for (int r = row; r < sheet->rowsCount(); ++r)
	{
		for (int c = 0; c < sheet->columnsCount(); ++c)
		{
			CString data = GetCellData(sheet, r, c);
			if (c == 0)
			{
				if (data.IsEmpty())
				{
					m_err.Format(_T("[%d, %c]不能为空"), r + 1, c + 65);
					m_callback(m_sheet, m_err);
				}
				std::map<CString, int>::const_iterator itr = m_idList.find(data);
				if (itr != m_idList.end())
				{
					m_err.Format(_T("[%d, %c]数据和第[%d, A]重复"), r + 1, c + 65, itr->second);
					m_callback(m_sheet, m_err);
				}
			}
			std::map<int, CString>::iterator i = m_colInfo.find(c);
			if (i == m_colInfo.end())
			{
				return TRUE;
			}
			if (!CheckDataFormat(sheet, i->second, data))
			{
				m_err.Format(_T("[%d, %c]:%s不是:%s"), r + 1, c + 65, data.GetBuffer(),i->second.GetBuffer());
				m_callback(m_sheet, m_err);
			}
		}
	}
	return TRUE;
}

BOOL CExcel::CheckDataFormat(Excel::Sheet* sheet, const CString &type,  const CString &data)
{
	if (data.IsEmpty())
	{
		return TRUE;
	}
	auto itr = m_checkFuncList.find(type);
	if (itr == m_checkFuncList.end())
	{
		CString err = _T("类型:") + type + _T("没有找到检查函数,检查失败");
		return FALSE;
	}
	return itr->second(data);
}

BOOL CExcel::IsInt(const CString &data)
{
	int i = 0;
	if (data[0] == '-')
	{
		i = 1;
	}
	if (i == data.GetLength())
	{
		return FALSE;
	}
	if (data.GetLength() != 1 && data[i] == '0')
	{
		return FALSE;
	}
	for (; i < data.GetLength(); ++i)
	{
		if (data[i] < '0' || data[i] > '9')
		{
			return FALSE;
		}
	}
	return TRUE;
}

BOOL CExcel::IsFloat(const CString& data)
{
	return IsDigit(data);
}

BOOL CExcel::IsString(const CString& data)
{
	if (data[0] == '"')
	{
		if (data[data.GetLength() - 1] != '"')
		{
			return FALSE;
		}
	}
	return TRUE;
}

BOOL CExcel::IsMacros(const CString& data)
{
	return TRUE;
}

BOOL CExcel::IsMap(const CString& data)
{
	return MapIsValid(data);
}

BOOL CExcel::MapIsValid(const CString& data)
{
	if (data == _T("{}"))
	{
		return TRUE;
	}
	int len = data.GetLength();
	if (len < 2)
	{
		return FALSE;
	}

	if (data[0] != '{' || data[len - 1] != '}')
	{
		return FALSE;
	}

	CString value = data.Mid(1, len - 2);
	return MapDataIsValid(value);
}

BOOL CExcel::MapDataIsValid(const CString& data)
{
	int counter = 0, counter1 = 0;
	int i = 0, start = 0;
	int len = data.GetLength();
	while (i < len)
	{
		for (; i < len; ++i)
		{
			if (data[i] == '{')
			{
				counter++;
			}
			else if (data[i] == '[')
			{
				counter1++;
			}
			else if (data[i] == '}')
			{
				counter--;
			}
			else if (data[i] == ']')
			{
				counter1--;
			}
			if (data[i] == ',')
			{
				if (counter == 0 && counter1 == 0)
				{
					break;
				}
			}
		}
		CString value = data.Mid(start, i - start);
		if (!CheckMapDataFormat(value))
		{
			return FALSE;
		}
		i++;
		while (i < len && data[i] == ' ')
		{
			i++;
		}
		start = i;
		i++;
	}
	return TRUE;
}

BOOL CExcel::CheckMapDataFormat(const CString& data)
{
	if (data.IsEmpty())
	{
		return FALSE;
	}
	int pos = data.Find(':');
	if (-1 == pos)
	{
		return FALSE;
	}
	CString key = data.Left(pos);
	if (!MapKeyIsValid(key))
	{
		return FALSE;
	}
	while (data[pos + 1] == ' ')
	{
		pos++;
	}
	CString value = data.Right(data.GetLength() - pos - 1);
	if (!MapValueIsValid(value))
	{
		return FALSE;
	}
	return TRUE;
}

BOOL CExcel::MapKeyIsValid(const CString& data)
{
	if (data.IsEmpty())
	{
		return FALSE;
	}
	if (data[0] == '"')
	{
		return IsString(data);
	}
	return IsDigit(data);
}

BOOL CExcel::MapValueIsValid(const CString& data)
{
	if (data[0] == '{')
	{
		if (data[data.GetLength() - 1] != '}')
		{
			return FALSE;
		}
		return MapIsValid(data);
	}
	if (data[0] == '[')
	{
		if (data[data.GetLength() - 1] != ']')
		{
			return FALSE;
		}
		return ArrayIsValid(data);
	}
	return MapKeyIsValid(data);
}

BOOL CExcel::IsArray(const CString& data)
{
	if (data[0] != '[')
	{
		return ArrayIsValid(_T("[") + data + _T("]"));
	}
	return ArrayIsValid(data);
}

BOOL CExcel::ArrayIsValid(const CString& data)
{
	if (data == _T("[]"))
	{
		return TRUE;
	}
	if (data[0] != '[')
	{
		return FALSE;
	}
	if (data[data.GetLength() - 1] != ']')
	{
		return FALSE;
	}
	CString val = data.Mid(1, data.GetLength() - 2);
	std::vector<CString>  v;
	int c1 = 0, c2 = 0, c3 = 0, start = 0, i;
	for ( i = 0; i < val.GetLength(); ++i)
	{
		char ch = val.GetAt(i);
		switch (ch)
		{
		case '{':
			c1++;
			break;
		case '}':
			c1--;
			break;
		case '[':
			c2++;
			break;
		case ']':
			c2--;
			break;
		case '"':
			c3++;
			break;
		case ',':
		{
			if (c1 == 0 && c2 == 0 && (c3 % 2) == 0)
			{
				CString d = val.Mid(start, i - start);
				char c = d.GetAt(0);
				if (c != '"')
				{
					d.Insert(0, '"');
					d.AppendChar('"');
				}
				v.push_back(d);
				start = i+1;
			}
			break;
		}
		default:
			break;
		}
	}
	if (start < val.GetLength())
	{
		CString d = val.Mid(start, val.GetLength() - start);
		char c = d.GetAt(0);
		if (c != '"')
		{
			d.Insert(0, '"');
			d.AppendChar('"');
		}
		v.push_back(d);
	}
	for (std::vector<CString>::iterator i = v.begin(); i != v.end(); ++i)
	{
		char c = i->GetAt(0);
		if (c == '[')
		{
			if (!ArrayIsValid(*i))
			{
				return FALSE;
			}
		}
		else if (c == '{')
		{
			if (!MapIsValid(*i))
			{
				return FALSE;
			}
		}
		else if (c == '"')
		{
			if (!IsString(*i))
			{
				return FALSE;
			}
		}
		else
		{
			if (!IsDigit(*i))
			{
				return FALSE;
			}
		}
	}
	return TRUE;
}

BOOL CExcel::IsDigit(const CString& data)
{
	int start = 0;
	if (data[0] == '-')
	{
		start = 1;
	}
	int hasPoint = 0;
	for (int i = start; i < data.GetLength(); ++i)
	{
		if (data[i] == '.')
		{
			if (hasPoint)
			{
				return FALSE;
			}
			hasPoint = 1;
			continue;
		}
		if (data[i] < '0' || data[i] > '9')
		{
			return FALSE;
		}
	}
	return TRUE;
}

std::string CExcel::CString2String(CString from)
{
	USES_CONVERSION;
	char* p = T2A(from.GetBuffer());
	std::string result;
	if (NULL == p)
	{
		return result;
	}
	result = p;
	return result;
}

CString CExcel::DoubleToString(double n)
{
	CString str;
	str.Format(_T("%lf"), n);
	int i = str.GetLength() - 1;
	while (i > 0)
	{
		char ch = str.GetAt(i);
		if (ch == '.')
		{
			i--;
			break;
		}
		if (ch > '0')
		{
			break;
		}
	}
	if (i > 0)
	{
		return str.Left(i);
	}
	return str;
}

CString CExcel::GetCellData(Excel::Sheet* sheet, int row, int col)
{
	Excel::Cell::DataType type = sheet->cell(row, col).dataType();
	CString data;
	if (type == Excel::Cell::DataType::String)
	{
		data = sheet->cell(row, col).getString().c_str();
	}
	else if (type == Excel::Cell::DataType::Double)
	{
		double nData = sheet->cell(row, col).getDouble();
		data = DoubleToString(nData);
	}
	else if (type == Excel::Cell::DataType::Formula)	// 公式
	{
		const Excel::Cell& cell = sheet->cell(row, col);
		const Excel::Formula& formula = cell.getFormula();
		Excel::Formula::ValueType vt = formula.valueType();
		switch (vt)
		{
		case Excel::Formula::ValueType::DoubleValue:
		{
			data = DoubleToString(formula.getDouble());
			break;
		}
		case Excel::Formula::ValueType::StringValue:
		{
			break;
		}
		default:
			break;
		}
	}

	CString strRet;

	for (int i = 0; i < data.GetLength(); ++i)
	{
		char ch = data.GetAt(i);
		if (ch == '\t' || ch == '\r' || ch == '\n' || ch == ' ')
		{
			continue;
		}
		strRet.AppendChar(ch);
	}
	return strRet;
}

const Table & CExcel::Read(const CString& file, int nsheet)
{
	static Table empty;
	try
	{
		m_file = _T("[") + file + _T("]");
		Excel::Book book(CExcel::CString2String(file));
		if (book.sheetsCount() <= nsheet)
		{
			return empty;
		}
		Excel::Sheet* sheet = book.sheet(nsheet);
		if (NULL == sheet)
		{
			return empty;
		}
		std::map<int, Table>::iterator itr = m_data.find(nsheet);
		if (itr != m_data.end())
		{
			return itr->second;
		}
		Table table;
		for (int i = 0; i < sheet->rowsCount(); ++i)
		{
			Row row;
			for (int j = 0; j < sheet->columnsCount(); ++j)
			{
				CString data = GetCellData(sheet, i, j);
				row.push_back(data);
			}
			table.push_back(row);
		}
		m_data[nsheet] = table;
		return m_data[nsheet];
	}
	catch (const Excel::Exception& e)
	{
		return empty;
	}
	catch (const CompoundFile::Exception& e)
	{
		return empty;
	}
	catch (const std::exception& e)
	{
		return empty;
	}
	return empty;
}

BOOL CExcel::Check(const CString& file, std::vector<std::pair<CString, CString> >& vClientExportFile, std::vector<std::pair<CString, CString> >& vServerExportFile, std::function<int(int)> SetProgressBarCallBack, std::function<int(int)> UpdateProgressBarCallBack)
{
	try
	{
		std::map<CString, CString> mClientExportFile, mServerExportFile;
		m_file = _T("[") + file + _T("]");
		Excel::Book book(CString2String(file));
		int nSheets = book.sheetsCount();
		SetProgressBarCallBack(nSheets);
		BOOL ret = TRUE;
		for (int i = 0; i < nSheets; ++i)
		{
			m_sheet.Format(_T("%s"), book.sheetName(i).c_str());
			Excel::Sheet* sheet = book.sheet(i);
			if (NULL == sheet)
			{
				continue;
			}
			CString strClientExportFile, strServerExportFile;
			BOOL ret = CheckSheet(sheet, strClientExportFile, strServerExportFile);
			if (ret)
			{
				std::map<CString, CString>::iterator ci = mClientExportFile.find(strClientExportFile);
				if (ci != mClientExportFile.end())
				{
					m_callback(_T(""), _T("sheet[") + ci->second + _T("]和sheet[") + m_sheet + _T("]导出的客户端文件重名"));
					return FALSE;
				}
				std::map<CString, CString>::iterator si = mServerExportFile.find(strServerExportFile);
				if (ci != mServerExportFile.end())
				{
					m_callback(_T(""), _T("sheet[") + ci->second + _T("]和sheet[") + m_sheet + _T("]导出的服务端文件重名"));
					return FALSE;
				}
				vClientExportFile.push_back(std::make_pair(strClientExportFile, m_sheet));
				vServerExportFile.push_back(std::make_pair(strServerExportFile, m_sheet));
			}
			else
			{
				ret = FALSE;
			}
			UpdateProgressBarCallBack(i + 1);
		}
		return ret;
	}
	catch (const Excel::Exception& e)
	{
		m_callback(_T(""), e.whatAsWString().c_str());
		return FALSE;
	}
	catch (const CompoundFile::Exception& e)
	{
		m_callback(_T(""), _T("当前只支持检查.xls文件"));
		return FALSE;
	}
	catch (const std::exception& e)
	{
		m_callback(_T(""), CString(e.what()));
		return FALSE;
	}
	return TRUE;
}
