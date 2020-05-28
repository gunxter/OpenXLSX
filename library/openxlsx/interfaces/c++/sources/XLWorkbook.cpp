//
// Created by Troldal on 2019-01-16.
//

#include <XLWorkbook.h>
#include <XLWorkbook_Impl.h>
#include "XLSheet_Impl.h"
#include "XLWorksheet_Impl.h"
#include "XLChartsheet_Impl.h"

using namespace OpenXLSX;

/**********************************************************************************************************************/
XLWorkbook::XLWorkbook(Impl::XLWorkbook& workbook)
        : m_workbook(&workbook) {
}

/**********************************************************************************************************************/
XLSheet XLWorkbook::Sheet(unsigned int index) {

    return XLSheet(*m_workbook->Sheet(index));
}

/**********************************************************************************************************************/
const XLSheet XLWorkbook::Sheet(unsigned int index) const {

    return static_cast<const XLSheet>(*m_workbook->Sheet(index));
}

/**********************************************************************************************************************/
XLSheet XLWorkbook::Sheet(const XLString& sheetName) {

    return XLSheet(*m_workbook->Sheet(sheetName));
}

/**********************************************************************************************************************/
const XLSheet XLWorkbook::Sheet(const XLString& sheetName) const {

    return static_cast<const XLSheet>(*m_workbook->Sheet(sheetName));
}

/**********************************************************************************************************************/
XLWorksheet XLWorkbook::Worksheet(const XLString& sheetName) {

    return XLWorksheet(*m_workbook->Worksheet(sheetName));
}

/**********************************************************************************************************************/
const XLWorksheet XLWorkbook::Worksheet(const XLString& sheetName) const {

    return static_cast<const XLWorksheet>(*m_workbook->Worksheet(sheetName));
}

/**********************************************************************************************************************/
XLChartsheet XLWorkbook::Chartsheet(const XLString& sheetName) {

    return XLChartsheet(*m_workbook->Chartsheet(sheetName));
}

/**********************************************************************************************************************/
const XLChartsheet XLWorkbook::Chartsheet(const XLString& sheetName) const {

    return static_cast<const XLChartsheet>(*m_workbook->Chartsheet(sheetName));
}

/**********************************************************************************************************************/
void XLWorkbook::DeleteSheet(const XLString& sheetName) {

    m_workbook->DeleteSheet(sheetName);
}

/**********************************************************************************************************************/
void XLWorkbook::AddWorksheet(const XLString& sheetName, unsigned int index) {

    m_workbook->AddWorksheet(sheetName, index);
}

/**********************************************************************************************************************/
void XLWorkbook::CloneWorksheet(const XLString& extName, const XLString& newName, unsigned int index) {

    m_workbook->CloneWorksheet(extName, newName, index);
}

/**********************************************************************************************************************/
void XLWorkbook::AddChartsheet(const XLString& sheetName, unsigned int index) {

    m_workbook->AddChartsheet(sheetName, index);
}

/**********************************************************************************************************************/
void XLWorkbook::MoveSheet(const XLString& sheetName, unsigned int index) {

    m_workbook->MoveSheet(sheetName, index);
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::IndexOfSheet(const XLString& sheetName) const {

    return m_workbook->IndexOfSheet(sheetName);
}

/**********************************************************************************************************************/
XLSheetType XLWorkbook::TypeOfSheet(const XLString& sheetName) const {

    return m_workbook->TypeOfSheet(sheetName);
}

/**********************************************************************************************************************/
XLSheetType XLWorkbook::TypeOfSheet(unsigned int index) const {

    return m_workbook->TypeOfSheet(index);
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::SheetCount() const {

    return m_workbook->SheetCount();
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::WorksheetCount() const {

    return m_workbook->WorksheetCount();
}

/**********************************************************************************************************************/
unsigned int XLWorkbook::ChartsheetCount() const {

    return m_workbook->ChartsheetCount();
}

/**********************************************************************************************************************/
bool XLWorkbook::SheetName(unsigned int index, char* buf, unsigned int bufLen) const {

    const auto& list = m_workbook->SheetNames();
		if (index < list.size())
		{
			strncpy(buf, list[index].c_str(), bufLen);
			return true;
		}

		return false;
}

/**********************************************************************************************************************/
bool XLWorkbook::WorksheetName(unsigned int index, char* buf, unsigned int bufLen) const {

    const auto& list = m_workbook->WorksheetNames();
		if (index < list.size())
		{
			strncpy(buf, list[index].c_str(), bufLen);
			return true;
		}

		return false;
}

/**********************************************************************************************************************/
bool XLWorkbook::ChartsheetName(unsigned int index, char* buf, unsigned int bufLen) const {

		const auto& list = m_workbook->ChartsheetNames();
		if (index < list.size())
		{
			strncpy(buf, list[index].c_str(), bufLen);
			return true;
		}

		return false;
}

/**********************************************************************************************************************/
bool XLWorkbook::SheetExists(const XLString& sheetName) const {

    return m_workbook->SheetExists(sheetName);
}

/**********************************************************************************************************************/
bool XLWorkbook::WorksheetExists(const XLString& sheetName) const {

    return m_workbook->WorksheetExists(sheetName);
}

/**********************************************************************************************************************/
bool XLWorkbook::ChartsheetExists(const XLString& sheetName) const {

    return m_workbook->ChartsheetExists(sheetName);
}

/**********************************************************************************************************************/
void XLWorkbook::DeleteNamedRanges() {

    m_workbook->DeleteNamedRanges();
}