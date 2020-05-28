//
// Created by Troldal on 2019-01-10.
//

#include <XLDocument.h>
#include <XLDocument_Impl.h>

using namespace OpenXLSX;

/**********************************************************************************************************************/
XLDocument::XLDocument()
        : m_document(std::make_shared<Impl::XLDocument>()) {
}

/**********************************************************************************************************************/
XLDocument::XLDocument(const XLString& name)
        : m_document(std::make_shared<Impl::XLDocument>(name)) {
}

/**********************************************************************************************************************/
void XLDocument::OpenDocument(const XLString& fileName) {

    m_document = std::make_shared<Impl::XLDocument>();
    m_document->OpenDocument(fileName);
}

/**********************************************************************************************************************/
void XLDocument::CreateDocument(const XLString& fileName) {

    if (!m_document)
        m_document = std::make_shared<Impl::XLDocument>();
    m_document->CreateDocument(fileName);
}

/**********************************************************************************************************************/
void XLDocument::CloseDocument() {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    m_document->CloseDocument();
    m_document = nullptr;
}

/**********************************************************************************************************************/
bool XLDocument::SaveDocument() {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->SaveDocument();
}

/**********************************************************************************************************************/
bool XLDocument::SaveDocumentAs(const XLString& fileName) {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->SaveDocumentAs(fileName);
}

/**********************************************************************************************************************/
/*const XLString& XLDocument::DocumentName() const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->DocumentName();
}*/

/**********************************************************************************************************************/
/*const XLString& XLDocument::DocumentPath() const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->DocumentPath();
}*/

/**********************************************************************************************************************/
XLWorkbook XLDocument::Workbook() {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return XLWorkbook(*m_document->Workbook());
}

/**********************************************************************************************************************/
const XLWorkbook XLDocument::Workbook() const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return static_cast<const XLWorkbook>(*m_document->Workbook());
}

/**********************************************************************************************************************/
/*XLString XLDocument::GetProperty(XLProperty theProperty) const {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    return m_document->GetProperty(theProperty);
}*/

/**********************************************************************************************************************/
void XLDocument::SetProperty(XLProperty theProperty, const XLString& value) {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    m_document->SetProperty(theProperty, value);
}

/**********************************************************************************************************************/
void XLDocument::DeleteProperty(XLProperty theProperty) {

    if (!m_document)
        throw XLException("Invalid XLDocument object!");
    m_document->DeleteProperty(theProperty);
}
