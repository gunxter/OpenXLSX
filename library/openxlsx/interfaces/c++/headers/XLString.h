#pragma once

#include <string>

namespace OpenXLSX
{
	//! Wrapper around string to avoid using std::string in dll api
	class XLString
	{
	public:
		XLString() {}
		XLString(const char* s) : m_ptr(s) {}
		
		operator const char*() const { return m_ptr; }
		operator std::string() const { return m_ptr; }

	private:
		const char* m_ptr = nullptr;
	};
}