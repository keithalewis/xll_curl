// xll_string.cpp - std::string wrappers
#include "xll24/include/xll.h"

using namespace xll;

AddIn xai_string(
	Function(XLL_HANDLEX, L"xll_string", L"\\STRING")
	.Arguments({
		Arg(XLL_PSTRING4, L"value", L"is the string value.")
		})
	.Uncalced()
	.FunctionHelp(L"Return a handle to a std::string.")
	.HelpTopic(L"https://en.cppreference.com/w/cpp/string/basic_string")
);
HANDLEX WINAPI xll_string(const char* value)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<std::string> h_(new std::string(value + 1, value + 1 + value[0]));
		h = h_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return h;
}

AddIn xai_string_append(
	Function(XLL_HANDLEX, L"xll_string_append", L"STRING.APPEND")
	.Arguments({
		Arg(XLL_HANDLEX, L"handle", L"is a handle to a std::string."),
		Arg(XLL_PSTRING4, L"str", L"is a string to append."),
		})
		.Category(L"XLL")
	.FunctionHelp(L"Return appended string.")
	.HelpTopic(L"https://en.cppreference.com/w/cpp/string/basic_string/append")
);
HANDLEX WINAPI xll_string_append(HANDLEX s, const char* str)
{
#pragma XLLEXPORT
	try {
		handle<std::string> s_(s);
		ensure(s_);
		s_->append(str + 1, str[0]);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return s;
}

AddIn xai_string_substr(
	Function(XLL_LPOPER, L"xll_string_substr", L"STRING.SUBSTR")
	.Arguments({
		Arg(XLL_HANDLEX, L"handle", L"is a handle to a std::string."),
		Arg(XLL_UINT, L"pos", L"is the starting position."),
		Arg(XLL_UINT, L"count", L"is the length of the substring.")
		})
	.Category(L"XLL")
	.FunctionHelp(L"Return a substring of a string.")
	.HelpTopic(L"https://en.cppreference.com/w/cpp/string/basic_string/substr")
);
LPOPER WINAPI xll_string_substr(HANDLEX s, UINT pos, UINT count)
{
#pragma XLLEXPORT
	static OPER result;

	try {
		handle<std::string> s_(s);
		ensure(s_);
		result = OPER(s_->substr(pos, count ? count : 0x7FFE));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &result;
}