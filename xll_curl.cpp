// xll_curl.cpp - curl wrapper for Excel
#include <thread>
#include "xll_curl.h"

struct curl_init {
	CURLcode code;
	curl_init() : code(curl_global_init(CURL_GLOBAL_ALL | CURL_GLOBAL_WIN32))
	{
		if (CURLE_OK != code) {
			XLL_ERROR(curl_easy_strerror(code));
		}
	}
	curl_init(const curl_init&) = delete;
	curl_init& operator=(const curl_init&) = delete;
	~curl_init()
	{
		curl_global_cleanup();
	}
} init;


size_t writer(char* data, size_t size, size_t nmemb, std::string* writerData)
{
	if (!writerData )
		return 0;

	writerData->append(data, size * nmemb);

	return size * nmemb;
}

class curl_easy {
	CURL* curl;
public:
	curl_easy(const char* url = nullptr)
		: curl{curl_easy_init()}
	{
		if (url && *url) {
			CURLcode code = curl_easy_setopt(curl, CURLOPT_URL, url);
			if (CURLE_OK != code) {
				XLL_ERROR(curl_easy_strerror(code));
			}
			else {
				code = curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, writer);
				if (CURLE_OK != code) {
					XLL_ERROR(curl_easy_strerror(code));
				}
			}
		}
	}
	curl_easy(const curl_easy&) = delete;
	curl_easy& operator=(const curl_easy&) = delete;
	~curl_easy()
	{
		curl_easy_cleanup(curl);
	}
	operator CURL* () const
	{
		return curl;
	}
};

using namespace xll;

AddIn xai_curl_version_info(
	Function(XLL_LPOPER, L"xll_curl_version_info", CATEGORY L".VERSION.INFO")
	.Arguments({})
	.Category(CATEGORY)
	.FunctionHelp(L"Returns runtime libcurl version info.")
);
LPOPER WINAPI xll_curl_version_info()
{
#pragma XLLEXPORT
	static OPER result;

	curl_version_info_data* info = curl_version_info(CURLVERSION_NOW);
	result = OPER({
		OPER(L"version"), OPER(info->version), 
		OPER(L"ssl_version"), OPER(info->ssl_version), 
		OPER(L"libz_version"), OPER(info->libz_version)}
	);
	result.reshape(size(result) / 2, 2);

	return &result;	
}

AddIn xai_curl_easy(
	Function(XLL_HANDLEX, L"xll_curl_easy_init", L"\\" CATEGORY L".EASY.INIT")
	.Arguments({
		Arg(XLL_CSTRING4, L"url", L"is the URL to fetch.")
		})
	.Uncalced()
	.Category(CATEGORY)
	.FunctionHelp(L"Returns handle to a curl easy handle.")
	.HelpTopic(L"https://curl.se/libcurl/c/curl_easy_init.html")
);
HANDLEX WINAPI xll_curl_easy_init(const char* url)
{
#pragma XLLEXPORT
	HANDLEX h = INVALID_HANDLEX;

	try {
		handle<curl_easy> p(new curl_easy(url));
		h = p.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	
	return h;
}

void WINAPI perform(OPER async, CURL* curl, HANDLEX s)
{
	CURLcode code = curl_easy_setopt(curl, CURLOPT_WRITEDATA, to_pointer<std::string>(s));

	if (code != CURLE_OK) {
		XLL_ERROR(curl_easy_strerror(code));
	}
	code = curl_easy_perform(curl);
	if (CURLE_OK != code) {
		XLL_ERROR(curl_easy_strerror(code));
	}

	XLOPER12 result = { .val = {.num = s}, .xltype = xltypeNum };

	Excel12(xlAsyncReturn, 0, 2, &async, &result);
}

AddIn xai_curl_easy_perform(
	Function(XLL_VOID, L"xll_curl_easy_perform", CATEGORY L".EASY.PERFORM")
	.Arguments({
		Arg(XLL_HANDLEX, L"handle", L"is a handle to a curl easy handle."),
		Arg(XLL_HANDLEX, L"string", L"is a handle to a std::string.")
		})
	.Asynchronous()
	.Category(CATEGORY)
	.FunctionHelp(L"Return handle to a std::string by performing the transfer as described in the handle.")
	.HelpTopic(L"https://curl.se/libcurl/c/curl_easy_perform.html")
);
VOID WINAPI xll_curl_easy_perform(HANDLEX h, HANDLEX s, LPOPER async)
{	
#pragma XLLEXPORT
	try {
		handle<curl_easy> h_(h);
		ensure(h_);

		std::jthread t(perform, *async, h_->operator CURL*(), s);
		t.detach();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
}

AddIn xai_string(
	Function(XLL_HANDLEX, L"xll_string", L"\\STRING")
	.Arguments({
		Arg(XLL_CSTRING4, L"value", L"is the string value.")
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
		handle<std::string> h_(new std::string(value));
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
		Arg(XLL_CSTRING4, L"str", L"is a string to append."),
		})
	.Category(CATEGORY)
	.FunctionHelp(L"Return appended string.")
	.HelpTopic(L"https://en.cppreference.com/w/cpp/string/basic_string/append")
);
HANDLEX WINAPI xll_string_append(HANDLEX s, const char* str)
{
#pragma XLLEXPORT
	try {
		handle<std::string> s_(s);
		ensure(s_);
		s_->append(str);
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
	.Category(CATEGORY)
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
		result = OPER(s_->substr(pos, count ? count : std::string::npos));
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return &result;
}