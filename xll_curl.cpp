// xll_curl.cpp - curl wrapper for Excel
#include <thread>
#include "xll_curl.h"

// https://curl.se/libcurl/c/curl_global_init.html
curl::init init;

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
		OPER(L"version_num"), Num(info->version_num),
		OPER(L"host"), OPER(info->host),
		OPER(L"features"), OPER(info->features),
		OPER(L"brotli_version"), OPER(info->brotli_version),
		OPER(L"ssl_version"), OPER(info->ssl_version), 
		OPER(L"libz_version"), OPER(info->libz_version),
		OPER(L"zstd_version"), OPER(info->zstd_version),
		}
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
		handle<curl::easy> p(new curl::easy(url));
		h = p.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
	
	return h;
}

// Callback for curl_easy_perform.
void WINAPI perform(OPER async, curl::easy* pcurl, HANDLEX s)
{
	std::string* str;
	ensure (str = dynamic_cast<std::string*>(to_pointer<std::string>(s)));

	str->clear();
	pcurl->setopt(CURLOPT_WRITEDATA, str);
	pcurl->perform();

	XLOPER12 result = Num(s);

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
		handle<curl::easy> h_(h);
		ensure(h_);

		std::jthread t(perform, *async, h_.ptr(), s);
		t.detach();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}
}
