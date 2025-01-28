// xll_curl.cpp - curl wrapper for Excel
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
	Function(XLL_HANDLEX, L"xll_curl_easy_init", CATEGORY L".EASY.INIT")
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

AddIn xai_curl_easy_perform(
	Function(XLL_CSTRING, L"xll_curl_easy_perform", CATEGORY L".EASY.PERFORM")
	.Arguments({
		Arg(XLL_HANDLEX, L"handle", L"is a handle to a curl easy handle.")
		})
	.Category(CATEGORY)
	.FunctionHelp(L"Perform the transfer as described in the handle.")
	.HelpTopic(L"https://curl.se/libcurl/c/curl_easy_perform.html")
);
const char* WINAPI xll_curl_easy_perform(HANDLEX h)
{	
#pragma XLLEXPORT
	static std::string buffer;

	try {
		handle<curl_easy> p(h);
		ensure(p);
		CURLcode code = curl_easy_setopt(*p, CURLOPT_WRITEDATA, &buffer);
		if (code != CURLE_OK) {
			XLL_ERROR(curl_easy_strerror(code));
		}
		code = curl_easy_perform(*p);
		if (CURLE_OK != code) {
			XLL_ERROR(curl_easy_strerror(code));
		}
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());
	}

	return buffer.c_str();
}
