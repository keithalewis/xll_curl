#pragma once
#include "curl/curl.h"
#include "xll24/include/xll.h"

#define CATEGORY L"CURL"

namespace curl {

	inline CURLcode ok(CURLcode code)
	{
		if (CURLE_OK != code) {
			XLL_ERROR(curl_easy_strerror(code));
		}

		return code;
	}

	inline CURLUcode ok(CURLUcode code)
	{
		if (code) {
			XLL_ERROR(curl_url_strerror(code));
		}

		return code;
	}

	// https://curl.se/libcurl/c/curl_url.html
	class url {
		CURLU* purl;
	public:
		url()
			: purl{curl_url()}
		{ }
		url(const url& u)
		{
			purl = curl_url_dup(u.purl);
		}
		url& operator=(const url& u)
		{
			if (this != &u) {
				curl_url_cleanup(purl);
				purl = curl_url_dup(u.purl);
			}

			return *this;
		}
		~url()
		{
			curl_url_cleanup(purl);
		}

		operator CURLU* () const
		{
			return purl;
		}

		// https://curl.se/libcurl/c/curl_url_set.html
		// Set a URL part. Use set(part) to unset a part.
		url& set(CURLUPart part, const char* content = nullptr, unsigned int flags = 0)
		{
			ok(curl_url_set(purl, part, content, flags));

			return *this;
		}

		// Linear RAII string type
		class content {
			char* s;
		public:
			content(char* s = nullptr) : s{ s } { }
			content(const content&) = delete;
			content& operator=(const content&) = delete;
			~content() { curl_free(s); }

			operator char**()
			{
				return &s;
			}

			const char* get() const
			{
				return s;
			}
		};

		// https://curl.se/libcurl/c/curl_url_get.html
		// Extract a part from a URL.
		const content& get(CURLUPart part, unsigned int flags) const
		{
			content c;

			ok(curl_url_get(purl, part, c, flags));

			return c;
		}
	};

	// https://curl.se/libcurl/c/curl_global_init.html
	class init {
		CURLcode code;
	public:
		init(long flags = 0) 
			: code(curl_global_init(flags | CURL_GLOBAL_ALL | CURL_GLOBAL_WIN32))
		{
			ok(code);
		}
		init(const init&) = delete;
		init& operator=(const init&) = delete;
		~init()
		{
			curl_global_cleanup();
		}
	};

	// Write to string buffer.
	inline size_t writer(char* data, size_t size, size_t nmemb, std::string* buffer)
	{
		if (!buffer)
			return 0;

		buffer->append(data, size * nmemb);

		return size * nmemb;
	}

	// https://curl.se/libcurl/c/curl_easy_init.html
	class easy {
		CURL* curl;
	public:
		easy(const char* url = nullptr)
			: curl{ curl_easy_init() }
		{
			if (url && *url) {
				setopt(CURLOPT_URL, url);
				setopt(CURLOPT_WRITEFUNCTION, writer);
				// automatic decompression of HTTP downloads
				// https://curl.se/libcurl/c/CURLOPT_ACCEPT_ENCODING.html
				setopt(CURLOPT_ACCEPT_ENCODING, "");
			}
		}
		easy(const easy&) = delete;
		easy& operator=(const easy&) = delete;
		~easy()
		{
			curl_easy_cleanup(curl);
		}

		operator CURL* ()
		{
			return curl;
		}

		template<class Param>
		easy& setopt(CURLoption option, const Param& param)
		{
			ok(curl_easy_setopt(curl, option, param));
			
			return *this;
		}

		void perform()
		{
			ok(curl_easy_perform(curl));
		}
	};

} // namespace curl