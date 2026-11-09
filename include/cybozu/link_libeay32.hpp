#pragma once
/**
	@file
	@brief link libeay32.lib of openssl
	@author MITSUNARI Shigeo(@herumi)
*/
#if defined(_WIN32) && defined(_MT)
	#if _MSC_VER >= 1900 // VC2015
		// static OpenSSL libcrypto; the project supplies its directory (OpenSSLDir)
		#pragma comment(lib, "libcrypto.lib")
//	#elif _MSC_VER == 1800 // VC2013
	#else
		#pragma comment(lib, "mt/12/libeay32.lib")
	#endif
	#pragma comment(lib, "advapi32.lib")
	#pragma comment(lib, "gdi32.lib")
	#pragma comment(lib, "user32.lib")
	#pragma comment(lib, "crypt32.lib")
	#pragma comment(lib, "ws2_32.lib")
#endif
