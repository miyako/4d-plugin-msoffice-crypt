![version](https://img.shields.io/badge/version-17%2B-3E8B93)
![platform](https://img.shields.io/static/v1?label=platform&message=mac-intel%20|%20mac-arm%20|%20win-64&color=blue)
m%20|%20win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-msoffice-crypt)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-msoffice-crypt/total)

# 4d-plugin-msoffice-crypt
Tool to encrypt or decrypt OOXML documents.

Based on [herumi/msoffice](https://github.com/herumi/msoffice).

**Note**: The original source code links to a library that uses [Intel Intrinsics](https://www.intel.com/content/www/us/en/docs/intrinsics-guide/index.html). In particular, the file `uint32vec.hpp` goes like

```c
#ifdef _WIN32
	#include <winsock2.h>
	#include <intrin.h>
#else
	#ifdef __linux__
		#include <x86intrin.h>
	#else
#ifdef __x86_64__
        #include <emmintrin.h>
#endif
	#endif
#endif
```

meaning all non-Windows, non-Linux platform (i.e. macOS) defaults to Intel.

Thanks to [jratcliff63367/sse2neon](https://github.com/jratcliff63367/sse2neon) we can port such code to Apple Silicon.

## Usage

```4d
$XLSX:=Folder(fk resources folder).file("unprotected.xlsx").getContent()

$params:=New object("password";"1234";"secret";"abcd")

$status:=msoffice encrypt ($XLSX;$params)

If ($status.success)
	
	Folder(fk desktop folder).file("protected.xlsx").setContent($XLSX)
	
End if 

$XLSX:=Folder(fk resources folder).file("protected.xlsx").getContent()

$params:=New object("password";"1234";"secret";"abcd")

$status:=msoffice decrypt ($XLSX;$params)

If ($status.success)
	
	Folder(fk desktop folder).file("unprotected.xlsx").setContent($XLSX)
	
End if 
```
