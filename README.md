# 4d-plugin-msoffice-crypt
Tool to encrypt or decrypt OOXML documents.

Based on [herumi/msoffice](https://github.com/herumi/msoffice).

**Note**: macOS is Intel only; Xcode 12 clang fails to compile buitin ARM code.

```sh
use of undeclared identifier '__builtin_ia32_emms'; did you mean
      '__builtin_isless'?
    __builtin_ia32_emms();
```

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
