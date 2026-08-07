//%attributes = {}
$XLSX:=Folder:C1567(fk resources folder:K87:11).file("unprotected.xlsx").getContent()

$params:=New object:C1471("password"; "1234"; "secret"; "abcd")

$status:=msoffice encrypt($XLSX; $params)

If ($status.success)
	
	Folder:C1567(fk desktop folder:K87:19).file("protected.xlsx").setContent($XLSX)
	
End if 

$XLSX:=Folder:C1567(fk resources folder:K87:11).file("protected.xlsx").getContent()

$params:=New object:C1471("password"; "1234"; "secret"; "abcd")

$status:=msoffice decrypt($XLSX; $params)

If ($status.success)
	
	Folder:C1567(fk desktop folder:K87:19).file("unprotected.xlsx").setContent($XLSX)
	
End if 