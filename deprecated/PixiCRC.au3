;;;;
;;  Name  :  PixiClientRemoteControl
;;  Author:  Matthias Ries <matthias.ries@coffein-shock.de>
;;  Darum :  15.12.2009
;;  Version: 1.0
;;  Description: Bringt Pixi Kundenservice in den Vordergrund 
;;  und sucht nach dem übergebenen Parameter
;;;;  


Dim $pixi_cs_email      = "{TAB}"
Dim $pixi_cs_kdnr       = ""
Dim $pixi_cs_name       = "{TAB 2}"
Dim $pixi_cs_oid_pixi   = "{TAB 10}"
Dim $pixi_cs_oid_shop   = "{TAB 11}"
Dim $pixi_cs_invoicenr  = "{TAB 16}"
Dim $pixi_cs_trackingid = "{TAB 17}"

;If Not $CmdLine[1] Then
;	MsgBox(0,"Fehler", "Der Befehl wurde ohne Parameter ausgeführt")
;	Exit
;EndIf

If $CmdLine[0] = 0 Then
	MsgBox(0, "parameter", "Es fehlt ein Parameter")
	MsgBox(0, "Verfügbare Parameter", "'pixi-cs:email=matthias.ries@asmc.it',   ':name=', ':oid-pixi=', ':oid-shop=', ':invoicenr=', ':trackingid='")
	exit
EndIf

Dim $menge   = StringRegExp( $CmdLine[1], '(.*)\:(.*)\=(.*)',0);
Dim $params  = StringRegExp( $CmdLine[1], '(.*)\:(.*)\=(.*)',1);
; $params[0] => pixi-cs
; $params[1] => email
; $params[2] => wert

If $menge = 0 Then
	MsgBox(0, "parameter", "Es fehlt ein Parameter")
	MsgBox(0, "Verfügbare Parameter", "'pixi-cs:email=matthias.ries@asmc.it', ':name=', ':oid-pixi=', ':oid-shop=', ':invoicenr=', ':trackingid='")
	exit
EndIf


; Change into the WinTitleMatchMode that supports classnames and handles
AutoItSetOption("WinTitleMatchMode", 4)
AutoItSetOption("WinDetectHiddenText", 0)

Dim $name   = "pixi* Kundenservice [pixi_ASM]"

Dim $handle = WinGetHandle($name)

If @error Then
    MsgBox(4096, "Error", $name , "Konnte nicht gefunden werden")
	exit
Else
	
EndIf

WinSetState($name,"",@SW_SHOW )
WinActivate($name)
WinWaitActive($name)
sleep(10)
Send("{F5}")
sleep(50)
Dim $text = WinGetText("[active]")
Dim $match = StringRegExp ( $text, "Nur gesperrte Kunden")

If $match = 1 Then
Dim $key   = $params[1]
Dim $value = $params[2]
Dim $field

	Select
		Case $key == "kdnr"
			$field = $pixi_cs_kdnr
		Case $key == "email"
			$field = $pixi_cs_email
		Case $key == "name"
			$field = $pixi_cs_name
		Case $key == "oid-pixi"
			$field = $pixi_cs_oid_pixi
		Case $key == "oid-shop"
			$field = $pixi_cs_oid_shop
		Case $key == "invoicenr"
			$field = $pixi_cs_invoicenr
		Case $key == "trackingid"
			$field = $pixi_cs_trackingid
		Case Else
			MsgBox(0, "INFO", "Kein oder kein bekannter Parameter")
			exit;
		EndSelect 
		
;Eingabemaske löschen
Send("!c")
;An den Anfang springen
Send("!s")
;Eingabefeld wählen
Send($field)
;wert eingeben
Send($value)
;Absenden	 
Send("{ENTER}")
Else
	MsgBox(0, "ERROR", "Kundenservice will nicht wie vorgesehen: Ist vieleicht ein Unterfenster offen?")
EndIf
