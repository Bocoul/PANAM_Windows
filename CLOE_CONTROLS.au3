#include-once
#include "Outils Base\MAP.au3"
#include "Outils Base\TOOLS.au3"
;~ #include "Outils Base\CONTROLS.au3"

#region LesOptions ;****Option****
Opt("WinWaitDelay", 2600)
Opt("SendKeyDelay", 20)
Opt("SendKeyDownDelay", 10)
Opt("WinTitleMatchMode", -2) ;1=début du titre , 2=Chaine dans le titre, 3=exact, 4=mode avancé, -1 to -4=Nocase
Opt("WinSearchChildren", 1)
Opt("SendCapslockMode", 1)
Opt("MouseClickDelay", 1000)
Opt("PixelCoordMode", 1)
Opt("MouseCoordMode", 1)
#endregion LesOptions ;****Option****

#region LesDeclarations ;****LesDeclarations****
Local Const  $CRM_WinTitle =  "[REGEXPTITLE:CLOE.+Internet Explorer]"


Dim $WinWaitDelai = 500, $Timeout = 10000

Dim $aMsgRapportAll[3]
Dim $oIECLOE, $oIECLOEView, $oIECLOEFrameCourant,  $NomEcranCourant

;~ Dim $IECLOEDialog_Hwnd, $IEServerDialog_Hwnd, $IEServerDialog_Pos
Dim $IE_CLOE_Hwnd = 0, $IEServer_Hwnd = 0, $ATLChoixEcran = 0, $ATLChoixVue = 0,$ZonesVues_Hwnds, $ChampsEditsVisibles, $ChampsBtnVisibles, $ZoneVue_Hwnd, $ATL
Dim $ChampsEditsVisiblesSup, $ChampsEditsVisiblesInf, $ChampsBtnVisibles
Dim $IE_CLOE_Pos = 0, $IEServer_Pos = 0, $ATLChoixEcran_Pos = 0, $ATLChoixVue_Pos = 0
Dim $ZoneVue_Pos = 0, $ZoneBoutonSup_Pos, $ZoneBoutonInf_Pos, $BtnsCLOEA, $BtnsCLOEB, $BtnsCLOEAll, $BoutonsMenuEcran, $BoutonsMenuvue

Dim $Rapport = "Merci de patienter. La lecture des fichiers paramètres est en cours"

Dim $NomFenCLOE = "CLOE RISPI - Windows Internet Explorer", $NomFenCLOEProd = "CLOE - Windows Internet Explorer"
;~ Dim $WinWaitDelai = 15
Dim $sToolTipAnswer, $Consignes, $ConsignesFichierSource
Dim $AdresseCLOERISPY = "https://pcyyyxp5.pcy.edfgdf.fr/cloe1/start.swe?SWECmd=Login&SWEFullRefresh=1&TglPrtclRfrsh=1"
Dim $AdresseCLOEProd = "http://cloe78.edf.fr/cloe2/start.swe?SWECmd=Start"
Dim $TablDonnees, $RangInitialChampsEditsCLOE = 2

Global $g_aArrayEdits
HotKeySet("{END}", "Terminer")
Send("{CAPSLOCK OFF}")
#endregion LesDeclarations ;****LesDeclarations****


#region LesFonctions ;****LesFonctions****²
Func CLOEUDFWriteError($sSeverity, $sFunc, $sMessage = Default, $sStatus = Default)
	Local $sStr = "--> CLOE_Include.au3 " & $sSeverity & " from function " & $sFunc
	If Not($sMessage = Default) Then $sStr &= ", " & $sMessage
	If Not($sStatus = Default) Then $sStr &= " (" & $sStatus & ")"
	ConsoleWrite($sStr & @CRLF)
	Return SetError($sStatus, 0, 1) ; restore calling @error
EndFunc   ;==>CLOEUDFWriteError

Func CLOEPixelSearch($FenHwnd, $sText, $ControlID, $aCouleurs, $aCoords = Default, $Shade = 0, $Step = 1, $CapturTest = "")
	Opt("PixelCoordMode", 1)

	If IsArray($aCoords) = 0 Then
		$aCoords = ControlGetPos($FenHwnd, $sText, $ControlID)
		$aCoords[0] = 1
		$aCoords[1] = 1
		$aCoords[2] -= 1
		$aCoords[3] -= 1
	EndIf

	Local $SensH = ($aCoords[0] < $aCoords[2]), $SensV = ($aCoords[1] < $aCoords[3])
	Local $CoordsSearch = [ _
			($SensH ? $aCoords[0] : $aCoords[2]), _
			($SensV ? $aCoords[1] : $aCoords[3]), _
			($SensH ? $aCoords[2] : $aCoords[0]), _
			($SensV ? $aCoords[3] : $aCoords[1])]
	WinActivate($FenHwnd)

	$CoordsSearch = RectClientToScreen($ControlID, $CoordsSearch)

	If $CapturTest <> "" Then _ScreenCapture_Capture($CapturTest & "_" & @HOUR & @MIN & @SEC & ".bmp", $CoordsSearch[0], $CoordsSearch[1], $CoordsSearch[2], $CoordsSearch[3])

	For $i = 0 to UBound($aCouleurs) - 1
		Local $coordsClic = PixelSearch( _
				($SensH ? $CoordsSearch[0] : $CoordsSearch[2]), _
				($SensV ? $CoordsSearch[1] : $CoordsSearch[3]), _
				($SensH ? $CoordsSearch[2] : $CoordsSearch[0]), _
				($SensV ? $CoordsSearch[3] : $CoordsSearch[1]), _
				$aCouleurs[$i], $Shade, $Step)
		If @error = 0 Then Return SetError(0, 0, $coordsClic)
	Next
	SetError(@error, 0, 0)
EndFunc   ;==>CLOEPixelSearch

Func CRM_EnumCtrlFields()
	$g_aArrayEdits = CRM_EnumWindows("Button|Edit")

	Local $i = 0
	While $i < UBound($g_aArrayEdits)
		If StringInStr($g_aArrayEdits[$i][9] , "ATL")  Then
			$i += 1
		Else
			_ArrayDelete($g_aArrayEdits, $i)
		EndIf
	WEnd
	Return ($g_aArrayEdits)
EndFunc

Func CRM_EnumWindows( $sClassname = ".*",$bOrdreVertical = False )

	If WaitFunc ("CRM_getIEServer") = 0 Then SetError(1, 9, 0)

	Local $aCtrlWnds = WaitFunc("_WinGetChildrenByClassname", 500, 20000, $IEServer_Hwnd, $sClassname, $bOrdreVertical, True)
	If UBound($aCtrlWnds)  = 0 Then Return SetError(1, 9, 0)  ; ,"CRM_EnumWindows", @ScriptLineNumber, "UBound($aCtrlWnds)  = 0 : $sClassname " & $sClassname)
	Return SetError(0, 0, $aCtrlWnds)
EndFunc

Func CRM_getIEServer()
	Local $sTitle = $CRM_WinTitle , $sText = ""

	If $IE_CLOE_Hwnd = 0 or WinGetHandle($sTitle, $sText) <> $IE_CLOE_Hwnd then $IE_CLOE_Hwnd = Winwait($sTitle, $sText,  10)
	If $IE_CLOE_Hwnd  = 0 Then Return MySetError(1, 9, 0,"CRM_getIEServer", @ScriptLineNumber, "$IE_CLOE_Hwnd = 0 ")

	Local $aCtrlWnds = _WinGetChildrenByClassname($IE_CLOE_Hwnd, "Internet Explorer_Server")
	If UBound($aCtrlWnds)  = 0 Then Return MySetError(1, 9, 0,"CRM_getIEServer", @ScriptLineNumber, "UBound($aCtrlWnds)  = 0 ")
	$IEServer_Hwnd = $aCtrlWnds[0][0]
	Return SetError(0, 0, $IEServer_Hwnd)
EndFunc

Func EstZoneListeVide_($FenHwnd, $sText, $ControlID, $CouleurTexte = 0xEFEF99, $CoordLeft = 0, $CoordTop = 0, $CoordRight = 0, $CoordBottom = 0, $Shade = 0, $Step = 1)
	Opt("PixelCoordMode", 1)

	Local $CoordsSearch = [$CoordLeft, $CoordTop, $CoordRight, $CoordBottom]
	FermerDialog()
	WinActivate($FenHwnd)

	$CoordsSearch = RectClientToScreen($ControlID, $CoordsSearch)
	Local $coordsClic = PixelSearch($CoordsSearch[0], $CoordsSearch[1], $CoordsSearch[2], $CoordsSearch[3], $CouleurTexte, $Shade, $Step)
	If @error = 1 Then
		Return 1
	Else
		Return 0
	EndIf
EndFunc   ;==>EstZoneListeVide_

Func CLOEGetZoneHwnd($sTitle = Default, $sText = "", $iTimeout = 5000,  $WinWaitDelai = 250, $OrderEditsV = False, $GRC = Default)
	$Hwnd = WaitFunctionExt("CLOEGetZoneHwnd_", $WinWaitDelai, $iTimeout, $sTitle, $stext)
	If @extended = 9 Then Return MySetError(@error, @extended, $Hwnd , "CLOEGetZoneHwnd" , @ScriptLineNumber , "hWnd CLOE introuvable")

   ;tous les ATL edits
	$ATL = _ControlGetInstancesVisiblesByClassName($IE_CLOE_Hwnd, $IEServer_Hwnd, "[REGEXPCLASS:ATL]", $OrderEditsV)
	$ChampsEditsVisibles = _ControlGetInstancesVisiblesByClassName($IE_CLOE_Hwnd, $IEServer_Hwnd, "[REGEXPCLASS:Edit|Button]", $OrderEditsV)

	If $ChampsEditsVisibles[0] = 0 Then
		CLOEUDFWriteError("Warning", "CLOEGetZoneHwnd", "$CLOEError_ChampsVisiblesKO")
		$CLOEError += $CLOEError_ChampsVisiblesKO
	Else
		$CLOEStatus += $CLOEStatus_ChampsVisibles_OK
	EndIf

	; La zone vue inferieure
	$ZonesVues_Hwnds = _ControlGetInstancesVisiblesByClassName($IE_CLOE_Hwnd, $IEServer_Hwnd, "[REGEXPCLASS:#32770]")
	If $ZonesVues_Hwnds[0] = 0 Then
		CLOEUDFWriteError("Warning", "CLOEGetZoneHwnd", "$CLOEError_ZoneVueKO")
		$CLOEError += $CLOEError_ZoneVueKO
	Else
		$CLOEStatus += $CLOEStatus_ZoneVue_OK
		$ZoneVue_Hwnd = $ZonesVues_Hwnds[0]
		$ZoneVue_Pos = ControlGetPos($IE_CLOE_Hwnd, "", $ZoneVue_Hwnd)
	EndIf
	CLOEUDFWriteError("Succès", "CLOEGetZoneHwnd", "$CLOEError_Success")
	Return SetError($CLOEError, $CLOEStatus, 1)

EndFunc   ;==>CLOEGetZoneHwnd

Func CLOEGetZoneHwnd_($sTitle = Default, $sText = "",  $OrderEditsV = False)
	$sTitle = ($sTitle = Default ? "CLOE" : $sTitle)
	$CLOEError = 0
	$CLOEStatus = 0

	If $IE_CLOE_Hwnd = 0 or WinGetHandle("[REGEXPTITLE:" & $sTitle & ".+Internet Explorer]", $sText) <> $IE_CLOE_Hwnd then
		$IE_CLOE_Hwnd = Winwait("[REGEXPTITLE:" & $sTitle & ".+Internet Explorer]", $sText,  10)
	EndIf

	If $IE_CLOE_Hwnd = 0 then Return MySetError($CLOEError_hWnd_KO, 9, 0, "CLOEGetZoneHwnd_", "IE introuvable")   ;SetError ($CLOEError_hWnd_KO, 9, "Handle de la  fenetre est nul")

	; Zone client
	$IEServer_Hwnd = _ControlGetInstancesVisiblesByClassName($IE_CLOE_Hwnd, "", "[CLASS:Internet Explorer_Server]")
	$IEServer_Hwnd = $IEServer_Hwnd[0]
	If $IEServer_Hwnd = 0 Then Return MySetError($CLOEError_hWnd_KO, 9, 0, "CLOEGetZoneHwnd_", "IEServer introuvable")

	WinActivate($IE_CLOE_Hwnd)

	WinSetState($IE_CLOE_Hwnd, "", @SW_MAXIMIZE)

	$CLOEStatus += $CLOEStatus_IE_OK
	$IE_CLOE_Pos = WinGetPos($IE_CLOE_Hwnd)
	$CLOEStatus += $CLOEStatus_IEServer_OK
	$IEServer_Pos = ControlGetPos($IE_CLOE_Hwnd, "", $IEServer_Hwnd)

	;ATL menus choix  de l'écran  objet et  vues rattachées
	$ATL = _ControlGetInstancesVisiblesByClassName($IE_CLOE_Hwnd, $IEServer_Hwnd, "[REGEXPCLASS:ATL]", True)

	If $ATL[0] = 0 Then
		CLOEUDFWriteError("ECHEC", "CLOEGetZoneHwnd", "$CLOEError_ATL_MenuEcranKO")
		$CLOEError += $CLOEError_ATL_MenuEcranKO
		Return MySetError($CLOEError, 9, 0, "CLOEGetZoneHwnd_", "ATL menu superieur introuvable") ;MySetError(@error, @extended, $Hwnd , "CLOEGetZoneHwnd" , @ScriptLineNumber , "hWnd CLOE introuvable")
	EndIf

	; Recherche  menu Ecran
	Local $NumATL = 0
	$ATLChoixEcran = 0
	While $NumATL < UBound($ATL)
		$ATLChoixEcran_Pos = ControlGetPos($IE_CLOE_Hwnd, "", $ATL[$NumATL])
		If $ATLChoixEcran_Pos[2] > ($IEServer_Pos[2] * 0.80) Then
			$ATLChoixEcran = $ATL[$NumATL]
			$CLOEStatus += $CLOEStatus_ATL_MenuEcran_OK
			ExitLoop
		EndIf
		$NumATL += 1
	WEnd

	If $ATLChoixEcran = 0 Then
		CLOEUDFWriteError("ECHEC", "CLOEGetZoneHwnd", "$CLOEError_ATL_MenuEcranKO")
		$CLOEError += $CLOEError_ATL_MenuEcranKO
		Return SetError($CLOEError, 9, "Menu supérieur absent")
	EndIf

	$NumATL += 1
	While $NumATL < UBound($ATL)
		$ATLChoixVue_Pos = ControlGetPos($IE_CLOE_Hwnd, "", $ATL[$NumATL])
		If $ATLChoixVue_Pos[2] > ($IEServer_Pos[2] * 0.80) Then
			$ATLChoixVue = $ATL[$NumATL]
			$CLOEStatus += $CLOEStatus_ATL_MenuVue_OK
			ExitLoop
		EndIf
		$NumATL += 1
	WEnd
	If $ATLChoixVue = 0 Then
		CLOEUDFWriteError("Warning", "CLOEGetZoneHwnd", "$CLOEError_ATL_MenuVueKO")
		$CLOEError += $CLOEError_ATL_MenuVueKO
	EndIf

	Return  SetError($CLOEError, 0, 1)
EndFunc   ;==>CLOEGetZoneHwnd

Func SetRapport($MsgRapport, $TitreTraitement = "", $MsgRapport1="", $MsgRapport2="")

	$aMsgRapportAll[0] =  ($MsgRapport = "" ? $aMsgRapportAll[0]  : $MsgRapport)
	$aMsgRapportAll[1] = $MsgRapport1 ; ($MsgRapport1 = "" ? $aMsgRapportAll[1]  : $MsgRapport1)
	$aMsgRapportAll[2] = $MsgRapport2 ; ($MsgRapport2 = "" ? $aMsgRapportAll[2]  : $MsgRapport2)
	Local $sTitre  =  $NomAppl & ($TitreTraitement = "" ? $TypeTraitement : $TitreTraitement)
	Local $sTexte  = _ArrayToString($aMsgRapportAll, @CRLF)

	$sToolTipAnswer = ToolTip($sTexte & @CRLF & "[" & @HOUR & ":" & @MIN & ":" & @SEC & "]" & @CRLF, @DesktopWidth * 0.61, 0, _
					$sTitre, _
					1, _
					4)
	Sleep(50)
	_FileWriteLog(@ScriptDir & "\" & $NomFichierLog, $MsgRapport)
	Return $MsgRapport
EndFunc   ;==>SetRapport
#cs
Func SetRapportold($MsgRapport)

;~ 	If $NomFichierLog = "" Then $NomFichierLog = $NomAppl & "_Journal_" & @MDAY & @MON & @YEAR & ".log"
;~ 	$FichierModelNbreLignesFaites = $FichierModelLigneCourante - $FichierModelLigneEntete
;~ 	$MsgRapport = "Références traitées: " & $FichierModelNbreLignesFaites & " sur " & $FichierModelNbreTotalLignes  _
;~ 	& "| Nombre d'erreurs successives: " & $CompteurErreur  & "| nombre de succès: " & $CompteurSucces& " <<" & AfficherDate() & ">>" & @CRLF & $MsgRapport
;~ 	& @CRLF & $MsgRapport ; & " <<" & AfficherDate() & ">>"
	Local $i = $FichierModelLigneCourante * 100 / $FichierModelNbreTotalLignes
	GUICtrlSetData($Progress1, $i)
	$sToolTipAnswer = ToolTip("[**Ref. CLOE: " & $RefCompte & "**]" & _
			($FichierModelLigneCourante = 0 ? "" : "[**Ligne courante : " & $FichierModelLigneCourante & "**] ... ") & _
			($FichierModelNbreTotalLignes = 0 ? "" : " [** " & "Réf. traitées " & $FichierModelNbreLignesFaites & " sur " & $FichierModelNbreTotalLignes & "**] ... ") & _
			($CompteurErreur <= 1 ? "" : "[** " & $CompteurErreur & " erreurs successives**] ... ") & _
			($CompteurSucces <= 1 ? "" : "[**" & $CompteurSucces & " succès**] ") & _
			"[" & @HOUR & @MIN & @SEC & "]" & @CRLF & _
			$MsgRapport & @CRLF & "***Merci de ne pas utiliser ce poste avant la fin de l'opération***", 560, 0, $NomAppl & $TypeTraitement, 1, 4)
	Sleep(10)
	_FileWriteLog(@ScriptDir & "\" & $NomFichierLog, $MsgRapport)
	Return $MsgRapport
EndFunc   ;==>SetRapportold
#ce
;~
#CS
Func CLOEBoutonClick____($FenHwnd, $sText, $ControlID, $CouleurTexte = 0x000000, $SearchCoords = 0, $Shade = 0, $Step = 1, $DelaiMilliSeconde = 1000, $BtnCaption = "", $CapturTest = "")
;~ 	ConsoleWrite('@@ (317) :(' & @MIN & ':' & @SEC & ') CLOEBoutonClick()' & @CR) ;### Function Trace
;~ 	Local $VarClic = CliquerBoutonCLOE($FenHwnd, $sText, $ControlID, $CouleurTexte, $SearchCoords[0], $SearchCoords[1], $SearchCoords[2], $SearchCoords[3], $DelaiMilliSeconde, $Shade, $Step, $CapturTest)
;~ 	return SetError(@error, @extended, $VarClic)
;~ EndFunc   ;==>CLOEBoutonClick

#CE

; #FUNCTION# ====================================================================================================================
; Name ..........: CliquerBoutonCLOE
; Description ...:
; Syntax ........: CliquerBoutonCLOE($FenHwnd, $sText, $ControlID[, $CouleurTexte = 0x000000[, $CoordLeft = 0[, $CoordTop = 0[,
;                  $CoordRight = 0[, $CoordBottom = 0[, $DelaiMilliSeconde = 1000[, $Shade = 0[, $Step = 1[,
;                  $CapturTest = ""]]]]]]]]])
; Parameters ....: $FenHwnd             - an unknown value.
;                  $sText               - a string value.
;                  $ControlID           - an unknown value.
;                  $CouleurTexte        - [optional] an unknown value. Default is 0x000000.
;                  $CoordLeft           - [optional] an unknown value. Default is 0.
;                  $CoordTop            - [optional] an unknown value. Default is 0.
;                  $CoordRight          - [optional] an unknown value. Default is 0.
;                  $CoordBottom         - [optional] an unknown value. Default is 0.
;                  $DelaiMilliSeconde   - [optional] an unknown value. Default is 1000.
;                  $Shade               - [optional] an unknown value. Default is 0.
;                  $Step                - [optional] an unknown value. Default is 1.
;                  $CapturTest          - [optional] an unknown value. Default is "".
; Return values .:
;             	-1 si la couleur n'est pas  trouvée
;              	 0 Si la couleur est trouvé et le clic ne provoque  pas de changement d'écran attendu
;                1 Si la couleur est trouvé et le clic provoque le changement d'écran attendu
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================

Func CLOEBoutonClick($FenHwnd, $sText, $ControlID, $CouleurTexte = 0x000000, $SearchCoords = 0, $Shade = 0, $Step = 1, $CapturTest = "")
	Local $DelaiMilliSeconde = 500
	$WinWaitDelai = 10
	Opt("PixelCoordMode", 1)

	Local $aPointDebutZone = [$SearchCoords[0], $SearchCoords[1]]
	Local $aPointFinZone = [$SearchCoords[2], $SearchCoords[3]]

	WinActivate($FenHwnd)
   If $ControlID = 0 Then Return SetError(1, 9, -1)
	$aPointDebutZone = _ClientToScreen($ControlID, $aPointDebutZone)
   If $aPointDebutZone = 0 Then Return SetError(1, 9, -1)
	$aPointFinZone = _ClientToScreen($ControlID, $aPointFinZone)

	If $CapturTest <> ""  And not @Compiled Then
;~ 		_ScreenCapture_Capture($CapturTest & "_" & @HOUR & @MIN & @SEC & ".bmp",$aPointDebutZone[0], $aPointDebutZone[1], $aPointFinZone[0], $aPointFinZone[1])
	EndIf

	Local $coordsClic = PixelSearch($aPointDebutZone[0], _
			$aPointDebutZone[1], _
			$aPointFinZone[0], _
			$aPointFinZone[1], _
			$CouleurTexte, $Shade, $Step)
	If @error = 1 Then
		Return SetError(1, 0, -1)
	EndIf
	Local $SommePxEcran = PixelChecksum($SearchCoords[0], $SearchCoords[1], $SearchCoords[2], $SearchCoords[3], $ControlID)

	$coordsClic = _ScreenToClient($ControlID, $coordsClic)
	ControlClick($FenHwnd, $sText, $ControlID, "left", 1, $coordsClic[0], $coordsClic[1])
	Opt("WinWaitDelay", $DelaiMilliSeconde)
	WinWait($FenHwnd, "", $WinWaitDelai)

   Return SetError(0, 9, ($SommePxEcran <> PixelChecksum($SearchCoords[0], $SearchCoords[1], $SearchCoords[2], $SearchCoords[3])) ? 1 : 0)
EndFunc   ;==>CLOEBoutonClick



Func OuvrirSessionCLOE($NNISesame = "", $mdpSesame = "")
	Opt("WinWaitDelay", 2000)

	Local $Html_s_swepi_22
	Local $NbreLoop
	While 1
		$NbreLoop = $NbreLoop + 1
		If $NbreLoop > $NbreMaxTentativesConnexion Then
			Return SetError(1, 0, 0)
		EndIf
		$IE_CLOE_Hwnd = WinWait("[Class:IEFrame]", "", $WinWaitDelai)
		If $IE_CLOE_Hwnd = 0 Then
			Return SetError(1, 0, 0)
		EndIf

		Local $oIE = _IEAttach($IE_CLOE_Hwnd, "hwnd")
		If isobj($oIE) = 0 Then
			Sleep(1000)

			ContinueLoop
		EndIf
		_IELoadWait($oIE, 1000, 20000)

		Local $Html_overridelink = $oIE.Document.body.all("overridelink")
		If IsObj($Html_overridelink) = 0 Then
			Local $Html_SWEUserName = $oIE.Document.body.all("s_swepi_1")
			If IsObj($Html_SWEUserName) = 0 Then
				Return WinWait($IE_CLOE_Hwnd, "", $WinWaitDelai)
			Else
				With $oIE.Document.body
					if $NNISesame <> "" and $mdpSesame <> "" Then
						.all("s_swepi_1").Value = $NNISesame
						.all("s_swepi_2").Value = $mdpSesame
						$Html_s_swepi_22 = .all("s_swepi_22")
						$Html_s_swepi_22.click()
					EndIf
					WinWait($IE_CLOE_Hwnd, "", $WinWaitDelai)
				EndWith
			EndIf
		Else
			$Html_overridelink.click()
			WinWait($IE_CLOE_Hwnd, "Terminé", $WinWaitDelai)
		EndIf
	WEnd
EndFunc   ;==>OuvrirSessionCLOE

Func My_IECreate($s_Url = "about:blank", $f_visible = 1, $f_wait = 5)

	Local $IECLOE_Process = ShellExecute("iexplore.exe", $s_Url)

	$IE_CLOE_Hwnd = winwait("[Class:IEFrame]", "", $WinWaitDelai)

	If WinGetProcess($IE_CLOE_Hwnd) = $IECLOE_Process Then
		Return _IEAttach($IE_CLOE_Hwnd, "HWND")
	Else
		Return MySetError(1, 0, 0, "My_IECreate",@ScriptLineNumber , "")
	EndIf
EndFunc   ;==>My_IECreate


Func InitialiserCLOE($Adresse = "", $NNISesame = "", $mdpSesame = "")
	Opt("WinWaitDelay", 2000)

	;Input Adresse URL CLOE
	;Return le handle de  la  fênetre CLOE, si succès
	;Return 0 et @error = 1 si impossible  de recuperer un objet IE
	;Return 0 et @error = 2 si Connexion session est un échec
	;Return 0 et @error = 3 menu choix de l'objet CLOE introuvable

	FermerIE()
	$oIECLOE = My_IECreate($Adresse)
	If(@error > 0) Or($IE_CLOE_Hwnd = 0) Then Return SetError(1, 0, 0)

	$IE_CLOE_Hwnd = WinWait("[Class:IEFrame]", "", $WinWaitDelai)
	If $IE_CLOE_Hwnd = 0 Then Return SetError(1, 0, 0)

	If(OuvrirSessionCLOE($NNISesame, $mdpSesame) <> $IE_CLOE_Hwnd) Then Return SetError(2, 0, 0)

	Return $IE_CLOE_Hwnd
EndFunc   ;==>InitialiserCLOE

Func CLOEGetFieldHwnd()  ;($TitreFen, $ControlAscendantHwnd = 0, $ControlAdvancedClassName = "", $NbreCol= 4, $CadrePos = Default)
	CLOEGetZoneHwnd()
	Local $Array1D =  _ControlGetInstancesVisiblesByClassName($IE_CLOE_Hwnd, $IEServer_Hwnd,  "[REGEXPCLASS:Edit|Button]", False,  Default)

;~ 	Local $Array2D[1][$NbreCol], $i = 0, $j = 0
;~
;~ 	For $X = 0 to  UBound($Array1D) - 1
;~ 		If
;~ 	Next
EndFunc

#EndRegion LesFonctions ;****LesFonctions****²
;****LesFonctions****
