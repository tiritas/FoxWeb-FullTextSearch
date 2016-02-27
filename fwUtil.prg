* Constants used by the PushSetting and PopSetting methods
#DEFINE PUSH_EXACT_SUBSCRIPT 1
#DEFINE PUSH_ERROR_SUBSCRIPT 2
#DEFINE PUSH_FULLPATH_SUBSCRIPT 3
#DEFINE PUSH_COLLATE_SUBSCRIPT 4
#DEFINE PUSH_TOT_SETTINGS 4
#DEFINE PUSH_SETTING_DELIM '}){'

#DEFINE CR CHR(13)
#DEFINE LF CHR(10)
#DEFINE CRLF CHR(13) + CHR(10)

* Store the version of VFP that compiled the program to VFP_COMPILE_VERSION
#IF TYPE('VERSION(5)') = 'U'
	#IF VAL(SUBSTR(VERSION(), 15, AT('.', SUBSTR(VERSION(), 15), 1) -1)) = 3
		#DEFINE VFP_COMPILE_VERSION 3
	#ELIF VAL(SUBSTR(VERSION(), 15, AT('.', SUBSTR(VERSION(), 15), 1) -1)) = 5
		#DEFINE VFP_COMPILE_VERSION 5
	#ELSE
		* This should never happen
		#DEFINE VFP_COMPILE_VERSION -1
	#ENDIF
#ELSE
	#IF INT(VERSION(5)/100) = 6
		#DEFINE VFP_COMPILE_VERSION 6
	#ELIF INT(VERSION(5)/100) = 7
		#DEFINE VFP_COMPILE_VERSION 7
	#ELIF INT(VERSION(5)/100) = 8
		#DEFINE VFP_COMPILE_VERSION 8
	#ELIF INT(VERSION(5)/100) = 9
		#DEFINE VFP_COMPILE_VERSION 9
	#ELIF INT(VERSION(5)/100) = 10
		#DEFINE VFP_COMPILE_VERSION 10
	#ELSE
		* Unrecognized version
		#DEFINE VFP_COMPILE_VERSION -1
	#ENDIF
#ENDIF


**************************************************************
* The following is called with DO fwUtil.prg
**************************************************************
* Check if the object fwUtil exists
IF TYPE("M.fwUtil") != 'O'
	IF TYPE("M.fwUtil") = 'U'
		* Variable does not even exist, so it must be declared
		* public in order to survive after we exit this function
		PUBLIC fwUtil
	ENDIF
	fwUtil = CREATEOBJECT('classFwUtil')
ENDIF
RETURN



**************************************************************
* Definition of the classFwUtil class
**************************************************************
DEFINE CLASS classFwUtil AS Custom
	PROTECTED aSettingsStack(1), RegExp
	RegExp = .F.
	DIMENSION aSettingsStack(1)


***********************************************************
PROCEDURE Init
	IF VFP_COMPILE_VERSION <> INT(IIF(type('VERSION(5)') <> 'U', VERSION(5), INT(VAL(CHRTRAN(SUBSTR(VERSION(), 15, AT('.', SUBSTR(VERSION(), 15), 2) -1), '.', '')))) / 100)
		ERROR 'This class was compiled with VFP version ' + TRANSFORM(VFP_COMPILE_VERSION)
	ENDIF
	* Initialize variable that will hold setting info for the PushSetting and PopSetting methods
	LOCAL i
	DIMENSION THIS.aSettingsStack(PUSH_TOT_SETTINGS)
	FOR i = 1 TO PUSH_TOT_SETTINGS
		THIS.aSettingsStack(M.i) = ''
	NEXT
	THIS.RegExp = CREATEOBJECT("VBScript.RegExp")
ENDPROC


***********************************************************
FUNCTION RegExpReplace
* Uses regular expressions to replace a pattern with a string
	LPARAMETERS Original, Pattern, Replacement
	THIS.RegExp.Global = .T.
	THIS.RegExp.Multiline = .T.
	THIS.RegExp.IgnoreCase = .T.
	THIS.RegExp.Pattern = M.Pattern
	RETURN THIS.RegExp.Replace(M.Original, M.Replacement)
ENDFUNC


***********************************************************
FUNCTION RegExpTest
* Uses regular expressions to test a string against a pattern
	LPARAMETERS Original, Pattern
	THIS.RegExp.Global = .F.
	THIS.RegExp.Multiline = .F.
	THIS.RegExp.IgnoreCase = .F.
	THIS.RegExp.Pattern = M.Pattern
	THIS.RegExp.Pattern = M.Pattern
	RETURN THIS.RegExp.Test(M.Original)
ENDFUNC


***********************************************************
FUNCTION ParseURL
* Parses a URL into its parts
	LPARAMETERS strURL
	RETURN CREATEOBJECT('classURL', M.strURL, THIS)
ENDFUNC


***********************************************************
FUNCTION URLDecode
	LPARAMETERS TextString
	LOCAL oMatchCollection, oMatch
	THIS.RegExp.Global = .T.
	THIS.RegExp.Multiline = .T.
	THIS.RegExp.IgnoreCase = .T.
	THIS.RegExp.Pattern = "\+"
	M.TextString = THIS.RegExp.Replace(M.TextString, " ")
	THIS.RegExp.Pattern = "%[0-9,A-F]{2}"
	M.oMatchCollection = THIS.RegExp.Execute(M.TextString)
	FOR EACH oMatch IN oMatchCollection
		M.TextString = STRTRAN(M.TextString, oMatch.value, CHR(VAL("0x" + RIGHT(oMatch.Value,2))))
	NEXT
	RETURN M.TextString
ENDFUNC


***********************************************************
FUNCTION ValidateEmail
* Validates an email address
	LPARAMETERS email
	RETURN THIS.RegExpTest(m.email, "^[\w\.-]+@[a-zA-Z0-9\.-]+\.[a-zA-Z]+$")
ENDFUNC


***********************************************************
FUNCTION PadZero
	* Converts a number to a string padding it with zeros
	LPARAMETERS Value, TotSize
	RETURN RIGHT(REPLICATE("0", TotSize) + LTRIM(STR(M.Value)), M.TotSize)
ENDFUNC


***********************************************************
FUNCTION TrimChar
	* Strips surrounding characters from a string
	LPARAMETERS TextString, CharsToTrim
	LOCAL i
	* First trim on the left
	FOR M.i = 1 TO LEN(M.TextString)
		IF ! SUBSTR(M.TextString, M.i, 1) $ M.CharsToTrim
			EXIT
		ENDIF
	NEXT
	M.TextString = SUBSTR(M.TextString, M.i)
	* Now trim on the right
	FOR M.i = LEN(M.TextString) TO 1 STEP -1
		IF ! SUBSTR(M.TextString, M.i, 1) $ M.CharsToTrim
			EXIT
		ENDIF
	NEXT
	RETURN LEFT(M.TextString, M.i)
ENDFUNC


***********************************************************
FUNCTION JustTime
	* Returns the time portion of a DateTime expression as a string
	LPARAMETERS DateTimeExpr, military
	LOCAL AMPM, TimeExpr, nHour
	M.nHour = HOUR(M.DateTimeExpr)
	IF M.Military
		RETURN THIS.PadZero(M.nHour, 2) + ":" + THIS.PadZero(MINUTE(M.DateTimeExpr), 2)
	ELSE
		RETURN THIS.PadZero(IIF(M.nHour > 12, M.nHour - 12, M.nHour), 2) + ":" + THIS.PadZero(MINUTE(M.DateTimeExpr), 2) + " " + IIF(M.nHour >= 12, 'PM', 'AM')
	ENDIF
ENDFUNC


***********************************************************
PROCEDURE OpenTable
* Opens a table if it is not already open
	LPARAMETERS FilePath, Alias, Exclusive
	IF TYPE("M.alias") <> "C"
		M.alias = JustFName(JustStem(M.FilePath))
		IF '!' $ M.alias
			* Strip database delimiter
			M.alias = SUBSTR(M.alias, RAT('!', M.alias) + 1)
		ENDIF
	ENDIF
	IF USED(Alias)
		SELECT (Alias)
		IF FULLPATH(DBF()) <> FULLPATH(M.FilePath) OR (M.Exclusive <> ISEXCLUSIVE())
			IF M.Exclusive
				USE (M.FilePath) AGAIN ALIAS (M.Alias) EXCLUSIVE
			ELSE
				USE (M.FilePath) AGAIN ALIAS (M.Alias) SHARED
			ENDIF
		ENDIF
	ELSE
		SELECT 0
		IF M.Exclusive
			USE (M.FilePath) AGAIN ALIAS (M.Alias) EXCLUSIVE
		ELSE
			USE (M.FilePath) AGAIN ALIAS (M.Alias) SHARED
		ENDIF
	ENDIF
ENDPROC


***********************************************************
PROCEDURE AddDelimString
	LPARAMETERS NewString, OriginalString, delimiter
	DO CASE
	CASE EMPTY(M.NewString)
		* Do nothing
	CASE EMPTY(M.OriginalString)
		M.OriginalString = M.NewString
	OTHERWISE
		M.OriginalString = TRIM(M.OriginalString) + M.delimiter + LTRIM(M.NewString)
	ENDCASE
	RETURN M.OriginalString
ENDPROC


***********************************************************
FUNCTION FormatMailto
* Formats a mailto URL
	PARAMETERS User, UserEmail
	IF TYPE("M.User") <> "C"
		M.User = M.UserEmail
		IF TYPE("M.User") <> "C"
			RETURN ''
		ENDIF
	ENDIF
	IF TYPE("M.UserEmail") <> "C" OR EMPTY(M.UserEmail)
		RETURN M.User
	ELSE
		RETURN '<A HREF="mailto:' + M.UserEmail + '">' + M.User + '</A>'
	ENDIF
ENDFUNC


***********************************************************
FUNCTION GetDelimitedSet
* Returns a set of items from a delimited string
LPARAMETERS TextString, SetNumber, ItemsPerSet, SetDelim
RETURN SUBSTR(M.TextString, AT(M.SetDelim, M.SetDelim + M.TextString + M.SetDelim, ((M.SetNumber - 1) * M.ItemsPerSet) + 1), ;
	MIN(AT(M.SetDelim, M.SetDelim + M.TextString + REPLICATE(M.SetDelim,M.ItemsPerSet) , (M.SetNumber * M.ItemsPerSet) + 1) - ;
	AT(M.SetDelim, M.SetDelim + M.TextString + M.SetDelim, ((M.SetNumber - 1) * M.ItemsPerSet) + 1), LEN(M.TextString) + 1) - 1)
ENDFUNC


***********************************************************
FUNCTION PushSetting
* Stores current value of a setting in a stack and sets new value
	LPARAMETERS setting, NewValue
	LOCAL OldValue
	M.setting = UPPER(M.setting)
	DO CASE
	CASE M.setting = 'EXACT'
		OldValue = SET('EXACT')
		THIS.aSettingsStack(PUSH_EXACT_SUBSCRIPT) = THIS.aSettingsStack(PUSH_EXACT_SUBSCRIPT) + PUSH_SETTING_DELIM + M.OldValue
		SET EXACT &NewValue
	CASE M.setting = 'FULLPATH'
		OldValue = SET('FULLPATH')
		THIS.aSettingsStack(PUSH_FULLPATH_SUBSCRIPT) = THIS.aSettingsStack(PUSH_FULLPATH_SUBSCRIPT) + PUSH_SETTING_DELIM + M.OldValue
		SET EXACT &NewValue
	CASE M.setting = 'COLLATE'
		OldValue = SET('COLLATE')
		THIS.aSettingsStack(PUSH_COLLATE_SUBSCRIPT) = THIS.aSettingsStack(PUSH_COLLATE_SUBSCRIPT) + PUSH_SETTING_DELIM + M.OldValue
		SET COLLATE TO (M.NewValue)
	CASE M.setting = 'ERROR'
		OldValue = ON('ERROR')
		THIS.aSettingsStack(PUSH_ERROR_SUBSCRIPT) = THIS.aSettingsStack(PUSH_ERROR_SUBSCRIPT) + PUSH_SETTING_DELIM + M.OldValue
		ON ERROR &NewValue
	OTHERWISE
		ERROR 'Unsupported setting argument: "' + M.setting + '"'
	ENDCASE
	RETURN
ENDFUNC


***********************************************************
FUNCTION PopSetting
* Retrieves and sets previously saved value from the stack (LIFO)
	LPARAMETERS setting
	LOCAL LastSettingPos, NewValue
	M.setting = UPPER(M.setting)
	NewValue = ''
	DO CASE
	CASE M.setting = 'EXACT'
		M.NewValue = THIS.ReadSettingFromStack(PUSH_EXACT_SUBSCRIPT)
		IF ISNULL(M.NewValue)
			M.NewValue = SET('EXACT')
		ELSE
			SET EXACT &NewValue
		ENDIF
	CASE M.setting = 'FULLPATH'
		M.NewValue = THIS.ReadSettingFromStack(PUSH_FULLPATH_SUBSCRIPT)
		IF ISNULL(M.NewValue)
			M.NewValue = SET('FULLPATH')
		ELSE
			SET FULLPATH &NewValue
		ENDIF
	CASE M.setting = 'COLLATE'
		M.NewValue = THIS.ReadSettingFromStack(PUSH_COLLATE_SUBSCRIPT)
		IF ISNULL(M.NewValue)
			M.NewValue = SET('COLLATE')
		ELSE
			SET COLLATE TO (M.NewValue)
		ENDIF
	CASE M.setting = 'ERROR'
		M.NewValue = THIS.ReadSettingFromStack(PUSH_ERROR_SUBSCRIPT)
		IF ISNULL(M.NewValue)
			M.NewValue = ON('ERROR')
		ELSE
			ON ERROR &NewValue
		ENDIF
	OTHERWISE
		ERROR 'Unsupported setting argument: "' + M.setting + '"'
	ENDCASE
	RETURN M.NewValue
ENDFUNC


***********************************************************
PROTECTED FUNCTION ReadSettingFromStack
* Reads and deletes previously saved value from the stack
	LPARAMETERS SettingSubScript
	LOCAL NewValue
	M.LastSettingPos = RAT(PUSH_SETTING_DELIM, THIS.aSettingsStack(M.SettingSubScript))
	IF M.LastSettingPos = 0
		* Setting not found in stack
		M.NewValue = .NULL.
	ELSE
		* Read and delete setting
		M.NewValue = SUBSTR(THIS.aSettingsStack(M.SettingSubScript), M.LastSettingPos + LEN(PUSH_SETTING_DELIM))
		THIS.aSettingsStack(M.SettingSubScript) = LEFT(THIS.aSettingsStack(M.SettingSubScript), M.LastSettingPos - 1)
	ENDIF
	RETURN M.NewValue
ENDFUNC


**********************************************************
FUNCTION Split
* Splits a delimited string into an array
	LPARAMETERS SourceString, aElements, Delim
	IF TYPE("M.Delim") <> 'C'
		* No delimiter was passed.  Use comma by default
		M.Delim = ","
	ENDIF
	LOCAL TotElements, CurElement
	M.TotElements = THIS.GetWordCount(m.SourceString, M.Delim)
	DIMENSION M.aElements(M.TotElements)
	FOR M.CurElement = 1 TO M.TotElements
		M.aElements(M.CurElement) = ALLTRIM(THIS.GetWordNum(M.SourceString, M.CurElement, M.Delim))
	NEXT
	RETURN M.TotElements
ENDFUNC


**********************************************************
FUNCTION GetWordCount
* Simulates VFP 7's GETWORDCOUNT function
	LPARAMETERS M.SourceString, M.Delim

	#IF VFP_COMPILE_VERSION < 7
		THIS.GetWordFixString(@M.SourceString, @M.Delim)
		RETURN OCCURS(M.Delim, M.SourceString) + 1
	#ELSE
		IF PCOUNT() > 1
			RETURN GETWORDCOUNT(M.SourceString, M.Delim)
		ELSE
			RETURN GETWORDCOUNT(M.SourceString)
		ENDIF
	#ENDIF
ENDFUNC


**********************************************************
FUNCTION GetWordNum
* Simulates VFP 7's GETWORDNUM function
	LPARAMETERS M.SourceString, M.ElementNum, M.Delim

	#IF VFP_COMPILE_VERSION < 7
		LOCAL DelimLen, StartPos
		THIS.GetWordFixString(@M.SourceString, @M.Delim)
		M.DelimLen = LEN(M.Delim)
		M.StartPos = AT(M.Delim, M.Delim + M.SourceString, M.ElementNum)
		RETURN SUBSTR(M.SourceString, M.StartPos, AT(M.Delim, M.SourceString + M.Delim, M.ElementNum) - M.StartPos)
	#ELSE
		IF VARTYPE(M.Delim) == 'C'
			RETURN GETWORDNUM(M.SourceString, M.ElementNum, M.Delim)
		ELSE
			RETURN GETWORDNUM(M.SourceString, M.ElementNum)
		ENDIF
	#ENDIF
ENDFUNC


**********************************************************
PROTECTED FUNCTION GetWordFixString
* Prepares SourceString for GetWordCount and GetWordNum functions
	LPARAMETERS M.SourceString, M.Delim
	IF TYPE('M.Delim') <> 'C'
		* The default behavior uses spaces and tabs as delimiters if no delimiter is specified
		M.Delim = ' '
		* Replace tabs with spaces
		M.SourceString = CHRTRAN(M.SourceString, CHR(9), ' ')
	ENDIF
	* Remove consecutive instances of a delimiter
	DO WHILE M.Delim + M.Delim $ M.SourceString
		M.SourceString = STRTRAN(M.SourceString, M.Delim + M.Delim, M.Delim)
	ENDDO
	* Remove leading delimiter
	IF LEFT(M.SourceString, 1) = M.Delim
		M.SourceString = SUBSTR(M.SourceString, 2)
	ENDIF
	* Remove trailing delimiter
	IF RIGHT(M.SourceString, 1) = M.Delim
		M.SourceString = LEFT(M.SourceString, LEN(M.SourceString) - 1)
	ENDIF
ENDFUNC


**********************************************************
FUNCTION ToString
* Converts a variant of type C, D, T, M, N, L to a string
* or returns an empty string
**********************************************************
LPARAMETERS vInput
	LOCAL VarType
	#IF VFP_COMPILE_VERSION < 6
		M.VarType = TYPE("M.vInput")
		DO CASE
		CASE M.VarType $ "CM"
			RETURN M.vInput
		CASE M.VarType $ "NY"
			IF M.vInput = 0
				RETURN '0'
			ELSE
				LOCAL StrNum, StrExp
				StrNum = LTRIM(STR(M.vInput, 25, 18))
				IF 'E' $ M.StrNum
					StrExp = SUBSTR(M.StrNum, RAT('E', M.StrNum))
					StrNum = LEFT(M.StrNum, RAT('E', M.StrNum) - 1)
				ELSE
					StrExp = ''
				ENDIF
				IF '.' $ M.StrNum
					* Strip zeros past the decimal point
					DO WHILE RIGHT(M.StrNum, 1) == '0'
						StrNum = LEFT(M.StrNum, LEN(M.StrNum) - 1)
					ENDDO
					* Strip decimal point if it is at the end
					IF RIGHT(M.StrNum, 1) == '.'
						StrNum = LEFT(M.StrNum, LEN(M.StrNum) - 1)
					ENDIF
				ENDIF
				RETURN M.StrNum + M.StrExp
			ENDIF
		CASE M.VarType = "D"
			RETURN IIF(EMPTY(M.vInput), '', DTOC(M.vInput))
		CASE M.VarType = "T"
			RETURN IIF(EMPTY(M.vInput), '', TTOC(M.vInput))
		CASE M.VarType = "L"
			RETURN IIF(M.vInput, ".T.", ".F.")
		OTHERWISE
			RETURN ""
		ENDCASE
	#ELSE
		M.VarType = VARTYPE(M.vInput)
		DO CASE
		CASE M.VarType $ 'CM'
			RETURN M.vInput
		CASE M.VarType $ 'NYDTL'
			RETURN TRANSFORM(M.vInput)
		OTHERWISE
			RETURN ''
		ENDCASE
	#ENDIF
ENDFUNC

ENDDEFINE
**********************************************************
*       END OF DEFINITION OF classFwUtil CLASS
**********************************************************


**************************************************************
* Definition of the classURL class
**************************************************************
DEFINE CLASS classURL AS Custom
	PROTECTED aSettingsStack(1), RegExp
	RegExp = .F.
	Scheme = ''
	Authority = ''
	Path = ''
	QueryString = ''
	QueryStringParts = .F.
	Fragment = ''
	Valid = .F.

	PROTECTED PROCEDURE Init
		LPARAMETERS M.URL, fwUtil
		LOCAL Matches, QueryStringPart, QueryStringPartNum, QueryStringKey, QueryStringValue
		THIS.QueryStringParts = CREATEOBJECT("classFWCollection")
		THIS.RegExp = CREATEOBJECT("VBScript.RegExp")
		THIS.RegExp.Global = .F.
		THIS.RegExp.Multiline = .F.
		THIS.RegExp.IgnoreCase = .T.
		THIS.RegExp.Pattern = "^(?:([^:/\?#]+):)?(?://([^/\?#]*))?([^\?#]*)(?:\?([^#]*))?(?:#(.*))?"
		* First make sure that this is indeed a URL
		IF THIS.RegExp.Test(M.URL) Then
			* Parse the URL into its parts
			M.Matches = THIS.RegExp.Execute(M.URL)
			THIS.Scheme = NVL(M.Matches.Item(0).SubMatches(0), "")
			THIS.Authority = NVL(M.Matches.Item(0).SubMatches(1), "")
			THIS.Path = NVL(M.Matches.Item(0).SubMatches(2), "")
			THIS.QueryString = NVL(M.Matches.Item(0).SubMatches(3), "")
			THIS.Fragment = NVL(M.Matches.Item(0).SubMatches(4), "")
			IF NOT EMPTY(THIS.QueryString)
				FOR M.QueryStringPartNum = 1 TO GETWORDCOUNT(THIS.QueryString, "&")
					M.QueryStringPart = GETWORDNUM(THIS.QueryString, M.QueryStringPartNum, "&")
					IF "=" $ M.QueryStringPart
						M.QueryStringKey = LOWER(fwUtil.URLDecode(LEFT(M.QueryStringPart, AT("=", M.QueryStringPart) - 1)))
						M.QueryStringValue = fwUtil.URLDecode(SUBSTR(M.QueryStringPart, AT("=", M.QueryStringPart) + 1))
					ELSE
						M.QueryStringKey = LOWER(fwUtil.URLDecode(M.QueryStringPart))
						M.QueryStringValue = ""
					ENDIF
					IF THIS.QueryStringParts.Exists(M.QueryStringKey)
						* An item with the same key already exists
						* It seems like we can't edit an existing collection item,
						* so we are removing it first and re-adding below.
						M.QueryStringValue = THIS.QueryStringParts.Item(M.QueryStringKey) + "," + M.QueryStringValue
						THIS.QueryStringParts.Remove(M.QueryStringKey)
					ENDIF
					THIS.QueryStringParts.Add(M.QueryStringValue, M.QueryStringKey)
				NEXT
			ENDIF
			THIS.Valid = .T.
		ELSE
			THIS.Valid = .F.
		ENDIF
	ENDPROC
ENDDEFINE
**********************************************************
*       END OF DEFINITION OF classURL CLASS
**********************************************************


**************************************************************
* Definition of the classURL class
**************************************************************
DEFINE CLASS classFWCollection AS Collection

	FUNCTION Exists
		LPARAMETERS cKey
		LOCAL bResult
		M.bResult = .T.
		TRY
			=THIS.Item(M.cKey)
		CATCH
			M.bResult = .F.
		ENDTRY
		RETURN M.bResult
	ENDFUNC

	PROCEDURE RemoveAll
		LOCAL i
		FOR i = 1 TO THIS.Count
			THIS.Remove(1)
		NEXT
	ENDPROC
ENDDEFINE
**********************************************************
*       END OF DEFINITION OF classFWCollection CLASS
**********************************************************
