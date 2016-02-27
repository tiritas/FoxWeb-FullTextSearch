		****************************************************
		* ************************************************ *
		* *   fwFullText.PRG                             * *
		* *   (C) Aegis Group 2015                       * *
		* ************************************************ *
		****************************************************

#DEFINE COMPONENT_VERSION 2.8

#DEFINE MAX_KEYWORD_LEN 20
#DEFINE SEARCH_TERM_DELIMS	' ,'	&& List of characters that can separate search terms -- other characters cause words to be joined together.
#DEFINE MAX_ITEMS_PER_IN_CLAUSE 10
* Constants used by the Errors and Error objects
#DEFINE ERROR_LEVEL_INFO 0
#DEFINE ERROR_LEVEL_WARNING 1
#DEFINE ERROR_LEVEL_ERROR 2

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
* Definition of the classFullText class
**************************************************************
DEFINE CLASS classFullText AS Custom
	PROTECTED IndexTable, InFoxWeb, KeyField, KeywordIDSize, NoiseWords, Proximity, ;
			RejectedSearchWords, TablePath, TableName, TextFields(1), TotTextFields
	Errors = .NULL.
	ExclusiveIndexing = .F.
	fwUtil = .NULL.
	DIMENSION HighlightPatterns(1)
	InFoxWeb = .F.
	KeyField = ''
	KeywordIDSize = 2
	LastCorrectedSearchPhrase = ''
	LastSearchTally = 0
	MaxHits = 200
	MaxHitsPercent = 50
	NoiseWords = ''
	Proximity = .T.
	RejectedSearchWords = ''
	Table = ''
	TablePath = ''
	TableName = ''
	DIMENSION TextFields(1)
	TotTextFields = 0


**********************************************************
PROTECTED PROCEDURE Init
**********************************************************
	IF VFP_COMPILE_VERSION <> INT(IIF(TYPE('VERSION(5)') <> 'U', VERSION(5), INT(VAL(CHRTRAN(SUBSTR(VERSION(), 15, AT('.', SUBSTR(VERSION(), 15), 2) -1), '.', '')))) / 100)
		ERROR 'This class was compiled with VFP ' + TRANSFORM(VFP_COMPILE_VERSION) + ', but you are calling it from VFP ' + TRANSFORM(INT(IIF(TYPE('VERSION(5)') <> 'U', VERSION(5), INT(VAL(CHRTRAN(SUBSTR(VERSION(), 15, AT('.', SUBSTR(VERSION(), 15), 2) -1), '.', '')))) / 100))
	ENDIF
	IF TYPE("M.fwUtil") = "O" AND LOWER(fwUtil.Class) = 'classfwutil'
		THIS.fwUtil = M.fwUtil
	ELSE
		LOCAL fwUtilModule
		M.fwUtilModule = 'fwUtil.fxp'
		THIS.fwUtil = NEWOBJECT('classFwUtil', M.fwUtilModule)
	ENDIF
	THIS.Errors = CREATEOBJECT('classErrors', THIS.Name + ' object')
	* Are we running withing FoxWeb?
	THIS.InFoxWeb = (TYPE("Server.ScriptTimeout") = 'N')
ENDPROC


**********************************************************
PROTECTED PROCEDURE Destroy
**********************************************************
	* Clean up
	THIS.fwUtil = .NULL.
	THIS.Errors = .NULL.
	CLEAR CLASS classFwUtil
	CLEAR CLASS classErrors
	CLEAR CLASS classError
ENDPROC


**********************************************************
PROCEDURE Search
**********************************************************
	LPARAMETERS SearchPhrase, ResultCursorName, SearchFieldList, AnyWords
	LOCAL SearchQuery, TotKeywords, WordPos, CurKeyword, CurKeywordNum, ;
		TotSourceRecords, TotGoodTerms, SearchFieldIdList, ;
		TotSearchTerms, CurSearchTerm, aSearchTerms, aKeywords, ;
		HasMultipleTextFields, TotKeywordSets, ;
		CurKeywordSet, inClassFullText, RunTimeError, aBinaryIDs

	EXTERNAL ARRAY aKeywords
	M.inClassFullText = .T.
	THIS.LastSearchTally = 0
	THIS.LastCorrectedSearchPhrase = ''
	THIS.RejectedSearchWords = ''
	DIMENSION THIS.HighlightPatterns(1)
	THIS.HighlightPatterns(1) = ''
	* Clear messages
	THIS.Errors.Clear()
	IF PCOUNT() < 2
		ERROR 'Invalid number of parameters'
		RETURN .F.
	ENDIF
	IF TYPE('M.SearchPhrase') <> 'C'
		ERROR 'Invalid searh string'
		RETURN .F.
	ENDIF
	IF TYPE('M.ResultCursorName') <> 'C'
		ERROR 'Invalid cursor name'
		RETURN .F.
	ENDIF
	IF PCOUNT() >= 4 AND TYPE('M.AnyWords') <> 'L'
		ERROR 'Invalid AnyWords argument'
		RETURN .F.
	ENDIF
	IF EMPTY(THIS.Table)
		THIS.Errors.Add('You must populate the Table property before calling this function', , ERROR_LEVEL_ERROR, THIS.Name + '.UpdateIndex')
		RETURN .F.
	ENDIF
	M.ResultCursorName = ALLTRIM(M.ResultCursorName)
	* Verify that the table has a full-text index
	IF ! THIS.IsIndexed()
		THIS.Errors.Add('Non-existent or invalid Full Text Index for table "' + THIS.Table + '"', , ERROR_LEVEL_ERROR, THIS.Name + '.Search')
		RETURN .F.
	ENDIF
	M.TotGoodTerms = 0
	IF ! THIS.OpenSourceTable() OR ! THIS.OpenIndexTables()
		RETURN .F.
	ENDIF
	SELECT (THIS.TableName)
	COUNT FOR ! DELETED() TO M.TotSourceRecords
	* Determine whether the selected table has multiple indexed fields
	M.HasMultipleTextFields = THIS.TotTextFields > 1
	* Read text field(s)
	IF TYPE("M.SearchFieldList") = "C"
		M.SearchFieldIdList = THIS.GetSearchFieldIdList(ALLTRIM(LOWER(M.SearchFieldList)), @M.aBinaryIDs)
		IF EMPTY(M.SearchFieldIdList)
			* GetSearchFieldIdList failed for some reason
			RETURN .F.
		ENDIF
	ELSE
		* No text field specified.  Search all fields
		M.SearchFieldIdList = ""
	ENDIF
	M.SearchPhrase = ALLTRIM(M.SearchPhrase)
	IF ! THIS.Proximity
		* Test whether the user entered a search phrase that requires proximity searches
		IF ! THIS.fwUtil.RegExpTest(M.SearchPhrase, '[A-Za-z0-9_\*\?' + SEARCH_TERM_DELIMS + '].')
			THIS.Errors.Add('Proximity searches are not supported by this index', , ERROR_LEVEL_WARNING, THIS.Name + '.Search')
		ENDIF
		* Index does not support proximity searches -- remove all extranneous characters from search phrase
		M.SearchPhrase = THIS.SearchTermSplit(M.SearchPhrase)
	ENDIF
	* Split search string into individual search terms (each term may contain words, phrases, or composite words)
	M.TotSearchTerms = THIS.ParseSearchPhrase(M.SearchPhrase, @M.aSearchTerms)
	* Construct SQL statement
	M.SearchQuery = ''
	FOR M.CurSearchTerm = 1 TO M.TotSearchTerms
		M.TotKeywords = THIS.GetKeywordIDs(M.aSearchTerms(M.CurSearchTerm, 2), @M.aKeywords, @M.aBinaryIDs)
		DO CASE
		CASE M.TotKeywords = 0
			* All words in Noise list, or length not between 2 and MAX_KEYWORD_LEN
			LOOP
		CASE M.TotKeywords = -1
			* At least one of the keywords was not found in the index
			* Search failed
			THIS.CloseIndexTables()
			RETURN .T.
		ENDCASE
		M.TotGoodTerms = M.TotGoodTerms + 1
		IF M.TotGoodTerms > 1
			M.SearchQuery = M.SearchQuery + ' UNION '
		ENDIF
		M.SearchQuery = M.SearchQuery + 'SELECT "' + PADR(M.aSearchTerms(M.CurSearchTerm, 2), MAX_KEYWORD_LEN) + '" CurTerm, ft1.indexvalue, '
		IF THIS.Proximity
			M.SearchQuery = M.SearchQuery + 'COUNT(*) WFrequency FROM FTIndex ft1 '
		ELSE
			M.SearchQuery = M.SearchQuery + 'SUM(CTOBIN(WordCount)) WFrequency FROM FTIndex ft1 '
		ENDIF
		FOR M.CurKeywordNum = 2 TO M.TotKeywords
			M.SearchQuery = M.SearchQuery + ' JOIN FTIndex ft' + TRANSFORM(M.CurKeywordNum) + ;
				' ON ft' + TRANSFORM(M.CurKeywordNum - 1) + '.indexvalue = ft' + TRANSFORM(M.CurKeywordNum) + '.indexvalue ' + ;
				' AND CTOBIN(ft' + TRANSFORM(M.CurKeywordNum) + '.wordpos) = CTOBIN(ft' + TRANSFORM(M.CurKeywordNum - 1) + '.wordpos) + ' + TRANSFORM(M.aKeywords(M.CurKeywordNum, 3)) + ' '
			IF M.HasMultipleTextFields
				M.SearchQuery = M.SearchQuery + ' AND ft' + TRANSFORM(M.CurKeywordNum - 1) + '.TextField = ft' + TRANSFORM(M.CurKeywordNum) + '.TextField '
			ENDIF
		NEXT
		M.SearchQuery = M.SearchQuery + ' WHERE '
		FOR M.CurKeywordNum = 1 TO M.TotKeywords
			IF M.CurKeywordNum > 1
				M.SearchQuery = M.SearchQuery + ' AND '
			ENDIF
			* Use an IN clause to find all words matched by wildcard searches
			M.SearchQuery = M.SearchQuery + ' ('
			M.TotKeywordSets = CEILING((OCCURS(',', M.aKeywords(M.CurKeywordNum, 2)) + 1) / MAX_ITEMS_PER_IN_CLAUSE)
			FOR M.CurKeywordSet = 1 TO M.TotKeywordSets
				* Split words matched by wildcard searches into sets of MAX_ITEMS_PER_IN_CLAUSE
				IF M.CurKeywordSet > 1
					M.SearchQuery = M.SearchQuery + ' OR '
				ENDIF
				M.SearchQuery = M.SearchQuery + 'ft' + TRANSFORM(M.CurKeywordNum) + '.KeywordID IN (' + THIS.fwUtil.GetDelimitedSet(M.aKeywords(M.CurKeywordNum, 2), M.CurKeywordSet, MAX_ITEMS_PER_IN_CLAUSE, ',') + ') '
			NEXT
			M.SearchQuery = M.SearchQuery + ') '
		NEXT
		IF M.HasMultipleTextFields AND ! EMPTY(M.SearchFieldIdList)
			M.SearchQuery = M.SearchQuery + ' AND ft1.TextField IN (' + M.SearchFieldIdList + ') '
		ENDIF
		IF THIS.Proximity
			M.SearchQuery = M.SearchQuery + ' GROUP BY CurTerm, ft1.indexvalue '
		ELSE
			M.SearchQuery = M.SearchQuery + ' GROUP BY ft1.indexvalue '
		ENDIF
	NEXT
	IF ! EMPTY(THIS.RejectedSearchWords)
		THIS.Errors.Add('The following word(s) were in the noise list, so they were skipped: ' + THIS.RejectedSearchWords + '.', , ERROR_LEVEL_INFO, THIS.Name + '.Search')
	ENDIF
	IF M.TotGoodTerms = 0
		THIS.Errors.Add('Invalid search criteria: All keywords were in noise list', , ERROR_LEVEL_INFO, THIS.Name + '.Search')
		THIS.CloseIndexTables()
		RETURN .T.
	ENDIF
	M.SearchQuery = M.SearchQuery + ' INTO CURSOR Search1'
	* Trap errors during SQL statement execution
	THIS.fwUtil.PushSetting('ERROR', 'RunTimeError = ERROR()')
	THIS.fwUtil.PushSetting('COLLATE', 'MACHINE')
	RunTimeError = -999
	&SearchQuery
	THIS.fwUtil.PopSetting('COLLATE')
	THIS.fwUtil.PopSetting('ERROR')
	IF M.RunTimeError <> -999
		* There was an error during execution of the select statement
		THIS.Errors.Add(MESSAGE(), M.RunTimeError, ERROR_LEVEL_ERROR, THIS.Name + '.Search (SQL)')
		THIS.CloseIndexTables()
		RETURN .F.
	ENDIF
	M.SearchQuery = 'SELECT IndexValue, COUNT(*) AS TotWords, SUM(WFrequency) AS Frequency ' + ;
			' FROM Search1 GROUP BY IndexValue '
	IF ! M.AnyWords
		M.SearchQuery = M.SearchQuery + ' HAVING TotWords = ' + TRANSFORM(M.TotGoodTerms)
	ENDIF
	M.SearchQuery = M.SearchQuery + ' ORDER BY TotWords DESC, Frequency DESC, IndexValue ' + ;
			' INTO CURSOR ' + M.ResultCursorName
	&SearchQuery
	THIS.LastSearchTally = _TALLY
	USE IN Search1
	THIS.CloseIndexTables()
	DO CASE
	CASE THIS.LastSearchTally > THIS.MaxHits
		THIS.Errors.Add('Your search yielded ' + TRANSFORM(THIS.LastSearchTally) + ' hits.  Please make your criteria more restrictive.', , ERROR_LEVEL_INFO, THIS.Name + '.Search')
		THIS.LastSearchTally = 0
		USE
	CASE THIS.LastSearchTally / M.TotSourceRecords > THIS.MaxHitsPercent * 100
		THIS.Errors.Add('Your search matched ' + TRANSFORM(THIS.LastSearchTally / M.TotSourceRecords * 100) + '% of the total records.  Please make your criteria more restrictive.', , ERROR_LEVEL_INFO, THIS.Name + '.Search')
		THIS.LastSearchTally = 0
		USE
	CASE THIS.LastSearchTally > 0
		* We had some results -- populate HighlightPatterns
		DIMENSION THIS.HighlightPatterns(M.TotSearchTerms)
		FOR M.CurSearchTerm = 1 TO M.TotSearchTerms
			THIS.HighlightPatterns(M.CurSearchTerm) = '\b(' + STRTRAN(STRTRAN(STRTRAN(aSearchTerms(M.CurSearchTerm, 2), '*', '\w*'), '?', '\w?'), ' ', '\W+') + ')\b'
		NEXT
	ENDCASE
	RETURN .T.
ENDPROC


***********************************************************
PROTECTED FUNCTION ParseSearchPhrase
* Splits SearchPhrase into the individual search terms (either words, or complete phrases)
	LPARAMETERS SearchPhrase, aSearchTerms
	LOCAL StringPos, Quoted, TotSearchTerms, PrevDelim, CurSearchTerm
	M.SearchPhrase = ALLTRIM(M.SearchPhrase)
	M.Quoted = .F.
	M.TotSearchTerms = 0
	M.PrevDelim = 0
	FOR M.StringPos = 1 TO LEN(M.SearchPhrase)
		M.CurChar = SUBSTR(M.SearchPhrase, M.StringPos, 1)
		DO CASE
		CASE M.CurChar = '"'
			THIS.ProcessSearchTerm(@M.SearchPhrase, @M.StringPos, @M.PrevDelim, @M.aSearchTerms, @M.TotSearchTerms, M.Quoted)
			M.Quoted = ! M.Quoted
		CASE M.CurChar $ SEARCH_TERM_DELIMS AND ! M.Quoted
			THIS.ProcessSearchTerm(@M.SearchPhrase, @M.StringPos, @M.PrevDelim, @M.aSearchTerms, @M.TotSearchTerms, M.Quoted)
		ENDCASE
	NEXT
	IF M.StringPos > M.PrevDelim
		THIS.ProcessSearchTerm(@M.SearchPhrase, @M.StringPos, @M.PrevDelim, @M.aSearchTerms, @M.TotSearchTerms, M.Quoted)
	ENDIF
	* Populate LastCorrectedSearchPhrase, by appending the various corrected search terms
	FOR M.CurSearchTerm = 1 TO M.TotSearchTerms
		THIS.LastCorrectedSearchPhrase = THIS.LastCorrectedSearchPhrase + ;
			IIF(M.CurSearchTerm > 1, ' ', '') + ;
			IIF(M.aSearchTerms(M.CurSearchTerm, 3), '"', '') + ;
			M.aSearchTerms(M.CurSearchTerm, 1) + ;
			IIF(M.aSearchTerms(M.CurSearchTerm, 3), '"', '')
	NEXT
	RETURN M.TotSearchTerms
ENDFUNC


***********************************************************
PROTECTED FUNCTION ProcessSearchTerm
* Used by ParseSearchPhrase to store a search term into aSearchTerms array
	LPARAMETERS SearchPhrase, StringPos, PrevDelim, aSearchTerms, TotSearchTerms, Quoted
	LOCAL SearchTerm, ProcessedSearchTerm, CurSearchTerm, DuplicateSearchTerm
	M.SearchTerm = THIS.fwUtil.TrimChar(SUBSTR(M.SearchPhrase, M.PrevDelim + 1, M.StringPos - M.PrevDelim - 1), SEARCH_TERM_DELIMS)
	M.ProcessedSearchTerm = M.SearchTerm
	* Strip any surrounding quotes
	IF LEFT(M.ProcessedSearchTerm, 1) = '"'
		M.SearcProcessedSearchTerm = SUBSTR(M.ProcessedSearchTerm , 2)
	ENDIF
	IF RIGHT(ProcessedSearchTerm , 1) = '"'
		M.ProcessedSearchTerm = LEFT(M.ProcessedSearchTerm , LEN(M.ProcessedSearchTerm ) -1)
	ENDIF
	M.ProcessedSearchTerm = ALLTRIM(THIS.SearchTermSplit(M.ProcessedSearchTerm))
	IF ! EMPTY(M.ProcessedSearchTerm)
		* Check if term was already seen before
		M.DuplicateSearchTerm = .F.
		FOR M.CurSearchTerm = 1 TO M.TotSearchTerms
			IF aSearchTerms(M.CurSearchTerm, 2) == M.ProcessedSearchTerm
				M.DuplicateSearchTerm = .T.
				EXIT
			ENDIF
		NEXT
		IF ! M.DuplicateSearchTerm
			* Store search term in array
			M.TotSearchTerms = M.TotSearchTerms + 1
			DIMENSION aSearchTerms(M.TotSearchTerms, 3)
			M.aSearchTerms(M.TotSearchTerms, 1) = M.SearchTerm
			M.aSearchTerms(M.TotSearchTerms, 2) = M.ProcessedSearchTerm
			M.aSearchTerms(M.TotSearchTerms, 3) = M.Quoted
		ENDIF
	ENDIF
	M.PrevDelim = M.StringPos
ENDFUNC


***********************************************************
PROTECTED FUNCTION GetKeywordIDs
* Creates an array with the Keyword IDs of each keyword in string
	LPARAMETERS CurSearchTerm, aKeywordIDs, aBinaryIDs
	LOCAL TotKeywords, KeywordPos, CurKeyword, TotGoodKeywords,;
		TotBinaryIDs, CurSQLKeyword

	M.TotGoodKeywords = 0	&& These are keywords that are not in the noise list
	M.TotKeywords = THIS.FWUtil.GetWordCount(M.CurSearchTerm)
	IF TYPE('aBinaryIDs') = 'L'
		M.TotBinaryIDs = 0
	ELSE
		M.TotBinaryIDs = ALEN(M.aBinaryIDs)
	ENDIF
	* Iterate through search words in the current search term
	M.KeywordOffset = 1
	FOR M.KeywordPos = 1 TO M.TotKeywords
		M.CurKeyword = LOWER(LEFT(THIS.FWUtil.GetWordNum(M.CurSearchTerm, M.KeywordPos), MAX_KEYWORD_LEN))
		IF THIS.IsNoiseWord(M.CurKeyword)
			* This is a noise word -- ignore
			* With wild-card searches it's possible that other words are also matched, but we can't use them, 
			* because this would prevent our search from finding records that would be matched based on a noise word
			THIS.RejectedSearchWords = THIS.fwUtil.AddDelimString("'" + M.CurKeyword + "'", THIS.RejectedSearchWords, ', ')
			M.KeywordOffset = M.KeywordOffset + 1
		ELSE
			M.TotGoodKeywords = M.TotGoodKeywords + 1
			DIMENSION aKeywordIDs(M.TotGoodKeywords, 3)
			M.aKeywordIDs(M.TotGoodKeywords, 1) = TRIM(M.CurKeyword)
			M.aKeywordIDs(M.TotGoodKeywords, 2) = ''
			M.aKeywordIDs(M.TotGoodKeywords, 3) = M.KeywordOffset
			M.KeywordOffset = 1
			M.CurSQLKeyword = CHRTRAN(M.CurKeyword, '*?', '%_')
			SELECT KeywordID FROM FTWordIndex WHERE keyword LIKE M.CurSQLKeyword INTO CURSOR KeywordIDs
			IF _TALLY > 0
				* Scan through all words matching current search word (there may be more than one, if the search word contains wildcards)
				SCAN
					* Add the actual binary ID to the aBinaryIDs array
					* We are using aBinaryIDs because binary characters cannot be part of the SQL statement
					M.TotBinaryIDs = M.TotBinaryIDs + 1
					DIMENSION aBinaryIDs(M.TotBinaryIDs)
					M.aBinaryIDs(M.TotBinaryIDs) = KeywordIDs.KeywordID
					* Now add a reference to the id in the aBinaryIDs array into the comma-delimited list of IDs
					* This list will be used to construct the SQL statement later
					M.aKeywordIDs(M.TotGoodKeywords, 2) = THIS.fwUtil.AddDelimString('M.aBinaryIDs(' + TRANSFORM(M.TotBinaryIDs) + ')', aKeywordIDs(M.TotGoodKeywords, 2), ',')
				ENDSCAN
				USE
			ELSE
				* The current search word was not found in the database
				RETURN -1
			ENDIF
		ENDIF
	NEXT
	RETURN M.TotGoodKeywords
ENDFUNC


***********************************************************
PROTECTED FUNCTION SearchTermSplit
* Splits SearchTerm into words
	LPARAMETERS SearchTerm
	* Strip non-word characters, except asterisks and underscores
	M.SearchTerm = THIS.fwUtil.RegExpReplace(M.SearchTerm, '[^A-Za-z0-9_\*\?]', ' ')
	RETURN M.SearchTerm
ENDFUNC


**********************************************************
PROTECTED FUNCTION GetSearchFieldIdList
* Convert array with search field names to comma delimited string of field IDs
	LPARAMETERS SearchFieldList, aBinaryIDs
	LOCAL SearchFieldIdList, CurSearchField, CurSearchFieldNum, CurTextFieldNum, ;
			FieldFound, TotBinaryIDs
	IF THIS.TotTextFields <= 0
		RETURN ''
	ENDIF
	IF TYPE('aBinaryIDs') = 'L'
		M.TotBinaryIDs = 0
	ELSE
		M.TotBinaryIDs = ALEN(M.aBinaryIDs)
	ENDIF
	M.SearchFieldIdList = ""
	FOR M.CurSearchFieldNum = 1 TO THIS.fwUtil.GetWordCount(SearchFieldList, ',')
		M.CurSearchField = THIS.fwUtil.GetWordNum(SearchFieldList, M.CurSearchFieldNum, ',')
		M.FieldFound = .F.
		FOR M.CurTextFieldNum = 1 TO THIS.TotTextFields
			IF THIS.TextFields(M.CurTextFieldNum, 1) == M.CurSearchField
				M.FieldFound = .T.
				EXIT
			ENDIF
		NEXT
		M.CurTextFieldNum = IIF(M.FieldFound, M.CurTextFieldNum, 0)
		IF M.CurTextFieldNum = 0
			THIS.Errors.Add('Search field "' + M.CurSearchField + '" not found', , ERROR_LEVEL_ERROR, THIS.Name + '.Search')
		ELSE
			* Add field reference to comma delimited list
			* First add the actual binary ID to the aBinaryIDs array
			* We are doing this, because binary characters cannot be part of the SQL statement
			M.TotBinaryIDs = M.TotBinaryIDs + 1
			DIMENSION aBinaryIDs(M.TotBinaryIDs)
			M.aBinaryIDs(M.TotBinaryIDs) = BINTOC(M.CurTextFieldNum, 1)
			* Now add a reference to the id in the aBinaryIDs array into the comma-delimited list of IDs
			* This list will be used to construct the SQL statement later
			M.SearchFieldIdList = THIS.fwUtil.AddDelimString('M.aBinaryIDs(' + TRANSFORM(M.TotBinaryIDs) + ')', M.SearchFieldIdList, ',')
		ENDIF
	NEXT
	RETURN M.SearchFieldIdList
ENDFUNC


**********************************************************
FUNCTION CreateIndex
**********************************************************
	LPARAMETERS KeyField, TextFieldList, Proximity, NoiseWords, KeywordIDSize
	LOCAL aFieldInfo, CurField, TotFields, KeyFieldType, KeyFieldWidth, KeyFieldDecimals, ;
		KeyFieldDefn, TableDefn, aTextFields, TotTextFields
	PRIVATE inClassFullText, InMethodCreateIndex
	M.inClassFullText = .T.
	M.InMethodCreateIndex = .T.
	* Clear messages
	THIS.Errors.Clear()
	IF PCOUNT() < 5
		ERROR 'Invalid number of arguments'
		RETURN .F.
	ENDIF
	IF TYPE('M.KeyField') <> 'C' OR EMPTY(M.KeyField)
		THIS.Errors.Add('Empty or invalid KeyField argument', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	IF TYPE('M.TextFieldList') <> 'C' OR EMPTY(M.TextFieldList)
		THIS.Errors.Add('Empty or invalid TextFieldList argument', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	IF TYPE('M.Proximity') <> 'L'
		THIS.Errors.Add('Invalid Proximity argument', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	IF TYPE('M.NoiseWords') <> 'C'
		THIS.Errors.Add('Invalid NoiseWords argument', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	IF TYPE('M.KeywordIDSize') <> 'N' OR NOT INLIST(M.KeywordIDSize, 2, 4)
		THIS.Errors.Add('Invalid KeywordIDSize argument', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	IF EMPTY(THIS.Table)
		THIS.Errors.Add('You must populate the Table property before calling this function', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	* Open the table being indexed
	IF ! THIS.OpenSourceTable()
		RETURN .F.
	ENDIF
	M.KeyField = ALLTRIM(M.KeyField)
	M.TextFieldList = ALLTRIM(M.TextFieldList)
	* Determine datatype and length of index field
	DIMENSION aFieldInfo(1)
	TotFields = AFIELDS(aFieldInfo)
	M.FoundField = .F.
	FOR M.CurField = 1 TO M.TotFields
		IF LOWER(M.aFieldInfo(M.CurField, 1)) == LOWER(M.KeyField)
			M.FoundField = .T.
			EXIT
		ENDIF
	NEXT
	IF ! M.FoundField
		* Did not find the field
		THIS.Errors.Add('Could not find field "' + M.KeyField + '" in table "' + THIS.Table + '"', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	IF ! aFieldInfo(M.CurField, 2) $ 'CBDINFT'
		THIS.Errors.Add('Key field is of an invalid data type (' + aFieldInfo(M.CurField, 2) + ')', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
		RETURN .F.
	ENDIF
	* nPrecision and nDecimals are ignored where not applicable
	M.KeyFieldDefn = aFieldInfo(M.CurField, 2) + ' (' + TRANSFORM(aFieldInfo(M.CurField, 3)) + ',' + TRANSFORM(aFieldInfo(M.CurField, 4)) + ') NULL'
	* Split the TextFieldString in an array
	M.TotTextFields = THIS.fwUtil.GetWordCount(M.TextFieldList, ',')
	DIMENSION M.aTextFields(M.TotTextFields, 2)
	FOR M.CurTextFieldNum = 1 TO M.TotTextFields
		M.CurTextField = LOWER(THIS.fwUtil.GetWordNum(M.TextFieldList, M.CurTextFieldNum, ','))
		IF ':' $ M.CurTextField
			* A content type was specified
			M.aTextFields(M.CurTextFieldNum, 1) = ALLTRIM(LEFT(M.CurTextField, AT(':', M.CurTextField) - 1))
			M.aTextFields(M.CurTextFieldNum, 2) = ALLTRIM(SUBSTR(M.CurTextField, AT(':', M.CurTextField) + 1))
		ELSE
			M.aTextFields(M.CurTextFieldNum, 1) = ALLTRIM(M.CurTextField)
			M.aTextFields(M.CurTextFieldNum, 2) = ''
		ENDIF
	NEXT
	* Check that all TextFields exist and that they are of type C or M
	FOR M.CurTextFieldNum = 1 TO M.TotTextFields
		M.FoundField = .F.
		FOR M.CurField = 1 TO M.TotFields
			IF LOWER(M.aFieldInfo(M.CurField, 1)) == LOWER(M.aTextFields(M.CurTextFieldNum, 1))
				M.FoundField = .T.
				EXIT
			ENDIF
		NEXT
		IF ! M.FoundField
			* Did not find the field
			THIS.Errors.Add('Could not find field "' + M.aTextFields(M.CurTextFieldNum, 1) + '" in table "' + LOWER(THIS.TableName) + '"', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
			RETURN .F.
		ELSE
			IF NOT INLIST(M.aFieldInfo(M.CurField, 2), 'C', 'M', 'V')
				THIS.Errors.Add('Field "' + M.aTextFields(M.CurTextFieldNum, 1) + '" is not of type CHARACTER or MEMO', , ERROR_LEVEL_ERROR, THIS.Name + '.CreateIndex')
				RETURN .F.
			ENDIF
		ENDIF
	NEXT
	* Delete full-text-index table and index if it already exists
	IF NOT THIS.DeleteIndexInternal()
		RETURN .F.
	ENDIF
	* Construct index table creation command and create table
	M.TableDefn = [CREATE TABLE (THIS.TablePath + FORCEEXT(THIS.TableName, 'ftt')) FREE (KeyWordID C(] + LTRIM(STR(M.KeywordIDSize)) + [) NULL, IndexValue ] + M.KeyFieldDefn + IIF(THIS.fwUtil.GetWordCount(M.TextFieldList, ',') > 1, [, TextField C(1)], '') + IIF(M.Proximity, [, WordPos C(2))], [, WordCount C(2))])
	&TableDefn
	USE DBF() ALIAS FTIndex EXCLUSIVE
	SELECT 0
	* Construct word index table creation command and create table
	M.TableDefn = [CREATE TABLE (THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwt')) FREE (KeyWord C(] + TRANSFORM(MAX_KEYWORD_LEN) + [), KeywordID C(] + LTRIM(STR(M.KeywordIDSize)) + [))]
	&TableDefn
	USE DBF() ALIAS FTWordIndex EXCLUSIVE
	INDEX ON Keyword TAG Keyword OF THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwx')
	* Store index and text field names in table and populate object properties
	THIS.StoreIndexAttributes(M.KeyField, @M.aTextFields, M.Proximity, M.NoiseWords, M.KeywordIDSize)
	* Update index table
	RETURN THIS.UpdateIndex()
ENDFUNC


**********************************************************
FUNCTION UpdateIndex
**********************************************************
	LOCAL TotRecs, StartSeconds, IndexValue, TotUniqueWords, ;
		CurSourceRecord, RegExp, inClassFullText, KeyField, ;
		TotNoiseWordOccurrences, TotWordOccurrences, ;
		TotShortWordOccurrences, TotLongWordOccurrences
	M.inClassFullText = .T.
	SET TALK OFF
	M.StartSeconds = SECONDS()
	IF PCOUNT() <> 0
		ERROR "UpdateIndex does not take any arguments"
		RETURN .F.
	ENDIF
	IF TYPE('M.InMethodCreateIndex') = 'U'
		* Method was called directly by an external program
		PRIVATE InMethodCreateIndex
		M.InMethodCreateIndex = .F.
		* Clear messages
		THIS.Errors.Clear()
		IF EMPTY(THIS.Table)
			THIS.Errors.Add('You must populate the Table property before calling this function', , ERROR_LEVEL_ERROR, THIS.Name + '.UpdateIndex')
			RETURN .F.
		ENDIF
		* Verify that the table has a full-text index
		IF ! THIS.IsIndexed()
			THIS.Errors.Add('Non-existent or invalid Full Text Index for table "' + THIS.Table + '"', , ERROR_LEVEL_ERROR, THIS.Name + '.UpdateIndex')
			RETURN .F.
		ENDIF
		* Open the source and index tables
		IF ! THIS.OpenSourceTable() OR ! THIS.OpenIndexTables(.T.)
			RETURN .F.
		ENDIF
		IF THIS.TotTextFields <= 0
			* Corrupt FT index
			RETURN .F.
		ENDIF
		SELECT FTWordIndex
		SELECT FTIndex
		* Deleting the index makes mass updates much faster
		DELETE TAG ALL OF THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx')
		SET SAFETY OFF
		ZAP
	ENDIF
	IF ! THIS.ExclusiveIndexing
		* NOTE: overall speed is slower because the index has to be updated
		* whenever a record is inserted, but the tables are available for
		* searching during indexing process
		THIS.IndexFTIndexTable()
		* Close and re-open index tables in shared mode
		USE IN FTIndex
		USE IN FTWordIndex
		IF ! THIS.OpenIndexTables()
			RETURN .F.
		ENDIF
	ENDIF
	M.KeyField = THIS.KeyField
	* Mark FT Index as locked
	THIS.SetLockFlag(.T.)
	* Create Regular Expression object so that the COM control is cached (big performance improvement)
	RegExp = CREATEOBJECT("VBScript.RegExp")
	THIS.fwUtil.PushSetting('EXACT', 'ON')
	SELECT (THIS.TableName)
	COUNT TO M.TotRecs
	M.TotUniqueWords = 0
	M.TotWordOccurrences = 0
	M.TotNoiseWordOccurrences = 0
	M.TotShortWordOccurrences = 0
	M.TotLongWordOccurrences = 0
	M.CurSourceRecord = 1
	* Scan through table being indexed
	SCAN
		IF MOD(M.CurSourceRecord, 10) = 0
			* This runs every 10 records indexed
			THIS.StatusMessage(M.CurSourceRecord, M.TotRecs, M.StartSeconds)
		ENDIF
		M.IndexValue = &KeyField
		IF NOT ISNULL(M.IndexValue)
			IF NOT THIS.IndexRecord(M.IndexValue, @M.TotUniqueWords, @M.TotWordOccurrences, @M.TotNoiseWordOccurrences, @M.TotShortWordOccurrences, @M.TotLongWordOccurrences, .F.)
				THIS.Errors.Add('Full-text indexing was not completed', , ERROR_LEVEL_ERROR, THIS.Name + IIF(M.InMethodCreateIndex, '.CreateIndex', '.UpdateIndex'))
				RETURN .F.
			ENDIF
		ENDIF
		M.CurSourceRecord = M.CurSourceRecord + 1
	ENDSCAN
	IF THIS.ExclusiveIndexing
		THIS.IndexFTIndexTable()
	ENDIF
	THIS.StatusMessage(M.TotRecs, M.TotRecs, M.StartSeconds)
	* Mark FT Index as free
	THIS.SetLockFlag(.F.)
	THIS.fwUtil.PopSetting('EXACT')
	M.ResultMessage = 'Finished indexing ' + TRANSFORM(M.TotRecs) + ' Records in ' + LTRIM(STR(SECONDS() - M.StartSeconds, 10, 2)) + ' seconds:' + CRLF + ;
			'      - Indexed ' + TRANSFORM(M.TotWordOccurrences) + ' occurrences of ' + TRANSFORM(M.TotUniqueWords) + ' unique words'
	IF M.TotLongWordOccurrences > 0
		M.ResultMessage = ResultMessage + CRLF + ;
			'      - Truncated ' + TRANSFORM(M.TotLongWordOccurrences) + ' occurrences of words longer than ' + TRANSFORM(MAX_KEYWORD_LEN) + ' characters'
	ENDIF
	IF OCCURS(CRLF, THIS.NoiseWords) - 1 > 0
		M.ResultMessage = ResultMessage + CRLF + ;
			'      - Skipped ' + TRANSFORM(M.TotNoiseWordOccurrences) + ' occurrences of ' + TRANSFORM(OCCURS(CRLF, THIS.NoiseWords) - 1) + ' noise words'
	ENDIF
	IF M.TotShortWordOccurrences > 0
		M.ResultMessage = ResultMessage + CRLF + ;
			'      - Skipped ' + TRANSFORM(M.TotShortWordOccurrences) + ' occurrences of single-character words'
	ENDIF
	THIS.Errors.Add(M.ResultMessage, , ERROR_LEVEL_INFO, THIS.Name + '.UpdateIndex')
	THIS.CloseIndexTables()
	RETURN .T.
ENDFUNC


**********************************************************
FUNCTION UpdateRecordIndex
**********************************************************
	LPARAMETERS IndexValue, DeleteRecord
	LOCAL KeyField, ErrNum
	PRIVATE inClassFullText

	M.inClassFullText = .T.
	* Clear messages
	THIS.Errors.Clear()
	IF EMPTY(THIS.Table)
		THIS.Errors.Add('You must populate the Table property before calling this function', , ERROR_LEVEL_ERROR, THIS.Name + '.UpdateRecordIndex')
		RETURN .F.
	ENDIF
	* This class cannot index tables with NULL IndexValues
	IF ISNULL(M.IndexValue)
		THIS.Errors.Add('The index value may not be NULL', , ERROR_LEVEL_ERROR, THIS.Name + '.UpdateRecordIndex')
		RETURN .F.
	ENDIF
	* Verify that the table has a full-text index
	IF ! THIS.IsIndexed()
		THIS.Errors.Add('Non-existent or invalid Full Text Index for table "' + THIS.Table + '"', , ERROR_LEVEL_ERROR, THIS.Name + '.UpdateRecordIndex')
		RETURN .F.
	ENDIF
	* Open the source and index tables
	IF ! THIS.OpenSourceTable() OR ! THIS.OpenIndexTables()
		RETURN .F.
	ENDIF
	SELECT FTWordIndex
	M.KeyField = THIS.KeyField
	IF THIS.TotTextFields <= 0
		* Corrupt FT index
		RETURN .F.
	ENDIF
	SELECT FTIndex
	SET ORDER TO IndexValue
	IF SEEK(M.IndexValue)
		M.ErrNum = 0
		THIS.fwUtil.PushSetting('ERROR', 'ErrNum = ERROR()')
		REPLACE IndexValue WITH .NULL., KeywordID WITH .NULL. FOR IndexValue = M.IndexValue
		THIS.fwUtil.PopSetting('ERROR')
		IF M.ErrNum = 1581
			* This index was created with a previous version of the class and doesn't support NULL values
			DELETE FOR IndexValue = M.IndexValue
		ENDIF
	ENDIF
	IF ! M.DeleteRecord
		THIS.fwUtil.PushSetting('EXACT', 'ON')
		SELECT (THIS.TableName)
		LOCATE FOR &KeyField = IndexValue
		IF FOUND()
			IF NOT THIS.IndexRecord(M.IndexValue, 0, 0, 0, 0, 0, .T.)
				RETURN .F.
			ENDIF
		ENDIF
		THIS.fwUtil.PopSetting('EXACT')
	ENDIF
	THIS.CloseIndexTables()
	RETURN .T.
ENDFUNC


**********************************************************
PROTECTED PROCEDURE IndexRecord
* Indexes the current record in the selected area
	LPARAMETERS IndexValue, TotUniqueWords, TotWordOccurrences, TotNoiseWordOccurrences, ;
				TotShortWordOccurrences, TotLongWordOccurrences, IndexingSingleRecord
	LOCAL TextFieldValue, WordPos, CurKeyword, CurTextField, CurTextFieldNum,;
		TotWords, FoundWordMatch
	FOR M.CurTextFieldNum = 1 TO THIS.TotTextFields
		M.CurTextField = THIS.TextFields(M.CurTextFieldNum, 1)
		M.CurTextFieldType = THIS.TextFields(M.CurTextFieldNum, 2)
		M.TextFieldValue = &CurTextField
		IF M.CurTextFieldType = 'html'
			* Strip HTML code
			M.TextFieldValue = THIS.StripHTML(M.TextFieldValue)
		ENDIF
		* Split the field into words
		M.TextFieldValue = THIS.WordSplit(M.TextFieldValue)
		* Loop through all keywords in current record and store them in full text index
		M.TotWords = MEMLINES(M.TextFieldValue)
		_MLINE = 0
		FOR M.WordPos = 1 TO M.TotWords
			M.CurKeyword = LOWER(MLINE(M.TextFieldValue, 1, _MLINE))
			M.CurKeywordLen = LEN(M.CurKeyword)
			DO CASE
			CASE M.CurKeywordLen <= 1
				M.TotShortWordOccurrences = M.TotShortWordOccurrences + 1
			CASE CRLF + M.CurKeyword+ CRLF $ THIS.NoiseWords
				M.TotNoiseWordOccurrences = M.TotNoiseWordOccurrences + 1
			OTHERWISE
				IF M.CurKeywordLen > MAX_KEYWORD_LEN
					M.TotLongWordOccurrences = M.TotLongWordOccurrences + 1
					M.CurKeyword = LEFT(M.CurKeyword, MAX_KEYWORD_LEN)
				ENDIF
				M.TotWordOccurrences = M.TotWordOccurrences + 1
				IF ! SEEK(M.CurKeyword, 'FTWordIndex')
					M.TotUniqueWords = M.TotUniqueWords + 1
					IF RECCOUNT('FTWordIndex') >= IIF(THIS.KeywordIDSize = 2, 65538, 4294967296)
						THIS.Errors.Add('Maximum number of keywords exceeded in "' + THIS.Table + '"', , ERROR_LEVEL_ERROR, THIS.Name + '.IndexRecord')
						RETURN .F.
					ENDIF
					APPEND BLANK IN 'FTWordIndex'
					* Use RECNO() - (32769|2147483648), so that we utilize the full range that can be represented with BINTOC:
					* KeywordIDSize: 2 => (–32,768 to 32,767) 4 => (–2,147,483,648 to 2,147,483,647)
					REPLACE FTWordIndex.Keyword WITH M.CurKeyword, FTWordIndex.KeywordID WITH BINTOC(RECNO('FTWordIndex') - IIF(THIS.KeywordIDSize = 2, 32769, 2147483648), THIS.KeywordIDSize)
				ENDIF
				SELECT FTIndex
				IF THIS.TotTextFields > 1
					IF ! THIS.Proximity
						IF THIS.ExclusiveIndexing AND ! M.IndexingSingleRecord
							M.FoundWordMatch = .F.
							GO BOTTOM
							DO WHILE ! BOF() AND FTIndex.IndexValue == M.IndexValue
								IF FTIndex.KeywordID = FTWordIndex.KeywordID AND FTIndex.TextField = BINTOC(M.CurTextFieldNum, 1)
									M.FoundWordMatch = .T.
									EXIT DO
								ENDIF
								SKIP -1
							ENDDO
						ELSE
							LOCATE FOR IndexValue = M.IndexValue AND KeywordID = FTWordIndex.KeywordID AND TextField = BINTOC(M.CurTextFieldNum, 1)
							M.FoundWordMatch = FOUND()
						ENDIF
						IF M.FoundWordMatch
							* Increment word count
							REPLACE FTIndex.WordCount WITH BINTOC(CTOBIN(FTIndex.WordCount) + 1, 2)
						ELSE
							IF M.IndexingSingleRecord AND SEEK(.NULL.)
								* Re-use previously-deleted record
								REPLACE KeywordID WITH FTWordIndex.KeywordID, IndexValue WITH M.IndexValue, TextField WITH BINTOC(M.CurTextFieldNum, 1), WordCount WITH BINTOC(1, 2)
							ELSE
								INSERT INTO FTIndex (KeywordID, IndexValue, TextField, WordCount) VALUES (FTWordIndex.KeywordID, M.IndexValue, BINTOC(M.CurTextFieldNum, 1), BINTOC(1, 2))
							ENDIF
						ENDIF
					ELSE
						IF M.IndexingSingleRecord AND SEEK(.NULL.)
							* Re-use previously-deleted record
							REPLACE KeywordID WITH FTWordIndex.KeywordID, IndexValue WITH M.IndexValue, TextField WITH BINTOC(M.CurTextFieldNum, 1), WordPos WITH BINTOC(M.WordPos, 2)
						ELSE
							INSERT INTO FTIndex (KeywordID, IndexValue, TextField, WordPos) VALUES (FTWordIndex.KeywordID, M.IndexValue, BINTOC(M.CurTextFieldNum, 1), BINTOC(M.WordPos, 2))
						ENDIF
					ENDIF
				ELSE
					IF ! THIS.Proximity
						IF THIS.ExclusiveIndexing AND ! M.IndexingSingleRecord
							M.FoundWordMatch = .F.
							GO BOTTOM
							DO WHILE ! BOF() AND FTIndex.IndexValue == M.IndexValue
								IF FTIndex.KeywordID = FTWordIndex.KeywordID
									M.FoundWordMatch = .T.
									EXIT DO
								ENDIF
								SKIP -1
							ENDDO
						ELSE
							LOCATE FOR IndexValue = M.IndexValue AND KeywordID = FTWordIndex.KeywordID
							M.FoundWordMatch = FOUND()
						ENDIF
						IF M.FoundWordMatch
							* Increment word count
							REPLACE FTIndex.WordCount WITH BINTOC(CTOBIN(FTIndex.WordCount) + 1, 2)
						ELSE
							IF M.IndexingSingleRecord AND SEEK(.NULL.)
								* Re-use previously-deleted record
								REPLACE KeywordID WITH FTWordIndex.KeywordID, IndexValue WITH M.IndexValue, WordCount WITH BINTOC(1, 2)
							ELSE
								INSERT INTO FTIndex (KeywordID, IndexValue, WordCount) VALUES (FTWordIndex.KeywordID, M.IndexValue, BINTOC(1, 2))
							ENDIF
						ENDIF
					ELSE
						IF M.IndexingSingleRecord AND SEEK(.NULL.)
							* Re-use previously-deleted record
							REPLACE KeywordID WITH FTWordIndex.KeywordID, IndexValue WITH M.IndexValue, WordPos WITH BINTOC(M.WordPos, 2)
						ELSE
							INSERT INTO FTIndex (KeywordID, IndexValue, WordPos) VALUES (FTWordIndex.KeywordID, M.IndexValue, BINTOC(M.WordPos, 2))
						ENDIF
					ENDIF
				ENDIF
				SELECT (THIS.TableName)
			ENDCASE
		NEXT
	NEXT
	RETURN .T.
ENDPROC


***********************************************************
PROTECTED FUNCTION StripHTML
* Strips all HTML code from a string
	LPARAMETERS TextBuffer
	* Strip "Sent From" sections
	M.TextBuffer = THIS.fwUtil.RegExpReplace(M.TextBuffer, '<DIV\s*CLASS="?msgFrom"?>.*?<!--EndMsgFrom-->\s*</DIV>', ' ')
	* Strip HTML tags
	M.TextBuffer = THIS.fwUtil.RegExpReplace(M.TextBuffer, '<.*?>', ' ')
	* Strip HTML codes (e.g. &nbsp;)
	M.TextBuffer = THIS.fwUtil.RegExpReplace(M.TextBuffer, '&\w+?;', ' ')
	RETURN M.TextBuffer
ENDFUNC


***********************************************************
PROTECTED FUNCTION WordSplit
* Splits text fields into words
	LPARAMETERS TextBuffer
	* FoxWeb-Forum-Specific: Strip certain phrases... ("FoxWeb Support Team", "support@foxweb.com", "Re:")
	M.TextBuffer = THIS.fwUtil.RegExpReplace(M.TextBuffer, '\bfoxweb\s*support\s*team\b|\bsupport\s*@\s*foxweb\s*\.\s*com\b|\bre:|\bSent by .*? on \d\d\/\d\d\/\d\d\d\d \d\d:\d\d:\d\d (?:AM|PM)\b', ' ')
	* Strip non-word characters
	M.TextBuffer = THIS.fwUtil.RegExpReplace(M.TextBuffer, '\W+', CHR(13))	&& same as '[^A-Za-z0-9_]'
	RETURN M.TextBuffer
ENDFUNC


**********************************************************
PROTECTED FUNCTION IsNoiseWord
* Returns .T. if word is a noise word, or is outside permitted size
* Note that the word may contain wildcard characters (* and ?)
	LPARAMETERS Keyword
	IF '*' $ M.Keyword OR '?' $ M.Keyword
		* This method takes much longer, so we only use it if necessary
		RETURN THIS.fwUtil.RegExpTest(THIS.NoiseWords, '\b' + STRTRAN(STRTRAN(M.Keyword, '*', '\w*'), '?', '\w?') + '\b') OR ! BETWEEN(LEN(CHRTRAN(M.Keyword, '*?', '')), 2, MAX_KEYWORD_LEN)
	ELSE
		RETURN CRLF + M.Keyword + CRLF $ THIS.NoiseWords OR ! BETWEEN(LEN(M.Keyword), 2, MAX_KEYWORD_LEN)
	ENDIF
ENDFUNC


**********************************************************
PROTECTED PROCEDURE IndexFTIndexTable
* Index full-text-index table
	SELECT FTIndex
	INDEX ON KeywordID TAG KeywordID OF THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx')
	INDEX ON IndexValue TAG IndexValue OF THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx')
	IF THIS.TotTextFields > 1
		INDEX ON TextField TAG TextField OF THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx')
	ENDIF
ENDPROC


**********************************************************
PROTECTED PROCEDURE StoreIndexAttributes
* Store full text index info in index table
	LPARAMETERS KeyField, aTextFields, Proximity, NoiseWords, KeywordIDSize
	LOCAL CurTextFieldNum, CurTextField, TotTextFields, CurNoiseWord, ;
			CurNoiseWordNum, NoiseWordsCount, GoodNoiseWordsCount, TextFieldDescr, ;
			CurTextFieldDescrChunk
	* Store index field name
	THIS.KeyField = M.KeyField
	INSERT INTO FTWordIndex (KeyWord) VALUES ('-1' + THIS.KeyField)
	* Store text field names
	THIS.TotTextFields = ALEN(M.aTextFields, 1)
	DIMENSION THIS.TextFields(THIS.TotTextFields, 2)
	ACOPY(M.aTextFields, THIS.TextFields)
	FOR M.CurTextFieldNum = 1 TO THIS.TotTextFields
		M.TextFieldDescr = THIS.TextFields(M.CurTextFieldNum, 1) + IIF(EMPTY(THIS.TextFields(M.CurTextFieldNum, 2)), '', ':' + THIS.TextFields(M.CurTextFieldNum, 2))
		* We have a loop to split the field description into multiple records if the length is too long
		FOR M.CurTextFieldDescrChunk = 1 TO CEILING(LEN(M.TextFieldDescr) / (MAX_KEYWORD_LEN - 6))
			INSERT INTO FTWordIndex (KeyWord) VALUES ('-2' + PADL(TRANSFORM(M.CurTextFieldNum), 2) + '+' + TRANSFORM(M.CurTextFieldDescrChunk) + SUBSTR(M.TextFieldDescr, ((M.CurTextFieldDescrChunk - 1) * (MAX_KEYWORD_LEN - 6)) + 1, MAX_KEYWORD_LEN - 6))
		NEXT
	NEXT
	* Store proximity flag
	THIS.Proximity = M.Proximity
	INSERT INTO FTWordIndex (KeyWord) VALUES ('-4' + IIF(THIS.Proximity, '1', '0'))
	* Store KeywordIDSize flag
	THIS.KeywordIDSize = M.KeywordIDSize
	INSERT INTO FTWordIndex (KeyWord) VALUES ('-5' + LTRIM(STR(THIS.KeywordIDSize)))
	* Noise Words
	M.NoiseWords = LOWER(STRTRAN(ALLTRIM(THIS.fwUtil.RegExpReplace(M.NoiseWords, '\W+', ' ')), ' ', CRLF))	&& 'W' is the same as '[^A-Za-z0-9_]'
	M.NoiseWordsCount = MEMLINES(M.NoiseWords)
	M.GoodNoiseWordsCount = 0
	THIS.NoiseWords = CRLF
	_MLINE = 0
	FOR M.CurNoiseWordNum = 1 TO M.NoiseWordsCount
		M.CurNoiseWord = MLINE(M.NoiseWords, 1, _MLINE)
		IF LEN(M.CurNoiseWord) <= MAX_KEYWORD_LEN AND ! SEEK(M.CurNoiseWord, 'FTWordIndex')
			THIS.NoiseWords = THIS.NoiseWords + M.CurNoiseWord + CRLF
			M.GoodNoiseWordsCount = M.GoodNoiseWordsCount + 1
			APPEND BLANK IN 'FTWordIndex'
			* Use RECNO() - (32769|2147483648), so that we utilize the full range that can be represented with BINTOC:
			* KeywordIDSize: 2 => (–32,768 to 32,767) 4 => (–2,147,483,648 to 2,147,483,647)
			REPLACE FTWordIndex.Keyword WITH M.CurNoiseWord, FTWordIndex.KeywordID WITH BINTOC(RECNO('FTWordIndex') - IIF(THIS.KeywordIDSize = 2, 32769, 2147483648), THIS.KeywordIDSize)
		ENDIF
	NEXT
	* Store the number of noise words
	INSERT INTO FTWordIndex (KeyWord) VALUES ('-6' + TRANSFORM(M.GoodNoiseWordsCount))
ENDPROC


**********************************************************
PROTECTED PROCEDURE OpenSourceTable
* Open the table being indexed
	RETURN THIS.OpenTable(THIS.Table)
ENDPROC


**********************************************************
PROTECTED FUNCTION OpenIndexTables
* Open the full text index tables and index files
	LPARAMETERS OpenExclusive
	LOCAL HadErrors, inClassFullText, CurTextField, CurTextFieldNum, TotTextFields, ;
			NoiseWordsCount, GoodNoiseWordsCount, CurRecord, TextFieldIndex
	M.inClassFullText = .T.
	M.HadErrors = .F.
	* Open the two index tables
	M.HadErrors = M.HadErrors OR ! THIS.OpenTable(;
			THIS.TablePath + FORCEEXT(THIS.TableName, 'ftt'), ;
			'FTIndex', ;
			M.OpenExclusive, ;
			THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx'))
	M.HadErrors = M.HadErrors OR ! THIS.OpenTable(;
			THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwt'), ;
			'FTWordIndex', ;
			M.OpenExclusive, ;
			THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwx'), ;
			'keyword')
	IF ! M.HadErrors
		THIS.fwUtil.PushSetting('EXACT', 'OFF')	&& SET EXACT OFF and remember current setting
		* Look for KeyField name
		IF SEEK('-1')
			THIS.KeyField = TRIM(SUBSTR(FTWordIndex.Keyword, 3))
		ELSE
			M.HadErrors = .T.
			THIS.Errors.Add('Corrupt Full-Text Index for table "' + THIS.Table + '" (no KeyField attribute)', , ERROR_LEVEL_ERROR)
		ENDIF
		IF SEEK('-2')
			IF SUBSTR(FTWordIndex.Keyword, 5, 1) = '+'
				* This is a new style table, which uses char pos 6 to store an index that ensures
				* correct order in situations where field descriptions have to be split into
				* multiple records
				M.FieldDescrOffset = 2
			ELSE
				M.FieldDescrOffset = 0
			ENDIF
			* First loop through the field descriptions to count the number of text fields
			M.TextFieldIndex = ''
			M.TotTextFields = 0
			SEEK '-2'
			SCAN WHILE FTWordIndex.Keyword = '-2'
				IF M.TextFieldIndex != SUBSTR(FTWordIndex.Keyword, 3, 2)
					* This record holds a new text field description
					M.TotTextFields = M.TotTextFields + 1
					M.TextFieldIndex = SUBSTR(FTWordIndex.Keyword, 3, 2)
				ENDIF
			ENDSCAN
			* Now loop through the field descriptions again to store them in an array
			DIMENSION THIS.TextFields(M.TotTextFields, 2)
			M.CurTextField = ''
			M.TextFieldIndex = ''
			M.CurTextFieldNum = 0
			SEEK '-2'
			SCAN WHILE FTWordIndex.Keyword = '-2'
				IF M.TextFieldIndex != SUBSTR(FTWordIndex.Keyword, 3, 2)
					* This record holds a new text field description
					IF M.CurTextFieldNum > 0
						* Store previous field description
						IF ':' $ M.CurTextField
							* A content type was specified
							THIS.TextFields(M.CurTextFieldNum, 1) = LEFT(M.CurTextField, AT(':', M.CurTextField) - 1)
							THIS.TextFields(M.CurTextFieldNum, 2) = TRIM(SUBSTR(M.CurTextField, AT(':', M.CurTextField) + 1))
						ELSE
							THIS.TextFields(M.CurTextFieldNum, 1) = TRIM(M.CurTextField)
							THIS.TextFields(M.CurTextFieldNum, 2) = ''
						ENDIF
					ENDIF
					M.CurTextFieldNum = M.CurTextFieldNum + 1
					M.CurTextField = SUBSTR(FTWordIndex.Keyword, 5 + M.FieldDescrOffset)
					M.TextFieldIndex = SUBSTR(FTWordIndex.Keyword, 3, 2)
				ELSE
					* This record holds a continuation of the text field description
					M.CurTextField = M.CurTextField + SUBSTR(FTWordIndex.Keyword, 5 + M.FieldDescrOffset)
				ENDIF
			ENDSCAN
			* Store final field description
			IF ':' $ M.CurTextField
				* A content type was specified
				THIS.TextFields(M.CurTextFieldNum, 1) = LEFT(M.CurTextField, AT(':', M.CurTextField) - 1)
				THIS.TextFields(M.CurTextFieldNum, 2) = TRIM(SUBSTR(M.CurTextField, AT(':', M.CurTextField) + 1))
			ELSE
				THIS.TextFields(M.CurTextFieldNum, 1) = TRIM(M.CurTextField)
				THIS.TextFields(M.CurTextFieldNum, 2) = ''
			ENDIF
			THIS.TotTextFields = M.TotTextFields
		ELSE
			THIS.TotTextFields = 0
			M.HadErrors = .T.
			THIS.Errors.Add('Corrupt Full Text Index for table "' + THIS.Table + '" (no TextField attributes)', , ERROR_LEVEL_ERROR)
		ENDIF
		* Look for IndexingInProcess flag
		IF SEEK('-3') AND KeywordID == '1'
			* FT index tables are locked by another call to the CreateIndex, or UpdateIndex method
			THIS.Errors.Add('Full-text index is currently being updated', , ERROR_LEVEL_WARNING)
		ENDIF
		* Look for Proximity flag
		IF SEEK('-4')
			THIS.Proximity = IIF(TRIM(SUBSTR(FTWordIndex.Keyword, 3)) == '1', .T., .F.)
		ELSE
			M.HadErrors = .T.
			THIS.Errors.Add('Corrupt Full-Text Index for table "' + THIS.Table + '" (no ProximityFlag attribute)', , ERROR_LEVEL_ERROR)
		ENDIF
		* Look for KeywordIDSize flag
		IF SEEK('-5')
			THIS.KeywordIDSize = VAL(SUBSTR(FTWordIndex.Keyword, 3))
		ELSE
			M.HadErrors = .T.
			THIS.Errors.Add('Corrupt Full-Text Index for table "' + THIS.Table + '" (no KeywordIDSize attribute)', , ERROR_LEVEL_ERROR)
		ENDIF
		* Look for NoiseWordsCount
		IF SEEK('-6')
			M.NoiseWordsCount = VAL(SUBSTR(FTWordIndex.Keyword, 3))
		ELSE
			M.NoiseWordsCount = 0
			M.HadErrors = .T.
			THIS.Errors.Add('Corrupt Full-Text Index for table "' + THIS.Table + '" (no NoiseWordsCount attribute)', , ERROR_LEVEL_ERROR)
		ENDIF
		THIS.NoiseWords = CRLF
		IF M.NoiseWordsCount > 0
			* Retrieve noise words
			M.GoodNoiseWordsCount = 0
			SET ORDER TO
			GO TOP
			FOR M.CurRecord = 1 TO M.NoiseWordsCount + 20	&& 20 is arbitrary
				DO CASE
				CASE LEFT(FTWordIndex.Keyword, 2) == '-6'
					EXIT FOR
				CASE LEFT(FTWordIndex.Keyword, 1) == '-'
					* This is an index attribute.  Skip record
				OTHERWISE
					THIS.NoiseWords = THIS.NoiseWords + TRIM(FTWordIndex.Keyword) + CRLF
					M.GoodNoiseWordsCount = M.GoodNoiseWordsCount + 1
				ENDCASE
				SKIP
			NEXT
			SET ORDER TO TAG keyword
			IF M.NoiseWordsCount <> M.GoodNoiseWordsCount
				THIS.Errors.Add('Corrupt Full-Text Index for table "' + THIS.Table + '" (noise words do not match NoiseWordsCount attribute)', , ERROR_LEVEL_ERROR)
			ENDIF
		ENDIF
		THIS.fwUtil.PopSetting('EXACT')
	ENDIF
	IF M.HadErrors
		THIS.CloseIndexTables
		RETURN .F.
	ELSE
		RETURN .T.
	ENDIF
ENDFUNC


**********************************************************
PROTECTED PROCEDURE CloseIndexTables
	IF USED('FTIndex')
		USE IN FTIndex
	ENDIF
	IF USED('FTWordIndex')
		USE IN FTWordIndex
	ENDIF
ENDPROC


***********************************************************
PROTECTED PROCEDURE OpenTable
* Opens a table if it is not already open
* This is the same as the OpenTable method in fwUtil, but it
* is duplicated here to allow better error handling
	LPARAMETERS FilePath, Alias, Exclusive, IndexFile, DefaultOrder
	LOCAL RunTimeError
	IF TYPE("M.alias") <> "C"
		M.alias = JustFName(JustStem(M.FilePath))
		IF '!' $ M.alias
			* Strip database delimiter
			M.alias = SUBSTR(M.alias, RAT('!', M.alias) + 1)
		ENDIF
	ENDIF
	THIS.fwUtil.PushSetting('ERROR', 'RunTimeError = ERROR()')
	RunTimeError = -999
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
	IF TYPE('M.IndexFile') = 'C' AND M.RunTimeError = -999
		SET INDEX TO (M.IndexFile) ADDITIVE
	ENDIF
	IF TYPE('M.DefaultOrder') = 'C' AND M.RunTimeError = -999
		SET ORDER TO (M.DefaultOrder)
	ENDIF
	THIS.fwUtil.PopSetting('ERROR')
	IF M.RunTimeError <> -999
		* There was an error opening the table
		THIS.Errors.Add(MESSAGE(), M.RunTimeError, ERROR_LEVEL_ERROR, THIS.Name + '.OpenTable(' + M.FilePath + ', ' + IIF(M.Exclusive, '.T.', '.F.') + ')')
		RETURN .F.
	ELSE
		RETURN .T.
	ENDIF
ENDPROC


**********************************************************
FUNCTION IsIndexed
* Check if a full text index exists
	IF EMPTY(THIS.Table)
		THIS.Errors.Add('You must populate the Table property before calling this function', , ERROR_LEVEL_ERROR, THIS.Name + '.UpdateIndex')
		RETURN .F.
	ENDIF
	RETURN FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftt')) AND ;
			FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx')) AND ;
			FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwt')) AND ;
			FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwx'))
ENDFUNC


**********************************************************
FUNCTION DeleteIndex
* Delete full-text-index table and index if it already exists
	THIS.Errors.Clear()
	RETURN THIS.DeleteIndexInternal()
ENDFUNC


**********************************************************
PROTECTED FUNCTION DeleteIndexInternal
* Delete full-text-index table and index if it already exists
	LOCAL RunTimeError
	* Trap errors while deleting tables
	RunTimeError = -999
	THIS.fwUtil.PushSetting('ERROR', 'RunTimeError = ERROR()')
	IF FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftt'))
		DELETE FILE THIS.TablePath + FORCEEXT(THIS.TableName, 'ftt')
	ENDIF
	IF FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx'))
		DELETE FILE THIS.TablePath + FORCEEXT(THIS.TableName, 'ftx')
	ENDIF
	IF FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwt'))
		DELETE FILE THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwt')
	ENDIF
	IF FILE(THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwx'))
		DELETE FILE THIS.TablePath + FORCEEXT(THIS.TableName, 'ftwx')
	ENDIF
	* Restore default error handler
	THIS.fwUtil.PopSetting('ERROR')
	IF M.RunTimeError <> -999
		THIS.Errors.Add(MESSAGE(), M.RunTimeError, ERROR_LEVEL_ERROR, THIS.Name + '.DeleteIndex')
		RETURN .F.
	ELSE
		RETURN .T.
	ENDIF
ENDFUNC


**********************************************************
PROTECTED PROCEDURE SetLockFlag
* Set the lock flag in FTWordINdex table
	LPARAMETERS LockFlag
	SELECT FTWordIndex
	SEEK '-3'
	IF FOUND()
		REPLACE KeywordID WITH IIF(M.LockFlag, '1', '0')
	ELSE
		INSERT INTO FTWordIndex (Keyword, KeywordID) VALUES ('-3', IIF(M.LockFlag, '1', '0'))
	ENDIF
ENDPROC


**********************************************************
FUNCTION GetIndexAttributes
**********************************************************
	LPARAMETERS KeyField, aTextFields, Proximity, NoiseWords, KeywordIDSize
	THIS.Errors.Clear()
	IF ! THIS.OpenIndexTables()
		RETURN .F.
	ENDIF
	M.KeyField = THIS.KeyField
	M.Proximity = THIS.Proximity
	M.NoiseWords = STRTRAN(ALLTRIM(CHRTRAN(THIS.NoiseWords, CRLF, ' ')), ' ', ', ')
	M.KeywordIDSize = THIS.KeywordIDSize
	DIMENSION aTextFields(1)
	ACOPY(THIS.TextFields, M.aTextFields)
	THIS.CloseIndexTables()
	RETURN .T.
ENDPROC


**********************************************************
PROTECTED PROCEDURE StatusMessage
* Displays status message -- this method can be sub-classed to handle special messaging interfaces
	LPARAMETERS CurRecord, TotRecords, StartSeconds
	IF M.CurRecord = M.TotRecords
		ResultMessage = ''
	ELSE
		M.ResultMessage = 'Indexing record ' + TRANSFORM(M.CurRecord) + ' of ' + TRANSFORM(M.TotRecords)
	ENDIF
	IF THIS.InFoxWeb
		* Increase Script timeout if necessary
		IF (SECONDS() - M.StartSeconds) + 5 >= Server.ScriptTimeout
			Server.AddScriptTimeout(5)
		ENDIF
		* Write current progress to browser
		Response.Write(M.ResultMessage + '<br>')
		Response.Flush
	ELSE
		_VFP.StatusBar = M.ResultMessage
	ENDIF
ENDPROC



**********************************************************
PROTECTED PROCEDURE ErrorHandler
* THIS IS NOT USED
*Error handler
	LPARAMETERS nError, cMethod, nLine
	THIS.Errors.Add(MESSAGE(), M.nError, ERROR_LEVEL_ERROR, THIS.Name + M.cMethod + ' line:' +  TRANSFORM(M.nLine))
	DODEFAULT()
ENDPROC


**********************************************************
PROTECTED PROCEDURE Table_ASSIGN
**********************************************************
	LPARAMETERS cTable
	IF TYPE('M.cTable') <> 'C' OR EMPTY(M.cTable)
		ERROR 'Invalid or empty ' + THIS.Name + '.Table property'
		RETURN .F.
	ENDIF
	IF EMPTY(JUSTEXT(M.cTable))
		M.cTable = FORCEEXT(M.cTable, 'dbf')
	ENDIF
	IF ! FILE(M.cTable)
		ERROR 'Table specified in ' + THIS.Name + '.Table property does not exist'
		RETURN .F.
	ENDIF
	THIS.Table = M.cTable
	THIS.TablePath = ADDBS(JUSTPATH(FULLPATH(M.cTable)))
	THIS.TableName = JUSTFNAME(M.cTable)
ENDPROC


**********************************************************
PROTECTED PROCEDURE ExclusiveIndexing_ASSIGN
**********************************************************
	LPARAMETERS lExclusiveIndexing
	IF TYPE('M.lExclusiveIndexing') <> 'L'
		ERROR 'Invalid ' + THIS.Name + '.ExclusiveIndexing property'
		RETURN .F.
	ENDIF
	THIS.ExclusiveIndexing = M.lExclusiveIndexing
ENDPROC


**********************************************************
FUNCTION Version
**********************************************************
	LPARAMETERS VersionType
	DO CASE
	CASE PCOUNT() = 0
		M.VersionType = 0
	CASE TYPE('M.VersionType') <> 'N' OR NOT BETWEEN(M.VersionType, 0, 1)
		ERROR 'Invalid VersionType argument'
		RETURN ''
	ENDCASE
	DO CASE
	CASE M.VersionType = 0
		RETURN COMPONENT_VERSION
	CASE M.VersionType = 1
		RETURN VFP_COMPILE_VERSION
	OTHERWISE
		* Should never happen
		RETURN -1
	ENDCASE
ENDFUNC


ENDDEFINE
**********************************************************
*       END OF DEFINITION OF classFullText CLASS
**********************************************************


**************************************************************
* Definition of the classErrors class
**************************************************************
DEFINE CLASS classErrors AS Custom
	PROTECTED DefaultSource
	Count = 0
	DefaultSource = ''
	DIMENSION Item(5)


PROTECTED PROCEDURE Init
	LPARAMETERS ErrorSource
	IF PCOUNT() < 1 OR TYPE('M.ErrorSource') <> 'C'
		M.ErrorSource = ''
	ENDIF
	THIS.DefaultSource = M.ErrorSource
ENDPROC


PROCEDURE Clear
	PRIVATE inClassErrors
	M.inClassErrors = .T.
	THIS.Count = 0
	THIS.Item = .NULL.
ENDPROC


PROCEDURE Add
	LPARAMETERS ErrorDescription, ErrorNumber, ErrorSeverity, ErrorSource
	LOCAL TotRows
	PRIVATE inClassErrors
	M.inClassErrors = .T.
	* Populate properties with default values if they were not passed
	IF PCOUNT() < 1 OR TYPE('M.ErrorDescription') <> 'C'
		M.ErrorDescription = ''
	ENDIF
	IF PCOUNT() < 2 OR TYPE('M.ErrorNumber') <> 'N'
		M.ErrorNumber = 0
	ENDIF
	IF PCOUNT() < 3 OR TYPE('M.ErrorSeverity') <> 'N'
		M.ErrorSeverity = 0
	ENDIF
	IF PCOUNT() < 4 OR TYPE('M.ErrorSource') <> 'C'
		M.ErrorSource = THIS.DefaultSource
	ENDIF
	THIS.Count = THIS.Count + 1
	* Add rows to array
	DIMENSION THIS.Item(THIS.Count)
	THIS.Item(THIS.Count) = CREATEOBJECT('classError')
	* Populate properties of error object
	THIS.Item(THIS.Count).Number = M.ErrorNumber
	THIS.Item(THIS.Count).Description = M.ErrorDescription
	THIS.Item(THIS.Count).Severity = M.ErrorSeverity
	THIS.Item(THIS.Count).Source = M.ErrorSource
ENDPROC


**********************************************************
PROTECTED PROCEDURE Count_ASSIGN
**********************************************************
	LPARAMETERS nCount
	IF TYPE('M.inClassErrors') <> 'L' OR ! M.inClassErrors
		ERROR 'Count property is read-only'
		RETURN .F.
	ENDIF
	THIS.Count = M.nCount
ENDPROC

ENDDEFINE
**********************************************************
*       END OF DEFINITION OF classErrors CLASS
**********************************************************


**************************************************************
* Definition of the classError class
**************************************************************
DEFINE CLASS classError AS Custom
	Number = 0
	Description = ''
	Severity = 0
	Source = ''

ENDDEFINE
**********************************************************
*       END OF DEFINITION OF classError CLASS
**********************************************************
