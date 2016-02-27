		****************************************************
		* ************************************************ *
		* *   FTWrapper.PRG                              * *
		* *   (C) Aegis Group 2004                       * *
		* ************************************************ *
		****************************************************

* This is a wrapper that sub-classes the classFullText class
* and illustrates how to add progress status handling via the
* StatusMessage method.

**************************************************************
* Definition of the classMyFullText class
**************************************************************

DEFINE CLASS classMyFullText AS classFullText
	ContainerForm = .NULL.


**********************************************************
PROCEDURE Init
**********************************************************
	LPARAMETERS ContainerForm
	IF PCOUNT() >= 1 AND TYPE('M.ContainerForm') = 'O'
		* A reference to the container form, which will be used to display status messages
		THIS.ContainerForm = M.ContainerForm
	ENDIF
	* Call default method code in classFullText
	DODEFAULT()	&& *** THIS IS REQUIRED ***
ENDPROC


**********************************************************
PROTECTED PROCEDURE StatusMessage
* Sub-class of the StatusMessage method in fwFullText.
* Displays status messages in an Edit Box in the calling form.
	LPARAMETERS CurRecord, TotRecords, StartSeconds

	IF M.CurRecord = M.TotRecords
		* We are done!
		ResultMessage = ''
	ELSE
		* Still in progress
		M.ResultMessage = 'Indexing record ' + TRANSFORM(M.CurRecord) + ' of ' + TRANSFORM(M.TotRecords)
	ENDIF
	IF NOT ISNULL(THIS.ContainerForm)
		* Display the current ResultMessage in the form reference that was passed to the Init method
		THIS.ContainerForm.txtResult.Value = M.ResultMessage
	ELSE
		* Display the current ResultMessage in the VFP status bar
		_VFP.StatusBar = M.ResultMessage
	ENDIF
ENDPROC

ENDDEFINE
