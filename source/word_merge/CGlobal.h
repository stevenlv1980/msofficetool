// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CGlobal 包装类

class CGlobal : public COleDispatchDriver
{
public:
	CGlobal(){} // 调用 COleDispatchDriver 默认构造函数
	CGlobal(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CGlobal(const CGlobal& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// _Global 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x3e9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x3ea, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Documents()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Windows()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveDocument()
	{
		LPDISPATCH result;
		InvokeHelper(0x3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ActiveWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x4, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Selection()
	{
		LPDISPATCH result;
		InvokeHelper(0x5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_WordBasic()
	{
		LPDISPATCH result;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_PrintPreview()
	{
		BOOL result;
		InvokeHelper(0x1b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_PrintPreview(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x1b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_RecentFiles()
	{
		LPDISPATCH result;
		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_NormalTemplate()
	{
		LPDISPATCH result;
		InvokeHelper(0x8, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_System()
	{
		LPDISPATCH result;
		InvokeHelper(0x9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AutoCorrect()
	{
		LPDISPATCH result;
		InvokeHelper(0xa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FontNames()
	{
		LPDISPATCH result;
		InvokeHelper(0xb, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_LandscapeFontNames()
	{
		LPDISPATCH result;
		InvokeHelper(0xc, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_PortraitFontNames()
	{
		LPDISPATCH result;
		InvokeHelper(0xd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Languages()
	{
		LPDISPATCH result;
		InvokeHelper(0xe, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Assistant()
	{
		LPDISPATCH result;
		InvokeHelper(0xf, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_FileConverters()
	{
		LPDISPATCH result;
		InvokeHelper(0x11, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Dialogs()
	{
		LPDISPATCH result;
		InvokeHelper(0x13, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CaptionLabels()
	{
		LPDISPATCH result;
		InvokeHelper(0x14, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AutoCaptions()
	{
		LPDISPATCH result;
		InvokeHelper(0x15, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AddIns()
	{
		LPDISPATCH result;
		InvokeHelper(0x16, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Tasks()
	{
		LPDISPATCH result;
		InvokeHelper(0x1c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_MacroContainer()
	{
		LPDISPATCH result;
		InvokeHelper(0x37, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CommandBars()
	{
		LPDISPATCH result;
		InvokeHelper(0x39, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SynonymInfo(LPCTSTR Word, VARIANT * LanguageID)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, Word, LanguageID);
		return result;
	}
	LPDISPATCH get_VBE()
	{
		LPDISPATCH result;
		InvokeHelper(0x3d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ListGalleries()
	{
		LPDISPATCH result;
		InvokeHelper(0x41, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_ActivePrinter()
	{
		CString result;
		InvokeHelper(0x42, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_ActivePrinter(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x42, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Templates()
	{
		LPDISPATCH result;
		InvokeHelper(0x43, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CustomizationContext()
	{
		LPDISPATCH result;
		InvokeHelper(0x44, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_CustomizationContext(LPDISPATCH newValue)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x44, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_KeyBindings()
	{
		LPDISPATCH result;
		InvokeHelper(0x45, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_KeysBoundTo(long KeyCategory, LPCTSTR Command, VARIANT * CommandParameter)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_BSTR VTS_PVARIANT ;
		InvokeHelper(0x46, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, KeyCategory, Command, CommandParameter);
		return result;
	}
	LPDISPATCH get_FindKey(long KeyCode, VARIANT * KeyCode2)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x47, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, parms, KeyCode, KeyCode2);
		return result;
	}
	LPDISPATCH get_Options()
	{
		LPDISPATCH result;
		InvokeHelper(0x5d, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CustomDictionaries()
	{
		LPDISPATCH result;
		InvokeHelper(0x5f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void put_StatusBar(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x61, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_ShowVisualBasicEditor()
	{
		BOOL result;
		InvokeHelper(0x68, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_ShowVisualBasicEditor(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x68, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_IsObjectValid(LPDISPATCH Object)
	{
		BOOL result;
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x6d, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, parms, Object);
		return result;
	}
	LPDISPATCH get_HangulHanjaDictionaries()
	{
		LPDISPATCH result;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL Repeat(VARIANT * Times)
	{
		BOOL result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x131, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Times);
		return result;
	}
	void DDEExecute(long Channel, LPCTSTR Command)
	{
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x136, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, Command);
	}
	long DDEInitiate(LPCTSTR App, LPCTSTR Topic)
	{
		long result;
		static BYTE parms[] = VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x137, DISPATCH_METHOD, VT_I4, (void*)&result, parms, App, Topic);
		return result;
	}
	void DDEPoke(long Channel, LPCTSTR Item, LPCTSTR Data)
	{
		static BYTE parms[] = VTS_I4 VTS_BSTR VTS_BSTR ;
		InvokeHelper(0x138, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel, Item, Data);
	}
	CString DDERequest(long Channel, LPCTSTR Item)
	{
		CString result;
		static BYTE parms[] = VTS_I4 VTS_BSTR ;
		InvokeHelper(0x139, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, Channel, Item);
		return result;
	}
	void DDETerminate(long Channel)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x13a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Channel);
	}
	void DDETerminateAll()
	{
		InvokeHelper(0x13b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long BuildKeyCode(long Arg1, VARIANT * Arg2, VARIANT * Arg3, VARIANT * Arg4)
	{
		long result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x13c, DISPATCH_METHOD, VT_I4, (void*)&result, parms, Arg1, Arg2, Arg3, Arg4);
		return result;
	}
	CString KeyString(long KeyCode, VARIANT * KeyCode2)
	{
		CString result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x13d, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, KeyCode, KeyCode2);
		return result;
	}
	BOOL CheckSpelling(LPCTSTR Word, VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * MainDictionary, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
	{
		BOOL result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x144, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, Word, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
		return result;
	}
	LPDISPATCH GetSpellingSuggestions(LPCTSTR Word, VARIANT * CustomDictionary, VARIANT * IgnoreUppercase, VARIANT * MainDictionary, VARIANT * SuggestionMode, VARIANT * CustomDictionary2, VARIANT * CustomDictionary3, VARIANT * CustomDictionary4, VARIANT * CustomDictionary5, VARIANT * CustomDictionary6, VARIANT * CustomDictionary7, VARIANT * CustomDictionary8, VARIANT * CustomDictionary9, VARIANT * CustomDictionary10)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0x147, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Word, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10);
		return result;
	}
	void Help(VARIANT * HelpType)
	{
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x149, DISPATCH_METHOD, VT_EMPTY, NULL, parms, HelpType);
	}
	LPDISPATCH NewWindow()
	{
		LPDISPATCH result;
		InvokeHelper(0x159, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString CleanString(LPCTSTR String)
	{
		CString result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x162, DISPATCH_METHOD, VT_BSTR, (void*)&result, parms, String);
		return result;
	}
	void ChangeFileOpenDirectory(LPCTSTR Path)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x163, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Path);
	}
	float InchesToPoints(float Inches)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x172, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Inches);
		return result;
	}
	float CentimetersToPoints(float Centimeters)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x173, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Centimeters);
		return result;
	}
	float MillimetersToPoints(float Millimeters)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x174, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Millimeters);
		return result;
	}
	float PicasToPoints(float Picas)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x175, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Picas);
		return result;
	}
	float LinesToPoints(float Lines)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x176, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Lines);
		return result;
	}
	float PointsToInches(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17c, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToCentimeters(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17d, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToMillimeters(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17e, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToPicas(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x17f, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToLines(float Points)
	{
		float result;
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x180, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points);
		return result;
	}
	float PointsToPixels(float Points, VARIANT * fVertical)
	{
		float result;
		static BYTE parms[] = VTS_R4 VTS_PVARIANT ;
		InvokeHelper(0x181, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Points, fVertical);
		return result;
	}
	float PixelsToPoints(float Pixels, VARIANT * fVertical)
	{
		float result;
		static BYTE parms[] = VTS_R4 VTS_PVARIANT ;
		InvokeHelper(0x182, DISPATCH_METHOD, VT_R4, (void*)&result, parms, Pixels, fVertical);
		return result;
	}
	LPDISPATCH get_LanguageSettings()
	{
		LPDISPATCH result;
		InvokeHelper(0x6f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AnswerWizard()
	{
		LPDISPATCH result;
		InvokeHelper(0x70, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_AutoCorrectEmail()
	{
		LPDISPATCH result;
		InvokeHelper(0x71, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// _Global 属性
public:

};
