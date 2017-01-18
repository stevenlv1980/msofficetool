// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDialogSheet 包装类

class CDialogSheet : public COleDispatchDriver
{
public:
	CDialogSheet(){} // 调用 COleDispatchDriver 默认构造函数
	CDialogSheet(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDialogSheet(const CDialogSheet& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// DialogSheet 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Activate()
	{
		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Copy(VARIANT Before, VARIANT After)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x227, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Before, &After);
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	CString get_CodeName()
	{
		CString result;
		InvokeHelper(0x55d, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	CString get__CodeName()
	{
		CString result;
		InvokeHelper(0x80010000, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put__CodeName(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x80010000, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Index()
	{
		long result;
		InvokeHelper(0x1e6, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void Move(VARIANT Before, VARIANT After)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x27d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Before, &After);
	}
	CString get_Name()
	{
		CString result;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_Name(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Next()
	{
		LPDISPATCH result;
		InvokeHelper(0x1f6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	CString get_OnDoubleClick()
	{
		CString result;
		InvokeHelper(0x274, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_OnDoubleClick(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x274, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_OnSheetActivate()
	{
		CString result;
		InvokeHelper(0x407, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_OnSheetActivate(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x407, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_OnSheetDeactivate()
	{
		CString result;
		InvokeHelper(0x439, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_OnSheetDeactivate(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x439, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_PageSetup()
	{
		LPDISPATCH result;
		InvokeHelper(0x3e6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Previous()
	{
		LPDISPATCH result;
		InvokeHelper(0x1f7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void __PrintOut(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x389, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate);
	}
	void PrintPreview(VARIANT EnableChanges)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x119, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &EnableChanges);
	}
	void _Protect(VARIANT Password, VARIANT DrawingObjects, VARIANT Contents, VARIANT Scenarios, VARIANT UserInterfaceOnly)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x11a, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Password, &DrawingObjects, &Contents, &Scenarios, &UserInterfaceOnly);
	}
	BOOL get_ProtectContents()
	{
		BOOL result;
		InvokeHelper(0x124, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_ProtectDrawingObjects()
	{
		BOOL result;
		InvokeHelper(0x125, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_ProtectionMode()
	{
		BOOL result;
		InvokeHelper(0x487, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	BOOL get_ProtectScenarios()
	{
		BOOL result;
		InvokeHelper(0x126, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void _SaveAs(LPCTSTR Filename, VARIANT FileFormat, VARIANT Password, VARIANT WriteResPassword, VARIANT ReadOnlyRecommended, VARIANT CreateBackup, VARIANT AddToMru, VARIANT TextCodepage, VARIANT TextVisualLayout)
	{
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x11c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &AddToMru, &TextCodepage, &TextVisualLayout);
	}
	void Select(VARIANT Replace)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Replace);
	}
	void Unprotect(VARIANT Password)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x11d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Password);
	}
	long get_Visible()
	{
		long result;
		InvokeHelper(0x22e, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Visible(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x22e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Shapes()
	{
		LPDISPATCH result;
		InvokeHelper(0x561, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void _Dummy29()
	{
		InvokeHelper(0x1001d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Arcs(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x2f8, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy31()
	{
		InvokeHelper(0x1001f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy32()
	{
		InvokeHelper(0x10020, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Buttons(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x22d, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy34()
	{
		InvokeHelper(0x10022, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL get_EnableCalculation()
	{
		BOOL result;
		InvokeHelper(0x590, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_EnableCalculation(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x590, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	void _Dummy36()
	{
		InvokeHelper(0x10024, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH ChartObjects(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x424, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH CheckBoxes(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x338, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void CheckSpelling(VARIANT CustomDictionary, VARIANT IgnoreUppercase, VARIANT AlwaysSuggest, VARIANT SpellLang)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x1f9, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &SpellLang);
	}
	void _Dummy40()
	{
		InvokeHelper(0x10028, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy41()
	{
		InvokeHelper(0x10029, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy42()
	{
		InvokeHelper(0x1002a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy43()
	{
		InvokeHelper(0x1002b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy44()
	{
		InvokeHelper(0x1002c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy45()
	{
		InvokeHelper(0x1002d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	BOOL get_DisplayAutomaticPageBreaks()
	{
		BOOL result;
		InvokeHelper(0x283, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayAutomaticPageBreaks(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x283, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH Drawings(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x304, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH DrawingObjects(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x58, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH DropDowns(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x344, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	BOOL get_EnableAutoFilter()
	{
		BOOL result;
		InvokeHelper(0x484, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_EnableAutoFilter(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x484, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_EnableSelection()
	{
		long result;
		InvokeHelper(0x591, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_EnableSelection(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x591, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_EnableOutlining()
	{
		BOOL result;
		InvokeHelper(0x485, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_EnableOutlining(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x485, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	BOOL get_EnablePivotTable()
	{
		BOOL result;
		InvokeHelper(0x486, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_EnablePivotTable(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x486, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	VARIANT Evaluate(VARIANT Name)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x1, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Name);
		return result;
	}
	VARIANT _Evaluate(VARIANT Name)
	{
		VARIANT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xfffffffb, DISPATCH_METHOD, VT_VARIANT, (void*)&result, parms, &Name);
		return result;
	}
	void _Dummy56()
	{
		InvokeHelper(0x10038, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void ResetAllPageBreaks()
	{
		InvokeHelper(0x592, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH GroupBoxes(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x342, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH GroupObjects(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x459, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH Labels(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x349, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH Lines(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x2ff, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH ListBoxes(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x340, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH get_Names()
	{
		LPDISPATCH result;
		InvokeHelper(0x1ba, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH OLEObjects(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x31f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy65()
	{
		InvokeHelper(0x10041, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy66()
	{
		InvokeHelper(0x10042, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy67()
	{
		InvokeHelper(0x10043, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH OptionButtons(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x33a, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy69()
	{
		InvokeHelper(0x10045, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Ovals(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x321, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void Paste(VARIANT Destination, VARIANT Link)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0xd3, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Destination, &Link);
	}
	void _PasteSpecial(VARIANT Format, VARIANT Link, VARIANT DisplayAsIcon, VARIANT IconFileName, VARIANT IconIndex, VARIANT IconLabel)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x403, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Format, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel);
	}
	LPDISPATCH Pictures(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x303, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy74()
	{
		InvokeHelper(0x1004a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy75()
	{
		InvokeHelper(0x1004b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy76()
	{
		InvokeHelper(0x1004c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Rectangles(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x306, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy78()
	{
		InvokeHelper(0x1004e, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy79()
	{
		InvokeHelper(0x1004f, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	CString get_ScrollArea()
	{
		CString result;
		InvokeHelper(0x599, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_ScrollArea(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x599, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH ScrollBars(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x33e, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy82()
	{
		InvokeHelper(0x10052, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy83()
	{
		InvokeHelper(0x10053, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Spinners(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x346, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy85()
	{
		InvokeHelper(0x10055, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy86()
	{
		InvokeHelper(0x10056, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH TextBoxes(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x309, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	void _Dummy88()
	{
		InvokeHelper(0x10058, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy89()
	{
		InvokeHelper(0x10059, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy90()
	{
		InvokeHelper(0x1005a, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH get_HPageBreaks()
	{
		LPDISPATCH result;
		InvokeHelper(0x58a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_VPageBreaks()
	{
		LPDISPATCH result;
		InvokeHelper(0x58b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_QueryTables()
	{
		LPDISPATCH result;
		InvokeHelper(0x59a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayPageBreaks()
	{
		BOOL result;
		InvokeHelper(0x59b, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayPageBreaks(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x59b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Comments()
	{
		LPDISPATCH result;
		InvokeHelper(0x23f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Hyperlinks()
	{
		LPDISPATCH result;
		InvokeHelper(0x571, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ClearCircles()
	{
		InvokeHelper(0x59c, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void CircleInvalid()
	{
		InvokeHelper(0x59d, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long get__DisplayRightToLeft()
	{
		long result;
		InvokeHelper(0x288, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put__DisplayRightToLeft(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x288, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_AutoFilter()
	{
		LPDISPATCH result;
		InvokeHelper(0x319, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	BOOL get_DisplayRightToLeft()
	{
		BOOL result;
		InvokeHelper(0x6ee, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_DisplayRightToLeft(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x6ee, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Scripts()
	{
		LPDISPATCH result;
		InvokeHelper(0x718, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void _PrintOut(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x6ec, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName);
	}
	void _CheckSpelling(VARIANT CustomDictionary, VARIANT IgnoreUppercase, VARIANT AlwaysSuggest, VARIANT SpellLang, VARIANT IgnoreFinalYaa, VARIANT SpellScript)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x719, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &SpellLang, &IgnoreFinalYaa, &SpellScript);
	}
	LPDISPATCH get_Tab()
	{
		LPDISPATCH result;
		InvokeHelper(0x411, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_MailEnvelope()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void SaveAs(LPCTSTR Filename, VARIANT FileFormat, VARIANT Password, VARIANT WriteResPassword, VARIANT ReadOnlyRecommended, VARIANT CreateBackup, VARIANT AddToMru, VARIANT TextCodepage, VARIANT TextVisualLayout, VARIANT Local)
	{
		static BYTE parms[] = VTS_BSTR VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x785, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Filename, &FileFormat, &Password, &WriteResPassword, &ReadOnlyRecommended, &CreateBackup, &AddToMru, &TextCodepage, &TextVisualLayout, &Local);
	}
	LPDISPATCH get_CustomProperties()
	{
		LPDISPATCH result;
		InvokeHelper(0x7ee, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_SmartTags()
	{
		LPDISPATCH result;
		InvokeHelper(0x7e0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Protection()
	{
		LPDISPATCH result;
		InvokeHelper(0xb0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void PasteSpecial(VARIANT Format, VARIANT Link, VARIANT DisplayAsIcon, VARIANT IconFileName, VARIANT IconIndex, VARIANT IconLabel, VARIANT NoHTMLFormatting)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x788, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Format, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel, &NoHTMLFormatting);
	}
	void Protect(VARIANT Password, VARIANT DrawingObjects, VARIANT Contents, VARIANT Scenarios, VARIANT UserInterfaceOnly, VARIANT AllowFormattingCells, VARIANT AllowFormattingColumns, VARIANT AllowFormattingRows, VARIANT AllowInsertingColumns, VARIANT AllowInsertingRows, VARIANT AllowInsertingHyperlinks, VARIANT AllowDeletingColumns, VARIANT AllowDeletingRows, VARIANT AllowSorting, VARIANT AllowFiltering, VARIANT AllowUsingPivotTables)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x7ed, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Password, &DrawingObjects, &Contents, &Scenarios, &UserInterfaceOnly, &AllowFormattingCells, &AllowFormattingColumns, &AllowFormattingRows, &AllowInsertingColumns, &AllowInsertingRows, &AllowInsertingHyperlinks, &AllowDeletingColumns, &AllowDeletingRows, &AllowSorting, &AllowFiltering, &AllowUsingPivotTables);
	}
	void _Dummy113()
	{
		InvokeHelper(0x10071, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy114()
	{
		InvokeHelper(0x10072, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void _Dummy115()
	{
		InvokeHelper(0x10073, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void PrintOut(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName)
	{
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x939, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName);
	}
	BOOL get_EnableFormatConditionsCalculation()
	{
		BOOL result;
		InvokeHelper(0x9cf, DISPATCH_PROPERTYGET, VT_BOOL, (void*)&result, NULL);
		return result;
	}
	void put_EnableFormatConditionsCalculation(BOOL newValue)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x9cf, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Sort()
	{
		LPDISPATCH result;
		InvokeHelper(0x370, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ExportAsFixedFormat(long Type, VARIANT Filename, VARIANT Quality, VARIANT IncludeDocProperties, VARIANT IgnorePrintAreas, VARIANT From, VARIANT To, VARIANT OpenAfterPublish, VARIANT FixedFormatExtClassPtr)
	{
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x9bd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Type, &Filename, &Quality, &IncludeDocProperties, &IgnorePrintAreas, &From, &To, &OpenAfterPublish, &FixedFormatExtClassPtr);
	}
	VARIANT get_DefaultButton()
	{
		VARIANT result;
		InvokeHelper(0x359, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_DefaultButton(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x359, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	LPDISPATCH get_DialogFrame()
	{
		LPDISPATCH result;
		InvokeHelper(0x347, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH EditBoxes(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x33c, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	VARIANT get_Focus()
	{
		VARIANT result;
		InvokeHelper(0x32e, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
	}
	void put_Focus(VARIANT newValue)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x32e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, &newValue);
	}
	BOOL Hide(VARIANT Cancel)
	{
		BOOL result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x32d, DISPATCH_METHOD, VT_BOOL, (void*)&result, parms, &Cancel);
		return result;
	}
	BOOL Show()
	{
		BOOL result;
		InvokeHelper(0x1f0, DISPATCH_METHOD, VT_BOOL, (void*)&result, NULL);
		return result;
	}

	// DialogSheet 属性
public:

};
