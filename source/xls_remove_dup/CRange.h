// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRange 包装类

class CRange : public COleDispatchDriver
{
public:
	CRange(){} // 调用 COleDispatchDriver 默认构造函数
	CRange(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRange(const CRange& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IRange 方法
public:
	STDMETHOD(get_Application)(Application * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x94, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Creator)(XlCreator * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x95, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Parent)(LPDISPATCH * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x96, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Activate)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x130, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_AddIndent)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x427, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_AddIndent)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x427, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_Address)(VARIANT RowAbsolute, VARIANT ColumnAbsolute, XlReferenceStyle ReferenceStyle, VARIANT External, VARIANT RelativeTo, long lcid, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_PBSTR ;
		InvokeHelper(0xec, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &RowAbsolute, &ColumnAbsolute, ReferenceStyle, &External, &RelativeTo, lcid, RHS);
		return result;
	}
	STDMETHOD(get_AddressLocal)(VARIANT RowAbsolute, VARIANT ColumnAbsolute, XlReferenceStyle ReferenceStyle, VARIANT External, VARIANT RelativeTo, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x1b5, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &RowAbsolute, &ColumnAbsolute, ReferenceStyle, &External, &RelativeTo, RHS);
		return result;
	}
	STDMETHOD(AdvancedFilter)(XlFilterAction Action, VARIANT CriteriaRange, VARIANT CopyToRange, VARIANT Unique, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x36c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Action, &CriteriaRange, &CopyToRange, &Unique, RHS);
		return result;
	}
	STDMETHOD(ApplyNames)(VARIANT Names, VARIANT IgnoreRelativeAbsolute, VARIANT UseRowColumnNames, VARIANT OmitColumn, VARIANT OmitRow, XlApplyNamesOrder Order, VARIANT AppendLast, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1b9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Names, &IgnoreRelativeAbsolute, &UseRowColumnNames, &OmitColumn, &OmitRow, Order, &AppendLast, RHS);
		return result;
	}
	STDMETHOD(ApplyOutlineStyles)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x1c0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Areas)(Areas * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x238, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(AutoComplete)(LPCTSTR String, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_PBSTR ;
		InvokeHelper(0x4a1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, String, RHS);
		return result;
	}
	STDMETHOD(AutoFill)(Range * Destination, XlAutoFillType Type, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_DISPATCH VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x1c1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Destination, Type, RHS);
		return result;
	}
	STDMETHOD(AutoFilter)(VARIANT Field, VARIANT Criteria1, XlAutoFilterOperator Operator, VARIANT Criteria2, VARIANT VisibleDropDown, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x319, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Field, &Criteria1, Operator, &Criteria2, &VisibleDropDown, RHS);
		return result;
	}
	STDMETHOD(AutoFit)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xed, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(AutoFormat)(XlRangeAutoFormat Format, VARIANT Number, VARIANT Font, VARIANT Alignment, VARIANT Border, VARIANT Pattern, VARIANT Width, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x72, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Format, &Number, &Font, &Alignment, &Border, &Pattern, &Width, RHS);
		return result;
	}
	STDMETHOD(AutoOutline)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x40c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(BorderAround)(VARIANT LineStyle, XlBorderWeight Weight, XlColorIndex ColorIndex, VARIANT Color, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x42b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &LineStyle, Weight, ColorIndex, &Color, RHS);
		return result;
	}
	STDMETHOD(get_Borders)(Borders * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1b3, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Calculate)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x117, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Cells)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xee, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Characters)(VARIANT Start, VARIANT Length, Characters * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x25b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Start, &Length, RHS);
		return result;
	}
	STDMETHOD(CheckSpelling)(VARIANT CustomDictionary, VARIANT IgnoreUppercase, VARIANT AlwaysSuggest, VARIANT SpellLang, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &CustomDictionary, &IgnoreUppercase, &AlwaysSuggest, &SpellLang, RHS);
		return result;
	}
	STDMETHOD(Clear)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x6f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ClearContents)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x71, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ClearFormats)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x70, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ClearNotes)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xef, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ClearOutline)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x40d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Column)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0xf0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ColumnDifferences)(VARIANT Comparison, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x1fe, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Comparison, RHS);
		return result;
	}
	STDMETHOD(get_Columns)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xf1, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_ColumnWidth)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xf2, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ColumnWidth)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xf2, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(Consolidate)(VARIANT Sources, VARIANT Function, VARIANT TopRow, VARIANT LeftColumn, VARIANT CreateLinks, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1e2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Sources, &Function, &TopRow, &LeftColumn, &CreateLinks, RHS);
		return result;
	}
	STDMETHOD(Copy)(VARIANT Destination, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x227, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Destination, RHS);
		return result;
	}
	STDMETHOD(CopyFromRecordset)(LPUNKNOWN Data, VARIANT MaxRows, VARIANT MaxColumns, long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_UNKNOWN VTS_VARIANT VTS_VARIANT VTS_PI4 ;
		InvokeHelper(0x480, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Data, &MaxRows, &MaxColumns, RHS);
		return result;
	}
	STDMETHOD(CopyPicture)(XlPictureAppearance Appearance, XlCopyPictureFormat Format, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0xd5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Appearance, Format, RHS);
		return result;
	}
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(CreateNames)(VARIANT Top, VARIANT Left, VARIANT Bottom, VARIANT Right, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1c9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Top, &Left, &Bottom, &Right, RHS);
		return result;
	}
	STDMETHOD(CreatePublisher)(VARIANT Edition, XlPictureAppearance Appearance, VARIANT ContainsPICT, VARIANT ContainsBIFF, VARIANT ContainsRTF, VARIANT ContainsVALU, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1ca, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Edition, Appearance, &ContainsPICT, &ContainsBIFF, &ContainsRTF, &ContainsVALU, RHS);
		return result;
	}
	STDMETHOD(get_CurrentArray)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1f5, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_CurrentRegion)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xf3, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Cut)(VARIANT Destination, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x235, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Destination, RHS);
		return result;
	}
	STDMETHOD(DataSeries)(VARIANT Rowcol, XlDataSeriesType Type, XlDataSeriesDate Date, VARIANT Step, VARIANT Stop, VARIANT Trend, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1d0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Rowcol, Type, Date, &Step, &Stop, &Trend, RHS);
		return result;
	}
	STDMETHOD(get__Default)(VARIANT RowIndex, VARIANT ColumnIndex, long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &RowIndex, &ColumnIndex, lcid, RHS);
		return result;
	}
	STDMETHOD(put__Default)(VARIANT RowIndex, VARIANT ColumnIndex, long lcid, VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &RowIndex, &ColumnIndex, lcid, &newValue);
		return result;
	}
	STDMETHOD(Delete)(VARIANT Shift, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x75, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Shift, RHS);
		return result;
	}
	STDMETHOD(get_Dependents)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x21f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(DialogBox)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xf5, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_DirectDependents)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x221, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_DirectPrecedents)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x222, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(EditionOptions)(XlEditionType Type, XlEditionOptionsOption Option, VARIANT Name, VARIANT Reference, XlPictureAppearance Appearance, XlPictureAppearance ChartSize, VARIANT Format, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x46b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, Option, &Name, &Reference, Appearance, ChartSize, &Format, RHS);
		return result;
	}
	STDMETHOD(get_End)(XlDirection Direction, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x1f4, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, Direction, RHS);
		return result;
	}
	STDMETHOD(get_EntireColumn)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xf6, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_EntireRow)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0xf7, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(FillDown)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xf8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(FillLeft)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xf9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(FillRight)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xfa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(FillUp)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xfb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Find)(VARIANT What, VARIANT After, VARIANT LookIn, VARIANT LookAt, VARIANT SearchOrder, XlSearchDirection SearchDirection, VARIANT MatchCase, VARIANT MatchByte, VARIANT SearchFormat, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x18e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &What, &After, &LookIn, &LookAt, &SearchOrder, SearchDirection, &MatchCase, &MatchByte, &SearchFormat, RHS);
		return result;
	}
	STDMETHOD(FindNext)(VARIANT After, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x18f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &After, RHS);
		return result;
	}
	STDMETHOD(FindPrevious)(VARIANT After, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x190, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &After, RHS);
		return result;
	}
	STDMETHOD(get_Font)(Font * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x92, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Formula)(long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x105, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_Formula)(long lcid, VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x105, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, &newValue);
		return result;
	}
	STDMETHOD(get_FormulaArray)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x24a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_FormulaArray)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x24a, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_FormulaLabel)(XlFormulaLabel * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x564, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_FormulaLabel)(XlFormulaLabel RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x564, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_FormulaHidden)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x106, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_FormulaHidden)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x106, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_FormulaLocal)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x107, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_FormulaLocal)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x107, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_FormulaR1C1)(long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x108, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_FormulaR1C1)(long lcid, VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x108, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, &newValue);
		return result;
	}
	STDMETHOD(get_FormulaR1C1Local)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x109, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_FormulaR1C1Local)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x109, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(FunctionWizard)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x23b, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(GoalSeek)(VARIANT Goal, Range * ChangingCell, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x1d8, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Goal, ChangingCell, RHS);
		return result;
	}
	STDMETHOD(Group)(VARIANT Start, VARIANT End, VARIANT By, VARIANT Periods, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x2e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Start, &End, &By, &Periods, RHS);
		return result;
	}
	STDMETHOD(get_HasArray)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x10a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_HasFormula)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x10b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Height)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x7b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Hidden)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x10c, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Hidden)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x10c, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_HorizontalAlignment)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x88, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_HorizontalAlignment)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x88, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_IndentLevel)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xc9, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_IndentLevel)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xc9, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(InsertIndent)(long InsertAmount)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x565, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, InsertAmount);
		return result;
	}
	STDMETHOD(Insert)(VARIANT Shift, VARIANT CopyOrigin, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0xfc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Shift, &CopyOrigin, RHS);
		return result;
	}
	STDMETHOD(get_Interior)(Interior * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x81, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Item)(VARIANT RowIndex, VARIANT ColumnIndex, long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &RowIndex, &ColumnIndex, lcid, RHS);
		return result;
	}
	STDMETHOD(put_Item)(VARIANT RowIndex, VARIANT ColumnIndex, long lcid, VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT ;
		InvokeHelper(0xaa, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &RowIndex, &ColumnIndex, lcid, &newValue);
		return result;
	}
	STDMETHOD(Justify)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x1ef, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Left)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x7f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_ListHeaderRows)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x4a3, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ListNames)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xfd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_LocationInTable)(XlLocationInTable * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x2b3, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Locked)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x10d, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Locked)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x10d, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(Merge)(VARIANT Across)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x234, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Across);
		return result;
	}
	STDMETHOD(UnMerge)()
	{
		HRESULT result;
		InvokeHelper(0x568, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(get_MergeArea)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x569, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_MergeCells)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xd0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_MergeCells)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xd0, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_Name)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x6e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Name)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x6e, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(NavigateArrow)(VARIANT TowardPrecedent, VARIANT ArrowNumber, VARIANT LinkNumber, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x408, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &TowardPrecedent, &ArrowNumber, &LinkNumber, RHS);
		return result;
	}
	STDMETHOD(get__NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Next)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1f6, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(NoteText)(VARIANT Text, VARIANT Start, VARIANT Length, BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PBSTR ;
		InvokeHelper(0x467, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Text, &Start, &Length, RHS);
		return result;
	}
	STDMETHOD(get_NumberFormat)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xc1, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_NumberFormat)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xc1, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_NumberFormatLocal)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x449, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_NumberFormatLocal)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x449, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_Offset)(VARIANT RowOffset, VARIANT ColumnOffset, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xfe, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &RowOffset, &ColumnOffset, RHS);
		return result;
	}
	STDMETHOD(get_Orientation)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x86, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Orientation)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x86, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_OutlineLevel)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x10f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_OutlineLevel)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x10f, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_PageBreak)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0xff, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_PageBreak)(long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xff, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(Parse)(VARIANT ParseLine, VARIANT Destination, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1dd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &ParseLine, &Destination, RHS);
		return result;
	}
	STDMETHOD(_PasteSpecial)(XlPasteType Paste, XlPasteSpecialOperation Operation, VARIANT SkipBlanks, VARIANT Transpose, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x403, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Paste, Operation, &SkipBlanks, &Transpose, RHS);
		return result;
	}
	STDMETHOD(get_PivotField)(PivotField * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x2db, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_PivotItem)(PivotItem * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x2e4, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_PivotTable)(PivotTable * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x2cc, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Precedents)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x220, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_PrefixCharacter)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x1f8, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Previous)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x1f7, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(__PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x389, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, RHS);
		return result;
	}
	STDMETHOD(PrintPreview)(VARIANT EnableChanges, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x119, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &EnableChanges, RHS);
		return result;
	}
	STDMETHOD(get_QueryTable)(QueryTable * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x56a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Range)(VARIANT Cell1, VARIANT Cell2, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Cell1, &Cell2, RHS);
		return result;
	}
	STDMETHOD(RemoveSubtotal)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x373, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Replace)(VARIANT What, VARIANT Replacement, VARIANT LookAt, VARIANT SearchOrder, VARIANT MatchCase, VARIANT MatchByte, VARIANT SearchFormat, VARIANT ReplaceFormat, BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PBOOL ;
		InvokeHelper(0xe2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &What, &Replacement, &LookAt, &SearchOrder, &MatchCase, &MatchByte, &SearchFormat, &ReplaceFormat, RHS);
		return result;
	}
	STDMETHOD(get_Resize)(VARIANT RowSize, VARIANT ColumnSize, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x100, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &RowSize, &ColumnSize, RHS);
		return result;
	}
	STDMETHOD(get_Row)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x101, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(RowDifferences)(VARIANT Comparison, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x1ff, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Comparison, RHS);
		return result;
	}
	STDMETHOD(get_RowHeight)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x110, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_RowHeight)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x110, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_Rows)(Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x102, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Run)(VARIANT Arg1, VARIANT Arg2, VARIANT Arg3, VARIANT Arg4, VARIANT Arg5, VARIANT Arg6, VARIANT Arg7, VARIANT Arg8, VARIANT Arg9, VARIANT Arg10, VARIANT Arg11, VARIANT Arg12, VARIANT Arg13, VARIANT Arg14, VARIANT Arg15, VARIANT Arg16, VARIANT Arg17, VARIANT Arg18, VARIANT Arg19, VARIANT Arg20, VARIANT Arg21, VARIANT Arg22, VARIANT Arg23, VARIANT Arg24, VARIANT Arg25, VARIANT Arg26, VARIANT Arg27, VARIANT Arg28, VARIANT Arg29, VARIANT Arg30, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x103, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Arg1, &Arg2, &Arg3, &Arg4, &Arg5, &Arg6, &Arg7, &Arg8, &Arg9, &Arg10, &Arg11, &Arg12, &Arg13, &Arg14, &Arg15, &Arg16, &Arg17, &Arg18, &Arg19, &Arg20, &Arg21, &Arg22, &Arg23, &Arg24, &Arg25, &Arg26, &Arg27, &Arg28, &Arg29, &Arg30, RHS);
		return result;
	}
	STDMETHOD(Select)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Show)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x1f0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ShowDependents)(VARIANT Remove, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x36d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Remove, RHS);
		return result;
	}
	STDMETHOD(get_ShowDetail)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x249, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ShowDetail)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x249, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(ShowErrors)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x36e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ShowPrecedents)(VARIANT Remove, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x36f, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Remove, RHS);
		return result;
	}
	STDMETHOD(get_ShrinkToFit)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xd1, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ShrinkToFit)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xd1, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(Sort)(VARIANT Key1, XlSortOrder Order1, VARIANT Key2, VARIANT Type, XlSortOrder Order2, VARIANT Key3, XlSortOrder Order3, XlYesNoGuess Header, VARIANT OrderCustom, VARIANT MatchCase, XlSortOrientation Orientation, XlSortMethod SortMethod, XlSortDataOption DataOption1, XlSortDataOption DataOption2, XlSortDataOption DataOption3, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x370, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Key1, Order1, &Key2, &Type, Order2, &Key3, Order3, Header, &OrderCustom, &MatchCase, Orientation, SortMethod, DataOption1, DataOption2, DataOption3, RHS);
		return result;
	}
	STDMETHOD(SortSpecial)(XlSortMethod SortMethod, VARIANT Key1, XlSortOrder Order1, VARIANT Type, VARIANT Key2, XlSortOrder Order2, VARIANT Key3, XlSortOrder Order3, XlYesNoGuess Header, VARIANT OrderCustom, VARIANT MatchCase, XlSortOrientation Orientation, XlSortDataOption DataOption1, XlSortDataOption DataOption2, XlSortDataOption DataOption3, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x371, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, SortMethod, &Key1, Order1, &Type, &Key2, Order2, &Key3, Order3, Header, &OrderCustom, &MatchCase, Orientation, DataOption1, DataOption2, DataOption3, RHS);
		return result;
	}
	STDMETHOD(get_SoundNote)(SoundNote * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x394, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(SpecialCells)(XlCellType Type, VARIANT Value, Range * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x19a, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, &Value, RHS);
		return result;
	}
	STDMETHOD(get_Style)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x104, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_Style)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x104, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(SubscribeTo)(LPCTSTR Edition, XlSubscribeToFormat Format, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x1e1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Edition, Format, RHS);
		return result;
	}
	STDMETHOD(Subtotal)(long GroupBy, XlConsolidationFunction Function, VARIANT TotalList, VARIANT Replace, VARIANT PageBreaks, XlSummaryRow SummaryBelowData, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x372, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, GroupBy, Function, &TotalList, &Replace, &PageBreaks, SummaryBelowData, RHS);
		return result;
	}
	STDMETHOD(get_Summary)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x111, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Table)(VARIANT RowInput, VARIANT ColumnInput, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x1f1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &RowInput, &ColumnInput, RHS);
		return result;
	}
	STDMETHOD(get_Text)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x8a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(TextToColumns)(VARIANT Destination, XlTextParsingType DataType, XlTextQualifier TextQualifier, VARIANT ConsecutiveDelimiter, VARIANT Tab, VARIANT Semicolon, VARIANT Comma, VARIANT Space, VARIANT Other, VARIANT OtherChar, VARIANT FieldInfo, VARIANT DecimalSeparator, VARIANT ThousandsSeparator, VARIANT TrailingMinusNumbers, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x410, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Destination, DataType, TextQualifier, &ConsecutiveDelimiter, &Tab, &Semicolon, &Comma, &Space, &Other, &OtherChar, &FieldInfo, &DecimalSeparator, &ThousandsSeparator, &TrailingMinusNumbers, RHS);
		return result;
	}
	STDMETHOD(get_Top)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x7e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Ungroup)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xf4, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_UseStandardHeight)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x112, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_UseStandardHeight)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x112, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_UseStandardWidth)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x113, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_UseStandardWidth)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x113, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_Validation)(Validation * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x56b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Value)(VARIANT RangeValueDataType, long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x6, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &RangeValueDataType, lcid, RHS);
		return result;
	}
	STDMETHOD(put_Value)(VARIANT RangeValueDataType, long lcid, VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x6, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &RangeValueDataType, lcid, &newValue);
		return result;
	}
	STDMETHOD(get_Value2)(long lcid, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_PVARIANT ;
		InvokeHelper(0x56c, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, lcid, RHS);
		return result;
	}
	STDMETHOD(put_Value2)(long lcid, VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x56c, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, lcid, &newValue);
		return result;
	}
	STDMETHOD(get_VerticalAlignment)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x89, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_VerticalAlignment)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x89, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(get_Width)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x7a, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Worksheet)(Worksheet * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x15c, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_WrapText)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x114, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_WrapText)(VARIANT RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x114, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, &newValue);
		return result;
	}
	STDMETHOD(AddComment)(VARIANT Text, Comment * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x56d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Text, RHS);
		return result;
	}
	STDMETHOD(get_Comment)(Comment * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x38e, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ClearComments)()
	{
		HRESULT result;
		InvokeHelper(0x56e, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(get_Phonetic)(Phonetic * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x56f, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_FormatConditions)(FormatConditions * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x570, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_ReadingOrder)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x3cf, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ReadingOrder)(long RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x3cf, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(get_Hyperlinks)(Hyperlinks * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x571, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_Phonetics)(Phonetics * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x713, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(SetPhonetic)()
	{
		HRESULT result;
		InvokeHelper(0x714, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(get_ID)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x715, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(put_ID)(LPCTSTR RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x715, DISPATCH_PROPERTYPUT, VT_HRESULT, (void*)&result, parms, newValue);
		return result;
	}
	STDMETHOD(_PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x6ec, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName, RHS);
		return result;
	}
	STDMETHOD(get_PivotCell)(PivotCell * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x7dd, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Dirty)()
	{
		HRESULT result;
		InvokeHelper(0x7de, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(get_Errors)(Errors * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x7df, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_SmartTags)(SmartTags * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x7e0, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Speak)(VARIANT SpeakDirection, VARIANT SpeakFormulas)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x7e1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &SpeakDirection, &SpeakFormulas);
		return result;
	}
	STDMETHOD(PasteSpecial)(XlPasteType Paste, XlPasteSpecialOperation Operation, VARIANT SkipBlanks, VARIANT Transpose, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x788, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Paste, Operation, &SkipBlanks, &Transpose, RHS);
		return result;
	}
	STDMETHOD(get_AllowEdit)(BOOL * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x7e4, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_ListObject)(ListObject * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x8d1, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_XPath)(XPath * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x8d2, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(get_ServerActions)(Actions * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PDISPATCH ;
		InvokeHelper(0x9bb, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(RemoveDuplicates)(VARIANT Columns, XlYesNoGuess Header)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_I4 ;
		InvokeHelper(0x9bc, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Columns, Header);
		return result;
	}
	STDMETHOD(PrintOut)(VARIANT From, VARIANT To, VARIANT Copies, VARIANT Preview, VARIANT ActivePrinter, VARIANT PrintToFile, VARIANT Collate, VARIANT PrToFileName, VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PVARIANT ;
		InvokeHelper(0x939, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &From, &To, &Copies, &Preview, &ActivePrinter, &PrintToFile, &Collate, &PrToFileName, RHS);
		return result;
	}
	STDMETHOD(get_MDX)(BSTR * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBSTR ;
		InvokeHelper(0x84b, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(ExportAsFixedFormat)(XlFixedFormatType Type, VARIANT Filename, VARIANT Quality, VARIANT IncludeDocProperties, VARIANT IgnorePrintAreas, VARIANT From, VARIANT To, VARIANT OpenAfterPublish, VARIANT FixedFormatExtClassPtr)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT ;
		InvokeHelper(0x9bd, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, &Filename, &Quality, &IncludeDocProperties, &IgnorePrintAreas, &From, &To, &OpenAfterPublish, &FixedFormatExtClassPtr);
		return result;
	}
	STDMETHOD(get_CountLarge)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x9c3, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(CalculateRowMajorOrder)(VARIANT * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x93c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}

	// IRange 属性
public:

};
