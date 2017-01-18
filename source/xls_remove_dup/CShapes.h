// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CShapes 包装类

class CShapes : public COleDispatchDriver
{
public:
	CShapes(){} // 调用 COleDispatchDriver 默认构造函数
	CShapes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CShapes(const CShapes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IShapes 方法
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
	STDMETHOD(get_Count)(long * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PI4 ;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(Item)(VARIANT Index, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(_Default)(VARIANT Index, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(get__NewEnum)(LPUNKNOWN * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PUNKNOWN ;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, RHS);
		return result;
	}
	STDMETHOD(AddCallout)(MsoCalloutType Type, float Left, float Top, float Width, float Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6b1, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(AddConnector)(MsoConnectorType Type, float BeginX, float BeginY, float EndX, float EndY, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6b2, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, BeginX, BeginY, EndX, EndY, RHS);
		return result;
	}
	STDMETHOD(AddCurve)(VARIANT SafeArrayOfPoints, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x6b7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &SafeArrayOfPoints, RHS);
		return result;
	}
	STDMETHOD(AddLabel)(MsoTextOrientation Orientation, float Left, float Top, float Width, float Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6b9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Orientation, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(AddLine)(float BeginX, float BeginY, float EndX, float EndY, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6ba, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, BeginX, BeginY, EndX, EndY, RHS);
		return result;
	}
	STDMETHOD(AddPicture)(LPCTSTR Filename, MsoTriState LinkToFile, MsoTriState SaveWithDocument, float Left, float Top, float Width, float Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BSTR VTS_I4 VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6bb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(AddPolyline)(VARIANT SafeArrayOfPoints, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x6be, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &SafeArrayOfPoints, RHS);
		return result;
	}
	STDMETHOD(AddShape)(MsoAutoShapeType Type, float Left, float Top, float Width, float Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6bf, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(AddTextEffect)(MsoPresetTextEffect PresetTextEffect, LPCTSTR Text, LPCTSTR FontName, float FontSize, MsoTriState FontBold, MsoTriState FontItalic, float Left, float Top, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_BSTR VTS_BSTR VTS_R4 VTS_I4 VTS_I4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6c0, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top, RHS);
		return result;
	}
	STDMETHOD(AddTextbox)(MsoTextOrientation Orientation, float Left, float Top, float Width, float Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6c6, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Orientation, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(BuildFreeform)(MsoEditingType EditingType, float X1, float Y1, FreeformBuilder * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x6c7, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, EditingType, X1, Y1, RHS);
		return result;
	}
	STDMETHOD(get_Range)(VARIANT Index, ShapeRange * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xc5, DISPATCH_PROPERTYGET, VT_HRESULT, (void*)&result, parms, &Index, RHS);
		return result;
	}
	STDMETHOD(SelectAll)()
	{
		HRESULT result;
		InvokeHelper(0x6c9, DISPATCH_METHOD, VT_HRESULT, (void*)&result, NULL);
		return result;
	}
	STDMETHOD(AddFormControl)(XlFormControl Type, long Left, long Top, long Width, long Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_I4 VTS_PDISPATCH ;
		InvokeHelper(0x6ca, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(AddOLEObject)(VARIANT ClassType, VARIANT Filename, VARIANT Link, VARIANT DisplayAsIcon, VARIANT IconFileName, VARIANT IconIndex, VARIANT IconLabel, VARIANT Left, VARIANT Top, VARIANT Width, VARIANT Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0x6cb, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &ClassType, &Filename, &Link, &DisplayAsIcon, &IconFileName, &IconIndex, &IconLabel, &Left, &Top, &Width, &Height, RHS);
		return result;
	}
	STDMETHOD(AddDiagram)(MsoDiagramType Type, float Left, float Top, float Width, float Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x880, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Type, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(AddCanvas)(float Left, float Top, float Width, float Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_PDISPATCH ;
		InvokeHelper(0x881, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Left, Top, Width, Height, RHS);
		return result;
	}
	STDMETHOD(AddChart)(VARIANT XlChartType, VARIANT Left, VARIANT Top, VARIANT Width, VARIANT Height, Shape * * RHS)
	{
		HRESULT result;
		static BYTE parms[] = VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_VARIANT VTS_PDISPATCH ;
		InvokeHelper(0xa69, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, &XlChartType, &Left, &Top, &Width, &Height, RHS);
		return result;
	}

	// IShapes 属性
public:

};
