// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CCanvasShapes 包装类

class CCanvasShapes : public COleDispatchDriver
{
public:
	CCanvasShapes(){} // 调用 COleDispatchDriver 默认构造函数
	CCanvasShapes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CCanvasShapes(const CCanvasShapes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// CanvasShapes 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x1f40, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x1f41, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Parent()
	{
		LPDISPATCH result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Count()
	{
		long result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	LPDISPATCH AddCallout(long Type, float Left, float Top, float Width, float Height)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0xa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, Left, Top, Width, Height);
		return result;
	}
	LPDISPATCH AddConnector(long Type, float BeginX, float BeginY, float EndX, float EndY)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, BeginX, BeginY, EndX, EndY);
		return result;
	}
	LPDISPATCH AddCurve(VARIANT * SafeArrayOfPoints)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, SafeArrayOfPoints);
		return result;
	}
	LPDISPATCH AddLabel(long Orientation, float Left, float Top, float Width, float Height)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Orientation, Left, Top, Width, Height);
		return result;
	}
	LPDISPATCH AddLine(float BeginX, float BeginY, float EndX, float EndY)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, BeginX, BeginY, EndX, EndY);
		return result;
	}
	LPDISPATCH AddPicture(LPCTSTR FileName, VARIANT * LinkToFile, VARIANT * SaveWithDocument, VARIANT * Left, VARIANT * Top, VARIANT * Width, VARIANT * Height)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_BSTR VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT VTS_PVARIANT ;
		InvokeHelper(0xf, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height);
		return result;
	}
	LPDISPATCH AddPolyline(VARIANT * SafeArrayOfPoints)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x10, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, SafeArrayOfPoints);
		return result;
	}
	LPDISPATCH AddShape(long Type, float Left, float Top, float Width, float Height)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0x11, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Type, Left, Top, Width, Height);
		return result;
	}
	LPDISPATCH AddTextEffect(long PresetTextEffect, LPCTSTR Text, LPCTSTR FontName, float FontSize, long FontBold, long FontItalic, float Left, float Top)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_BSTR VTS_BSTR VTS_R4 VTS_I4 VTS_I4 VTS_R4 VTS_R4 ;
		InvokeHelper(0x12, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top);
		return result;
	}
	LPDISPATCH AddTextbox(long Orientation, float Left, float Top, float Width, float Height)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0x13, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Orientation, Left, Top, Width, Height);
		return result;
	}
	LPDISPATCH BuildFreeform(long EditingType, float X1, float Y1)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 ;
		InvokeHelper(0x14, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, EditingType, X1, Y1);
		return result;
	}
	LPDISPATCH Range(VARIANT * Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_PVARIANT ;
		InvokeHelper(0x15, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, Index);
		return result;
	}
	void SelectAll()
	{
		InvokeHelper(0x16, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}

	// CanvasShapes 属性
public:

};
