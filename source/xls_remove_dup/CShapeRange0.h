// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CShapeRange0 包装类

class CShapeRange0 : public COleDispatchDriver
{
public:
	CShapeRange0(){} // 调用 COleDispatchDriver 默认构造函数
	CShapeRange0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CShapeRange0(const CShapeRange0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ShapeRange 方法
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
	long get_Count()
	{
		long result;
		InvokeHelper(0x76, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH Item(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xaa, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPDISPATCH _Default(VARIANT Index)
	{
		LPDISPATCH result;
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0x0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, parms, &Index);
		return result;
	}
	LPUNKNOWN get__NewEnum()
	{
		LPUNKNOWN result;
		InvokeHelper(0xfffffffc, DISPATCH_PROPERTYGET, VT_UNKNOWN, (void*)&result, NULL);
		return result;
	}
	void Align(long AlignCmd, long RelativeTo)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x6cc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, AlignCmd, RelativeTo);
	}
	void Apply()
	{
		InvokeHelper(0x68b, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Delete()
	{
		InvokeHelper(0x75, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Distribute(long DistributeCmd, long RelativeTo)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0x6ce, DISPATCH_METHOD, VT_EMPTY, NULL, parms, DistributeCmd, RelativeTo);
	}
	LPDISPATCH Duplicate()
	{
		LPDISPATCH result;
		InvokeHelper(0x40f, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void Flip(long FlipCmd)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x68c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, FlipCmd);
	}
	void IncrementLeft(float Increment)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x68e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	void IncrementRotation(float Increment)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x690, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	void IncrementTop(float Increment)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x691, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	LPDISPATCH Group()
	{
		LPDISPATCH result;
		InvokeHelper(0x2e, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void PickUp()
	{
		InvokeHelper(0x692, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void RerouteConnections()
	{
		InvokeHelper(0x693, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Regroup()
	{
		LPDISPATCH result;
		InvokeHelper(0x6d0, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ScaleHeight(float Factor, long RelativeToOriginalSize, VARIANT Scale)
	{
		static BYTE parms[] = VTS_R4 VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x694, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Factor, RelativeToOriginalSize, &Scale);
	}
	void ScaleWidth(float Factor, long RelativeToOriginalSize, VARIANT Scale)
	{
		static BYTE parms[] = VTS_R4 VTS_I4 VTS_VARIANT ;
		InvokeHelper(0x698, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Factor, RelativeToOriginalSize, &Scale);
	}
	void Select(VARIANT Replace)
	{
		static BYTE parms[] = VTS_VARIANT ;
		InvokeHelper(0xeb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, &Replace);
	}
	void SetShapesDefaultProperties()
	{
		InvokeHelper(0x699, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	LPDISPATCH Ungroup()
	{
		LPDISPATCH result;
		InvokeHelper(0xf4, DISPATCH_METHOD, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	void ZOrder(long ZOrderCmd)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x26e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ZOrderCmd);
	}
	LPDISPATCH get_Adjustments()
	{
		LPDISPATCH result;
		InvokeHelper(0x69b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_TextFrame()
	{
		LPDISPATCH result;
		InvokeHelper(0x69c, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_AutoShapeType()
	{
		long result;
		InvokeHelper(0x69d, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_AutoShapeType(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x69d, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Callout()
	{
		LPDISPATCH result;
		InvokeHelper(0x69e, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_ConnectionSiteCount()
	{
		long result;
		InvokeHelper(0x69f, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Connector()
	{
		long result;
		InvokeHelper(0x6a0, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ConnectorFormat()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a1, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Fill()
	{
		LPDISPATCH result;
		InvokeHelper(0x67f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_GroupItems()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	float get_Height()
	{
		float result;
		InvokeHelper(0x7b, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Height(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_HorizontalFlip()
	{
		long result;
		InvokeHelper(0x6a3, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	float get_Left()
	{
		float result;
		InvokeHelper(0x7f, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Left(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7f, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Line()
	{
		LPDISPATCH result;
		InvokeHelper(0x331, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_LockAspectRatio()
	{
		long result;
		InvokeHelper(0x6a4, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_LockAspectRatio(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6a4, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
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
	LPDISPATCH get_Nodes()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a5, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	float get_Rotation()
	{
		float result;
		InvokeHelper(0x3b, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Rotation(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x3b, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_PictureFormat()
	{
		LPDISPATCH result;
		InvokeHelper(0x65f, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Shadow()
	{
		LPDISPATCH result;
		InvokeHelper(0x67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_TextEffect()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a6, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ThreeD()
	{
		LPDISPATCH result;
		InvokeHelper(0x6a7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	float get_Top()
	{
		float result;
		InvokeHelper(0x7e, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Top(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7e, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_VerticalFlip()
	{
		long result;
		InvokeHelper(0x6a8, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	VARIANT get_Vertices()
	{
		VARIANT result;
		InvokeHelper(0x26d, DISPATCH_PROPERTYGET, VT_VARIANT, (void*)&result, NULL);
		return result;
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
	float get_Width()
	{
		float result;
		InvokeHelper(0x7a, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Width(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x7a, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_ZOrderPosition()
	{
		long result;
		InvokeHelper(0x6a9, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_BlackWhiteMode()
	{
		long result;
		InvokeHelper(0x6ab, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_BlackWhiteMode(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6ab, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	CString get_AlternativeText()
	{
		CString result;
		InvokeHelper(0x763, DISPATCH_PROPERTYGET, VT_BSTR, (void*)&result, NULL);
		return result;
	}
	void put_AlternativeText(LPCTSTR newValue)
	{
		static BYTE parms[] = VTS_BSTR ;
		InvokeHelper(0x763, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_DiagramNode()
	{
		LPDISPATCH result;
		InvokeHelper(0x875, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_HasDiagramNode()
	{
		long result;
		InvokeHelper(0x876, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Diagram()
	{
		LPDISPATCH result;
		InvokeHelper(0x877, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_HasDiagram()
	{
		long result;
		InvokeHelper(0x878, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Child()
	{
		long result;
		InvokeHelper(0x879, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_ParentGroup()
	{
		LPDISPATCH result;
		InvokeHelper(0x87a, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CanvasItems()
	{
		LPDISPATCH result;
		InvokeHelper(0x87b, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_ID()
	{
		long result;
		InvokeHelper(0x23a, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void CanvasCropLeft(float Increment)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x87c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	void CanvasCropTop(float Increment)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x87d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	void CanvasCropRight(float Increment)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x87e, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	void CanvasCropBottom(float Increment)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x87f, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Increment);
	}
	LPDISPATCH get_Chart()
	{
		LPDISPATCH result;
		InvokeHelper(0x7, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_HasChart()
	{
		long result;
		InvokeHelper(0xa62, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_TextFrame2()
	{
		LPDISPATCH result;
		InvokeHelper(0xa63, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_ShapeStyle()
	{
		long result;
		InvokeHelper(0xa64, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_ShapeStyle(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa64, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	long get_BackgroundStyle()
	{
		long result;
		InvokeHelper(0xa65, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_BackgroundStyle(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xa65, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_SoftEdge()
	{
		LPDISPATCH result;
		InvokeHelper(0xa66, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Glow()
	{
		LPDISPATCH result;
		InvokeHelper(0xa67, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_Reflection()
	{
		LPDISPATCH result;
		InvokeHelper(0xa68, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// ShapeRange 属性
public:

};
