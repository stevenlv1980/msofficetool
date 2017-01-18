// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CShapeNodes 包装类

class CShapeNodes : public COleDispatchDriver
{
public:
	CShapeNodes(){} // 调用 COleDispatchDriver 默认构造函数
	CShapeNodes(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CShapeNodes(const CShapeNodes& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ShapeNodes 方法
public:
	LPDISPATCH get_Application()
	{
		LPDISPATCH result;
		InvokeHelper(0x60020000, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_Creator()
	{
		long result;
		InvokeHelper(0x60020001, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
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
	LPDISPATCH Item(VARIANT Index)
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
	void Delete(long Index)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0xb, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index);
	}
	void Insert(long Index, long SegmentType, long EditingType, float X1, float Y1, float X2, float Y2, float X3, float Y3)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 VTS_I4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 VTS_R4 ;
		InvokeHelper(0xc, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3);
	}
	void SetEditingType(long Index, long EditingType)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0xd, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, EditingType);
	}
	void SetPosition(long Index, float X1, float Y1)
	{
		static BYTE parms[] = VTS_I4 VTS_R4 VTS_R4 ;
		InvokeHelper(0xe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, X1, Y1);
	}
	void SetSegmentType(long Index, long SegmentType)
	{
		static BYTE parms[] = VTS_I4 VTS_I4 ;
		InvokeHelper(0xf, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Index, SegmentType);
	}

	// ShapeNodes 属性
public:

};
