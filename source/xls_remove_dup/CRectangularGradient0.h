// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRectangularGradient0 包装类

class CRectangularGradient0 : public COleDispatchDriver
{
public:
	CRectangularGradient0(){} // 调用 COleDispatchDriver 默认构造函数
	CRectangularGradient0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRectangularGradient0(const CRectangularGradient0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// RectangularGradient 方法
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
	LPDISPATCH get_ColorStops()
	{
		LPDISPATCH result;
		InvokeHelper(0xac9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	double get_RectangleTop()
	{
		double result;
		InvokeHelper(0xaca, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
		return result;
	}
	void put_RectangleTop(double newValue)
	{
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0xaca, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	double get_RectangleBottom()
	{
		double result;
		InvokeHelper(0xacb, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
		return result;
	}
	void put_RectangleBottom(double newValue)
	{
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0xacb, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	double get_RectangleLeft()
	{
		double result;
		InvokeHelper(0xacc, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
		return result;
	}
	void put_RectangleLeft(double newValue)
	{
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0xacc, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	double get_RectangleRight()
	{
		double result;
		InvokeHelper(0xacd, DISPATCH_PROPERTYGET, VT_R8, (void*)&result, NULL);
		return result;
	}
	void put_RectangleRight(double newValue)
	{
		static BYTE parms[] = VTS_R8 ;
		InvokeHelper(0xacd, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// RectangularGradient 属性
public:

};
