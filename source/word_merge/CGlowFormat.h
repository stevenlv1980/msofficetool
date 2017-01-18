// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CGlowFormat 包装类

class CGlowFormat : public COleDispatchDriver
{
public:
	CGlowFormat(){} // 调用 COleDispatchDriver 默认构造函数
	CGlowFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CGlowFormat(const CGlowFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// GlowFormat 方法
public:
	float get_Radius()
	{
		float result;
		InvokeHelper(0x1, DISPATCH_PROPERTYGET, VT_R4, (void*)&result, NULL);
		return result;
	}
	void put_Radius(float newValue)
	{
		static BYTE parms[] = VTS_R4 ;
		InvokeHelper(0x1, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}
	LPDISPATCH get_Color()
	{
		LPDISPATCH result;
		InvokeHelper(0x2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// GlowFormat 属性
public:

};
