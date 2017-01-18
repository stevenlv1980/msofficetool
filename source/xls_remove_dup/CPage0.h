// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CPage0 包装类

class CPage0 : public COleDispatchDriver
{
public:
	CPage0(){} // 调用 COleDispatchDriver 默认构造函数
	CPage0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CPage0(const CPage0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// Page 方法
public:
	LPDISPATCH get_LeftHeader()
	{
		LPDISPATCH result;
		InvokeHelper(0x3fa, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CenterHeader()
	{
		LPDISPATCH result;
		InvokeHelper(0x3f3, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_RightHeader()
	{
		LPDISPATCH result;
		InvokeHelper(0x402, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_LeftFooter()
	{
		LPDISPATCH result;
		InvokeHelper(0x3f9, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_CenterFooter()
	{
		LPDISPATCH result;
		InvokeHelper(0x3f2, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_RightFooter()
	{
		LPDISPATCH result;
		InvokeHelper(0x401, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}

	// Page 属性
public:

};
