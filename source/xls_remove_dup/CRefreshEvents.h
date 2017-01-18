// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRefreshEvents 包装类

class CRefreshEvents : public COleDispatchDriver
{
public:
	CRefreshEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CRefreshEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRefreshEvents(const CRefreshEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// IRefreshEvents 方法
public:
	STDMETHOD(BeforeRefresh)(BOOL * Cancel)
	{
		HRESULT result;
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x63c, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Cancel);
		return result;
	}
	STDMETHOD(AfterRefresh)(BOOL Success)
	{
		HRESULT result;
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x63d, DISPATCH_METHOD, VT_HRESULT, (void*)&result, parms, Success);
		return result;
	}

	// IRefreshEvents 属性
public:

};
