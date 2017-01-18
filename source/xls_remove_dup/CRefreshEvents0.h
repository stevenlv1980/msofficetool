// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CRefreshEvents0 包装类

class CRefreshEvents0 : public COleDispatchDriver
{
public:
	CRefreshEvents0(){} // 调用 COleDispatchDriver 默认构造函数
	CRefreshEvents0(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CRefreshEvents0(const CRefreshEvents0& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// RefreshEvents 方法
public:
	void BeforeRefresh(BOOL * Cancel)
	{
		static BYTE parms[] = VTS_PBOOL ;
		InvokeHelper(0x63c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Cancel);
	}
	void AfterRefresh(BOOL Success)
	{
		static BYTE parms[] = VTS_BOOL ;
		InvokeHelper(0x63d, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Success);
	}

	// RefreshEvents 属性
public:

};
