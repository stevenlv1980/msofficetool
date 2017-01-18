// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CDocEvents 包装类

class CDocEvents : public COleDispatchDriver
{
public:
	CDocEvents(){} // 调用 COleDispatchDriver 默认构造函数
	CDocEvents(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CDocEvents(const CDocEvents& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// DocEvents 方法
public:
	void SelectionChange(LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x607, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target);
	}
	void BeforeDoubleClick(LPDISPATCH Target, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x601, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target, Cancel);
	}
	void BeforeRightClick(LPDISPATCH Target, BOOL * Cancel)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_PBOOL ;
		InvokeHelper(0x5fe, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target, Cancel);
	}
	void Activate()
	{
		InvokeHelper(0x130, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Deactivate()
	{
		InvokeHelper(0x5fa, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Calculate()
	{
		InvokeHelper(0x117, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void Change(LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x609, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target);
	}
	void FollowHyperlink(LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x5be, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target);
	}
	void PivotTableUpdate(LPDISPATCH Target)
	{
		static BYTE parms[] = VTS_DISPATCH ;
		InvokeHelper(0x86c, DISPATCH_METHOD, VT_EMPTY, NULL, parms, Target);
	}

	// DocEvents 属性
public:

};
