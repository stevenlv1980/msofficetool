// 从类型库向导中用“添加类”创建的计算机生成的 IDispatch 包装类

// CConnectorFormat 包装类

class CConnectorFormat : public COleDispatchDriver
{
public:
	CConnectorFormat(){} // 调用 COleDispatchDriver 默认构造函数
	CConnectorFormat(LPDISPATCH pDispatch) : COleDispatchDriver(pDispatch) {}
	CConnectorFormat(const CConnectorFormat& dispatchSrc) : COleDispatchDriver(dispatchSrc) {}

	// 属性
public:

	// 操作
public:


	// ConnectorFormat 方法
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
	void BeginConnect(LPDISPATCH ConnectedShape, long ConnectionSite)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_I4 ;
		InvokeHelper(0x6d6, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ConnectedShape, ConnectionSite);
	}
	void BeginDisconnect()
	{
		InvokeHelper(0x6d9, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	void EndConnect(LPDISPATCH ConnectedShape, long ConnectionSite)
	{
		static BYTE parms[] = VTS_DISPATCH VTS_I4 ;
		InvokeHelper(0x6da, DISPATCH_METHOD, VT_EMPTY, NULL, parms, ConnectedShape, ConnectionSite);
	}
	void EndDisconnect()
	{
		InvokeHelper(0x6db, DISPATCH_METHOD, VT_EMPTY, NULL, NULL);
	}
	long get_BeginConnected()
	{
		long result;
		InvokeHelper(0x6dc, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_BeginConnectedShape()
	{
		LPDISPATCH result;
		InvokeHelper(0x6dd, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_BeginConnectionSite()
	{
		long result;
		InvokeHelper(0x6de, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_EndConnected()
	{
		long result;
		InvokeHelper(0x6df, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	LPDISPATCH get_EndConnectedShape()
	{
		LPDISPATCH result;
		InvokeHelper(0x6e0, DISPATCH_PROPERTYGET, VT_DISPATCH, (void*)&result, NULL);
		return result;
	}
	long get_EndConnectionSite()
	{
		long result;
		InvokeHelper(0x6e1, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	long get_Type()
	{
		long result;
		InvokeHelper(0x6c, DISPATCH_PROPERTYGET, VT_I4, (void*)&result, NULL);
		return result;
	}
	void put_Type(long newValue)
	{
		static BYTE parms[] = VTS_I4 ;
		InvokeHelper(0x6c, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, newValue);
	}

	// ConnectorFormat 属性
public:

};
