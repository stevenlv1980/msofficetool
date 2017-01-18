
#include "stdafx.h"
#include "xls_remove_dup.h"
#ifdef _DEBUG
#define new DEBUG_NEW
#endif
#include "atlsafe.h"

#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets0.h"
#include "CRange0.h"
//#include "CSelection.h"
//#include "CRange.h"

// 唯一的应用程序对象
#include "get_rest.h"

CWinApp theApp;

using namespace std;


CApplication appexcel;

BOOL CreateDispatch()
{
   COleException* pe = new COleException;

   try 
   {
      // Create instance of Microsoft System Information Control 
      // by using ProgID.
      if (appexcel.CreateDispatch(_T("Excel.Application"), pe))
      {
         //_tprintf(_T("Failed to initialize OLE\n"));
		 //appexcel.put_Visible(TRUE);
		 return TRUE;
      }
      else
      {
         throw pe;
      }
   }
   //Catch control-specific exceptions.
    catch (COleDispatchException* pe) 
   {
      CString cStr;

      if (!pe->m_strSource.IsEmpty())
         cStr = pe->m_strSource + _T(" - ");
      if (!pe->m_strDescription.IsEmpty())
         cStr += pe->m_strDescription;
      else
         cStr += _T("unknown error");

      //AfxMessageBox(cStr, MB_OK, (pe->m_strHelpFile.IsEmpty()) ?  0 : pe->m_dwHelpContext);

      pe->Delete();
   }
   //Catch all MFC exceptions, including COleExceptions.
   // OS exceptions will not be caught.
   catch (CException* pe) 
   {
      TRACE(_T("%s(%d): OLE Execption caught: SCODE = %x"), 
         __FILE__, __LINE__, COleException::Process(pe));
      pe->Delete();
   }

   pe->Delete();

   return FALSE;
}

VARIANT & getVstr(VARIANT &v, TCHAR* str)
{
	CString s(str);
	v.vt = VT_BSTR;
	v.bstrVal = s.AllocSysString();
	return v;	
}

int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{



#ifdef CRASH_TEST

#endif
	int nRetCode = 0;

	::CoInitialize(NULL);
	//CoUninitialize(); 
	
	TCHAR* input_file = argv[1];
	
	// Initialize OLE libraries
	if (!AfxOleInit())
	{
		_tprintf(_T("Failed to initialize OLE\n"));
		//nRetCode = 1;
		return -1;
	}

	// 初始化 MFC 并在失败时显示错误
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
	{
		// TODO: 更改错误代码以符合您的需要
		_tprintf(_T("致命错误: MFC 初始化失败\n"));
		nRetCode = 1;
	}
	else
	{

#ifdef COMBINE_SHEETS
		//combine worksheet in excel
		if ( argc < 3 ) 
		{	
			_tprintf(_T("no input file. cmd format: combine_sheets.exe input1.xls input1.xls output.xls \n"));
			return -1;
		}
		TCHAR* output_file = argv[argc - 1];

		bool remove_empty = false;
		for(int x = 1; x < argc; x++)
		{
			CString str(argv[x]);
			if ( str == "remove_empty" )
			{
				remove_empty = true;
				break;
			}
		}

		BOOL ret;
		
		if ( !CreateDispatch() )
			return 1;

		try
		{		
			COleVariant vtOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR),
				vtTrue((short)TRUE),
				vtFalse((short)FALSE);
			CWorkbooks		works;
			CWorkbook		workout;
			CWorkbook		workin;
			CWorksheets0	sheetsout;
			CWorksheets0	sheets;
			CWorksheet		sheet;
			CWorksheet		sheetout;
			CRange0			range;

			works = appexcel.get_Workbooks();
			workout = works.Add(vtOptional);
			sheetsout = workout.get_Worksheets();
			sheetout = sheetsout.get_Item(COleVariant((short)(1)) ); 

			for(int x = 1; x < argc - 1; x++)
			{
				_tprintf(argv[x]);
				_tprintf(_T("\n"));
				CString str(argv[x]);
				if ( str == "remove_empty" )
				{
					continue;
				}
				
				workin = works.Open(str,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional);
				sheets = workin.get_Worksheets();
				int sheetct = sheets.get_Count();
				for( int x = 0; x < sheetct; x++ )
				{
					sheet = sheets.get_Item(COleVariant((short)(x+1)));
					range = sheet.get_UsedRange();
					CString strx = range.get_Address(vtOptional,vtOptional,	1,vtOptional,vtOptional);
					if ( strx == "$A$1" ) 
					{
						//empty sheet
						if ( remove_empty )
						{
							continue;
						}
					}
	
					VARIANT ct;
					ct.vt = VT_DISPATCH;
					ct.pdispVal = sheetout.m_lpDispatch;
					sheet.Copy(ct,vtOptional);
				}
				workin.Save();
				workin.Close(vtOptional,vtOptional,vtOptional);
			}

			if ( remove_empty )
			{
				//remove empty sheet
				int sheetct = sheetsout.get_Count();
				for( int x = sheetct; x > 0; x-- )
				{
					sheet = sheetsout.get_Item(COleVariant((short)(x)));
					range = sheet.get_UsedRange();
					CString strx = range.get_Address(vtOptional,vtOptional,	1,vtOptional,vtOptional);
					if ( strx == "$A$1" ) 
					{
						//empty sheet
						if ( remove_empty )
						{
							sheet.Delete();
							//sheetsout
						}
					}
				}
			}	
			VARIANT v;
			getVstr(v,output_file);
			workout.SaveAs(v,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,1,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional);
			appexcel.Quit();//Quit(vtFalse,vtOptional,vtOptional);
			appexcel.ReleaseDispatch();
		}
		catch(CException* pe)
		{
			char buf[1024];
			memset( buf, 0, sizeof(buf) );
			pe->GetErrorMessage( buf,1024 );
			//::MessageBox(NULL, buf, "COleException", MB_SETFOREGROUND | MB_OK);
			
			_tprintf(_T("Failed to remove\n"));
			TRACE(_T("%s(%d): OLE Execption caught: SCODE = %x"), 
				__FILE__, __LINE__, COleException::Process(pe));/**/
			pe->Delete();
		}
		
#elif defined SPLIT_WORKBOOK
		//split workbook to individual files
		if ( argc < 3 ) 
		{	
			_tprintf(_T("no input file. cmd format: combine_sheets.exe input1.xls c:\\output\\ \n"));
			return -1;
		}
		TCHAR* output_path = argv[argc - 1];
		CString stroutput(output_path);

		bool remove_empty = false;
		for(int x = 1; x < argc; x++)
		{
			CString str(argv[x]);
			if ( str == "remove_empty" )
			{
				remove_empty = true;
				break;
			}
		}
		CString strtmp(input_file);
		CString exten;

		exten = strtmp.Right( strtmp.GetLength() - strtmp.ReverseFind('.') );

		BOOL ret;
		
		if ( !CreateDispatch() )
			return 1;


		try
		{		
			COleVariant vtOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR),
				vtTrue((short)TRUE),
				vtFalse((short)FALSE);
		
			CWorkbooks		works;
			CWorkbook		work;
			CWorksheets0	sheets;
			CWorksheet		sheet;

			CWorkbook		worknew;
			CWorksheets0	sheetsnew;
			CWorksheet		sheetnew;
			CRange0			range;

			works = appexcel.get_Workbooks();
			work = works.Open(input_file,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional);
			sheets = work.get_Worksheets();
			int sheetct = sheets.get_Count();
			
			for( int x = 0; x < sheetct; x++ )
			{
				sheet = sheets.get_Item(COleVariant((short)(x+1)) ) ;
				range = sheet.get_UsedRange();
				CString strx = range.get_Address(vtOptional,vtOptional,	1,vtOptional,vtOptional);
				if ( strx == "$A$1" ) 
				{
					//empty sheet
					if ( remove_empty )
					{
						continue;
					}
				}
				//copy to new workbook
				worknew = works.Add(vtOptional);
				sheetsnew = worknew.get_Worksheets();
				sheetnew = sheetsnew.get_Item(COleVariant((short)(1)) ); 
				VARIANT ct;
				ct.vt = VT_DISPATCH;
				ct.pdispVal = sheetnew.m_lpDispatch;
				sheet.Copy(ct,vtOptional);
				//remove empty worksheet in new workbook
				if ( remove_empty )
				{
					int sheetct = sheetsnew.get_Count();
					for( int x = sheetct; x > 0; x-- )
					{
						sheetnew = sheetsnew.get_Item(COleVariant((short)(x)));
						range = sheetnew.get_UsedRange();
						CString strx = range.get_Address(vtOptional,vtOptional,	1,vtOptional,vtOptional);
						if ( strx == "$A$1" ) 
						{
							//empty sheet
							sheetnew.Delete();
						}
					}
				}
				//save to file
				CString filename = sheet.get_Name();
				CString fileoutput = stroutput + filename + exten;
				
				VARIANT v;
				getVstr(v,(TCHAR*)(LPCTSTR)fileoutput);
				worknew.SaveAs(v,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,1,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional);
			}

			work.Save();
			work.Close(vtFalse,vtOptional,vtOptional);
			appexcel.Quit();//Quit(vtFalse,vtOptional,vtOptional);
			appexcel.ReleaseDispatch();
		}
		catch(CException* pe)
		{
			char buf[1024];
			memset( buf, 0, sizeof(buf) );
			pe->GetErrorMessage( buf,1024 );
			//::MessageBox(NULL, buf, "COleException", MB_SETFOREGROUND | MB_OK);
			
			_tprintf(_T("Failed to remove\n"));
			TRACE(_T("%s(%d): OLE Execption caught: SCODE = %x"), 
				__FILE__, __LINE__, COleException::Process(pe));/**/
			pe->Delete();
		}
		
#else
		//remove duplicate rows
		if ( argc < 3 ) 
		{	
			_tprintf(_T("no input file. cmd format: xls_remove_dup.exe input1.xls output.xls \n"));
			return -1;
		}
		
		TCHAR* output_file = argv[2];

		BOOL ret;
		
		if ( !CreateDispatch() )
			return 1;
		
		try
		{
			COleVariant vtOptional((long)DISP_E_PARAMNOTFOUND,VT_ERROR),
					vtTrue((short)TRUE),
					vtFalse((short)FALSE);
			CWorkbooks		works;
			CWorkbook		work;
			CWorksheets0	sheets;
			CWorksheet		sheet;
			CRange0			range;

			works = appexcel.get_Workbooks();
			work = works.Open(input_file,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional);
			sheets = work.get_Worksheets();
			int sheetct = sheets.get_Count();
			
			for( int x = 0; x < sheetct; x++ )
			{
				sheet = sheets.get_Item(COleVariant((short)(x+1)) ) ;
				CRange0 cells = sheet.get_Cells();
				

				cells.RemoveDuplicates( vtTrue,1);
			}

			/*VARIANT v;
			getVstr(v,input_file);
			*/
			work.Save();//(&v,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional);
			
			VARIANT v;
			getVstr(v,output_file);
			work.SaveAs(v,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional,1,vtOptional,vtOptional,vtOptional,vtOptional,vtOptional);
		
			work.Close(vtFalse,vtOptional,vtOptional);
			appexcel.Quit();//Quit(vtFalse,vtOptional,vtOptional);
			appexcel.ReleaseDispatch();
		}
		catch(CException* pe)
		{
			char buf[1024];
			memset( buf, 0, sizeof(buf) );
			pe->GetErrorMessage( buf,1024 );
			//::MessageBox(NULL, buf, "COleException", MB_SETFOREGROUND | MB_OK);
			
			_tprintf(_T("Failed to remove\n"));
			TRACE(_T("%s(%d): OLE Execption caught: SCODE = %x"), 
				__FILE__, __LINE__, COleException::Process(pe));/**/
			pe->Delete();
		}
		
#endif

	}

	return nRetCode;
}

				/*CRange0 rows = cells.get_Rows();
				CRange0 cols = cells.get_Columns();

				int nRows = rows.get_Count();
				int nCols = cols.get_Count();				

				sheet.Activate();
				
				range = sheet.get_Range( COleVariant("A1"),COleVariant("C10"));
				*/

				//range.Clear();
				//sheet.get_Cells();
				/*CComSafeArray<int> out_sa;
				out_sa.Create(3);
				out_sa.SetAt(0,1);
				out_sa.SetAt(1,2);
				out_sa.SetAt(1,3);
				int colums[3];
				colums[0] = 1;
				colums[1] = 2;
				colums[2] = 3;
				
				
				COleSafeArray saRet;
				DWORD numElements[3];
				numElements[0] = 1;
				numElements[1] = 2;
				numElements[2] = 3;
				saRet.Create(VT_R4,1,numElements);
				int v  = 1;	
				long index = 0;
				*/

				/*v  = 1;index = 0;saRet.PutElement(&index, &v);
				v  = 2;index = 1;saRet.PutElement(&index, &v);
				v  = 3;index = 2;saRet.PutElement(&index, &v);*/