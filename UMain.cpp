//---------------------------------------------------------------------------

#include <vcl.h>
#include <assert.h>
#pragma hdrstop

#include "UMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TFMain *FMain;
#define MAX_KEYLEN 256
inline AnsiString Variant2Str(VARIANT& v)
{
   Variant var(v);
   return VarToStr(var);
}
#ifdef StrToInt
#undef StrToInt
#endif // StrToInt
//---------------------------------------------------------------------------
__fastcall TFMain::TFMain(TComponent* Owner)
   : TForm(Owner),
   m_vbActive(VARIANT_FALSE),
   m_hGroup(0),
   m_dwRate(100),
   m_fDeadBand(0.0f),
   m_hItem(0),
   m_dwID(0),
   m_dwCancelID(0)
{
   CoGetMalloc(MEMCTX_TASK, &m_ptrMalloc);
   assert(m_ptrMalloc != NULL);
}
//---------------------------------------------------------------------------
void __fastcall TFMain::FormShow(TObject *Sender)
{
        BtnServerList->Click();        
}
//---------------------------------------------------------------------------
void __fastcall TFMain::BtnServerListClick(TObject *Sender)
{
        GetServer();
}
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
void __fastcall TFMain::GetServer()
{
	HKEY hk = HKEY_CLASSES_ROOT;
	TCHAR szKey[MAX_KEYLEN];
	for(int nIndex = 0; ::RegEnumKey(hk, nIndex, szKey, MAX_KEYLEN) == ERROR_SUCCESS; nIndex++)
	{
		HKEY hProgID;
		TCHAR szDummy[MAX_KEYLEN];
		if(::RegOpenKey(hk, szKey, &hProgID) == ERROR_SUCCESS)
		{
			LONG lSize = MAX_KEYLEN;
			if(::RegQueryValue(hProgID, "OPC", szDummy, &lSize) == ERROR_SUCCESS)
			{
                        LstServer->Items->Add(szKey);
			}
			::RegCloseKey(hProgID);
		}
	}

        if (LstServer->Items->Count > 0)
        {
              BtnServerList->Visible = false;
              StatusBar->SimpleText ="选择服务器连接。。。";
        }
        else
              StatusBar->SimpleText ="没有找到OPC服务器";

}
void __fastcall TFMain::FormDestroy(TObject *Sender)
{
        Cleanup();
}
//---------------------------------------------------------------------------
void __fastcall TFMain::CleanupItem()
{
   if (0 != m_hItem)
   {
      assert(m_ptrGroup != NULL);
      // Get the item management interface of the group
      CComPtr<IOPCItemMgt> ptrItMgm;
      OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCItemMgt,
                              reinterpret_cast<LPVOID*>(&ptrItMgm)));
      assert(ptrItMgm != NULL);

      HRESULT* phResult = NULL;
      OLECHECK(ptrItMgm->RemoveItems(1, &m_hItem, &phResult));

      // Check the item result for errors
      assert(phResult != NULL);
      HRESULT hr = phResult[0];
      m_ptrMalloc->Free(phResult);
      OLECHECK((HRESULT)hr);

      // Final cleanup
      m_hItem = 0;
   }
}

//---------------------------------------------------------------------------
void __fastcall TFMain::Cleanup()
{
   // Free the current item
   CleanupItem();

   if (0 != m_hGroup)
   {
      // Disconnect group events
      m_DataCallbackSink.Disconnect();
      // Release the group
      m_ptrSyncIO.Release();
      m_ptrAsyncIO.Release();
      m_ptrGroup.Release();
      // Remove the group from the server itself
      OLECHECK(m_ptrServer->RemoveGroup(m_hGroup, FALSE));
      m_hGroup = 0;
   }
   assert(m_ptrSyncIO == NULL);
   assert(m_ptrAsyncIO == NULL);
   assert(m_ptrGroup == NULL);

   // Release the OPC server
   m_ptrServer.Release();

   // Set Deactivate Flag
   m_vbActive = VARIANT_FALSE;
}
void __fastcall TFMain::BtnConnectClick(TObject *Sender)
{
   static const LPCTSTR szConnect = _T("&Connect");
   static const LPCTSTR szDisconnect = _T("Dis&connect");

   if (BtnConnect->Tag == 0)
   {
      try
      {
         m_bstrServer.Empty();
         m_bstrServer = WideString(EdtServer->Text);

         ConnectToServer();
	 StatusBar->SimpleText ="连接到: "+EdtServer->Text +" 选择项目...";
         GetItems();

         BtnConnect->Tag = 1;
         BtnConnect->Caption = "断开";

         EnableButtons(true);
      }
      catch(...)
      {
         Cleanup();
         throw;
      }
   }
   else
   {
      try
      {
         Cleanup();

         BtnConnect->Tag = 0;
         BtnConnect->Caption = szConnect;

	 StatusBar->SimpleText ="断开...";
         EnableButtons(false);
      }
      catch(...)
      {
         throw;
      }
   }
}
//---------------------------------------------------------------------------
void __fastcall TFMain::ConnectToServer()
{
   assert(!m_vbActive);
   assert(m_bstrServer.Length() != 0);

   // INITIALIZATION:
   assert(m_ptrMalloc != NULL);

   // Create the OPC Server:
   CLSID clsid;
   OLECHECK(CLSIDFromProgID(m_bstrServer, &clsid));
   OLECHECK(CoCreateInstance(clsid, NULL, CLSCTX_ALL, IID_IOPCServer,
                                     reinterpret_cast<LPVOID*>(&m_ptrServer)));

   OLECHECK(m_ptrServer->QueryInterface(IID_IOPCBrowseServerAddressSpace,
                                        reinterpret_cast<LPVOID*>(&m_ptrBrowse)));
   // Create a group with a unique name:
   LPOLESTR pszGRPID = NULL;
   try
   {
      GUID guidGroupName;
      OLECHECK(CoCreateGuid(&guidGroupName));
      OLECHECK(StringFromCLSID(clsid, &pszGRPID));
      assert(pszGRPID != NULL);
      DWORD dwRevisedRate = 0;
      OLECHECK(m_ptrServer->AddGroup(pszGRPID, TRUE, m_dwRate, 0, 0, &m_fDeadBand,
			      0, &m_hGroup, &dwRevisedRate, IID_IUnknown,
   			   reinterpret_cast<LPUNKNOWN*>(&m_ptrGroup)));
      m_dwRate = dwRevisedRate;
      assert(m_ptrGroup != NULL);
      CoTaskMemFree(pszGRPID);
   }
   catch(...)
   {
      if (pszGRPID != NULL)
      {
         CoTaskMemFree(pszGRPID);
      }
      throw;
   }
   /*
   // Get the sync IO interface of the group:
   OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCSyncIO,
                           reinterpret_cast<LPVOID*>(&m_ptrSyncIO)));

   // Get the async IO interface of the group:
   OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCAsyncIO2,
                           reinterpret_cast<LPVOID*>(&m_ptrAsyncIO)));

   // Connect the event handlers to those of this form:
   m_DataCallbackSink.EvDataChange = OnDataChange;
   m_DataCallbackSink.EvReadComplete = OnReadComplete;
   m_DataCallbackSink.EvWriteComplete = OnWriteComplete;
   m_DataCallbackSink.EvCancelComplete = OnCancelComplete;

   // Connect the IOPCDataCallback sink to the group
   OLECHECK(m_DataCallbackSink.Connect(m_ptrGroup));
   */
}
void __fastcall TFMain::GetItems()
{
	// loop until all items are added
	char sz2[200];
	TCHAR szBuffer[256];
	HRESULT hr = 0;
	int nTestItem = 0; // how many items there are
	IEnumString* pEnumString = NULL;
	int nCount = 0;
	USES_CONVERSION;
        LstItem->Items->Clear();

   OLECHECK(m_ptrBrowse->BrowseOPCItemIDs(OPC_FLAT, L""/*NULL*/, VT_EMPTY, 0, &pEnumString));
	LPOLESTR pszName = NULL;
        ULONG count = 0;
        while((hr = pEnumString->Next(1, &pszName, &count)) == S_OK)
        {
            LstItem->Items->Add(OLE2T(pszName));
            ::CoTaskMemFree(pszName);
        }
        pEnumString->Release();
}
void __fastcall TFMain::EnableButtons(bool bEnabled)
{
   BtnRead->Enabled = bEnabled;
   BtnWrite->Enabled = bEnabled;
}
void __fastcall TFMain::DoLog(LPCTSTR pszMsg)
{
   MemLog->Lines->Add(pszMsg);
}

void __fastcall TFMain::OnDataChange(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
   DoLog(Format(_T("DataChange: Value = '%s'"),
               ARRAYOFCONST((Variant2Str(*pvValues)))).c_str());
}

//---------------------------------------------------------------------------
void __fastcall TFMain::OnReadComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMasterquality,
      /* [in] */ HRESULT hrMastererror,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *phClientItems,
      /* [size_is][in] */ VARIANT __RPC_FAR *pvValues,
      /* [size_is][in] */ WORD __RPC_FAR *pwQualities,
      /* [size_is][in] */ FILETIME __RPC_FAR *pftTimeStamps,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
   DoLog(Format(_T("ReadComplete: Value = '%s'"),
               ARRAYOFCONST((Variant2Str(*pvValues)))).c_str());
}

//---------------------------------------------------------------------------
void __fastcall TFMain::OnWriteComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup,
      /* [in] */ HRESULT hrMastererr,
      /* [in] */ DWORD dwCount,
      /* [size_is][in] */ OPCHANDLE __RPC_FAR *pClienthandles,
      /* [size_is][in] */ HRESULT __RPC_FAR *pErrors)
{
   DoLog(_T("Async Write completed!"));
}

//---------------------------------------------------------------------------
void __fastcall TFMain::OnCancelComplete(
      /* [in] */ DWORD dwTransid,
      /* [in] */ OPCHANDLE hGroup)
{
}
void __fastcall TFMain::LstServerClick(TObject *Sender)
{
        EdtServer->Text=LstServer->Items->Strings[LstServer->ItemIndex];
}
//---------------------------------------------------------------------------

void __fastcall TFMain::LstItemClick(TObject *Sender)
{
        EdtItem->Text=LstItem->Items->Strings[LstItem->ItemIndex];
        m_bstrItem.Empty();
        m_bstrItem = WideString(EdtItem->Text);
        ConnectToItem();
}
//---------------------------------------------------------------------------
void __fastcall TFMain::ConnectToItem()
{
   //assert(!m_vbActive);
   //assert(m_ptrGroup != NULL);
   //assert(m_bstrItem.Length() != 0);

   // Get the item management interface of the group
   CComPtr<IOPCItemMgt> ptrItMgm;
   OLECHECK(m_ptrGroup->QueryInterface(IID_IOPCItemMgt,
                           reinterpret_cast<LPVOID*>(&ptrItMgm)));

	OPCITEMDEF itemdef;
	HRESULT *phResult = NULL;
	OPCITEMRESULT *pItemState = NULL;

	// Define one item
	//
   USES_CONVERSION;
	itemdef.szItemID = m_bstrItem;
	itemdef.szAccessPath = T2OLE(_T(""));
	itemdef.bActive = TRUE;
	itemdef.hClient = reinterpret_cast<DWORD>(Handle);
	itemdef.dwBlobSize = 0;
	itemdef.pBlob = NULL;
	itemdef.vtRequestedDataType = 0;

	// Add then items and check the hresults
	//
	OLECHECK(ptrItMgm->AddItems(1, &itemdef, &pItemState, &phResult));

   // Check the item result for errors
   assert(phResult != NULL);

   if (SUCCEEDED(phResult[0]))
   {
      // Store item server handle for future use
      assert(pItemState != NULL);
      m_hItem = pItemState[0].hServer;

      if (pItemState[0].pBlob != NULL)
      {
         m_ptrMalloc->Free(pItemState[0].pBlob);
	   }

	   // Free the returned results
	   //
      m_ptrMalloc->Free(phResult);
	   m_ptrMalloc->Free(pItemState);

      // We always must call the SetEnabled of AsyncIO2:
      assert(m_ptrAsyncIO != NULL);
      OLECHECK(m_ptrAsyncIO->SetEnable(TRUE));

      // If it gets here, the server is active
      m_vbActive = VARIANT_TRUE;
   }
   else
   {
      HRESULT hr = phResult[0];

	   // Free the returned results
	   //
      m_ptrMalloc->Free(phResult);
	   m_ptrMalloc->Free(pItemState);

      OLECHECK((HRESULT)hr);// Always the exception is raised here
   }
}
void __fastcall TFMain::BtnReadClick(TObject *Sender)
{
        try
   {
      assert(m_ptrMalloc != NULL);
      assert(m_ptrSyncIO != NULL);

      HRESULT *phResult = NULL;
      OPCITEMSTATE* pItemState;
      OLECHECK(m_ptrSyncIO->Read(OPC_DS_DEVICE, 1, &m_hItem, &pItemState, &phResult));

      assert(phResult != NULL);
      if (SUCCEEDED(phResult[0]))
      {
         assert(pItemState != NULL);

         DoLog(Format(_T("Synchronous Read: Value = '%s'"),
                     ARRAYOFCONST((Variant2Str(pItemState[0].vDataValue)))).c_str());

         // Free results
         m_ptrMalloc->Free(pItemState);
         m_ptrMalloc->Free(phResult);
      }
      else
      {
         HRESULT hr = phResult[0];

	      // Free the returned results
	      //
         m_ptrMalloc->Free(phResult);
	      m_ptrMalloc->Free(pItemState);

         OLECHECK(hr);// Always the exception is raised here
      }
   }
   catch(...)
   {
   }
}
//---------------------------------------------------------------------------

void __fastcall TFMain::BtnWriteClick(TObject *Sender)
{
        // Only integers, by now:
   CComVariant vValue = Sysutils::StrToInt(EdtValue->Text);

   assert(m_ptrMalloc != NULL);
   assert(m_ptrSyncIO != NULL);

   HRESULT *phResult = NULL;
   OLECHECK(m_ptrSyncIO->Write(1, &m_hItem, &vValue, &phResult));

   // Check the item result for errors
   assert(phResult != NULL);
   HRESULT hr = phResult[0];
   m_ptrMalloc->Free(phResult);
   OLECHECK((HRESULT)hr);
}
//---------------------------------------------------------------------------

void __fastcall TFMain::Button1Click(TObject *Sender)
{
        try
   {
      m_bstrItem.Empty();
      m_bstrItem = WideString(EdtItem->Text);
      ConnectToItem();
      assert(m_ptrMalloc != NULL);
      assert(m_ptrSyncIO != NULL);

      HRESULT *phResult = NULL;
      OPCITEMSTATE* pItemState;
      OLECHECK(m_ptrSyncIO->Read( OPC_DS_DEVICE, 1, &m_hItem, &pItemState, &phResult));
      assert(phResult != NULL);
      if (SUCCEEDED(phResult[0]))
      {
         assert(pItemState != NULL);

         //{
                bool OPCReadValueX=pItemState[0].vDataValue.bVal;
         //}
         //else if(Type==2)
         //{
         //       OPCReadValueD=pItemState[0].vDataValue.dblVal;
         //}


         //DoLog(Format(_T("Synchronous Read: Value = '%s'"),
         //            ARRAYOFCONST((Variant2Str(pItemState[0].vDataValue)))).c_str());

         // Free results
         m_ptrMalloc->Free(pItemState);
         m_ptrMalloc->Free(phResult);
      }
      else
      {
         HRESULT hr = phResult[0];

	      // Free the returned results
	      //
         m_ptrMalloc->Free(phResult);
	      m_ptrMalloc->Free(pItemState);

         OLECHECK(hr);// Always the exception is raised here
      }
   }
   catch(...)
   {
   }
}
//---------------------------------------------------------------------------

