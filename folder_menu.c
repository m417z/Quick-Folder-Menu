#include <windows.h>
#include <shtypes.h>
#include <shobjidl.h>
#include <shlguid.h>
#include <shlobj.h>
#include <initguid.h>
#include "buffer.h"

#ifndef	SMSET_USEBKICONEXTRACTION
#define SMSET_USEBKICONEXTRACTION	0x00000008
#endif

// {ECD4FC4F-521C-11D0-B792-00A0C90312E1}
DEFINE_GUID(CLSID_MenuDeskBar, 0xecd4fc4f, 0x521c, 0x11d0, 0xb7, 0x92, 0x0, 0xa0, 0xc9, 0x3, 0x12, 0xe1);

IMenuBand *PopupMenu(long nX, long nY, WCHAR *pPath);

void main()
{
	if(OleInitialize(NULL) == S_OK)
	{
		POINT pt;
		IMenuBand *pIMenuBand;
		MSG msg;
		BOOL bRet;
		LRESULT lresult;

		GetCursorPos(&pt);

		pIMenuBand = PopupMenu(pt.x, pt.y, L"C:\\");
		if(pIMenuBand)
		{
			while((bRet = GetMessage(&msg, NULL, 0, 0)) != 0)
			{
				if(bRet == -1)
					break;

				switch(pIMenuBand->IsMenuMessage(&msg))
				{
				case S_OK:
					pIMenuBand->TranslateMenuMessage(&msg, &lresult);
					break;

				case E_FAIL:
					PostQuitMessage(0);
					break;

				default:
					TranslateMessage(&msg);
					DispatchMessage(&msg);
					break;
				}
			}

			pIMenuBand->Release();
		}
	}

	OleUninitialize();
	ExitProcess(0);
}

IMenuBand *PopupMenu(long nX, long nY, WCHAR *pPath)
{
	HRESULT hr;
	IShellMenu *pIShellMenu;
	IDeskBand *pIDeskBand;
	IMenuBand *pIMenuBand;

	hr = CoCreateInstance(CLSID_MenuBand, NULL, CLSCTX_INPROC_SERVER, IID_IShellMenu, (void**)&pIShellMenu);
	if(SUCCEEDED(hr))
	{
		hr = pIShellMenu->Initialize(NULL, -1, ANCESTORDEFAULT, SMINIT_TOPLEVEL|SMINIT_VERTICAL);

		IShellFolder *pIShellFolderDesktop;

		if(SUCCEEDED(hr))
			hr = SHGetDesktopFolder(&pIShellFolderDesktop);

		if(SUCCEEDED(hr))
		{
			LPITEMIDLIST pidl;

			hr = SHILCreateFromPath(pPath, &pidl, NULL);
			if(SUCCEEDED(hr))
			{
				IShellFolder *pIShellFolder;

				// Get the IShellFolder and PIDL of folder
				hr = pIShellFolderDesktop->BindToObject(pidl, NULL, IID_IShellFolder, (void **)&pIShellFolder);
				if(SUCCEEDED(hr))
				{
					// Folder assignment to menu
					hr = pIShellMenu->SetShellFolder(pIShellFolder, pidl, NULL, SMSET_BOTTOM | SMSET_USEBKICONEXTRACTION);	// I want to also specify: but the value is unknown | SMSET_HASEXPANDABLEFOLDERS)

					IMenuPopup *pIMenuPopup;

					if(SUCCEEDED(hr))
						hr = pIShellMenu->QueryInterface(&pIMenuPopup);

					if(SUCCEEDED(hr))
					{
						IDeskBar *pIDeskBar;

						hr = pIMenuPopup->QueryInterface(&pIDeskBar);
						if(SUCCEEDED(hr))
						{
							hr = pIDeskBar->QueryInterface(&pIDeskBand);
							pIDeskBar->Release();
						}

						pIMenuPopup->Release();
					}

					pIShellFolder->Release();
				}

				ILFree(pidl);
			}

			pIShellFolderDesktop->Release();
		}

		pIShellMenu->Release();
	}

	// Display the pop-up menu IShellMenu the (IDeskBand)
	if(SUCCEEDED(hr))
	{
		IUnknown *pIUnknown;

		// DeskBar for the menu "Menu Desk Bar"
		hr = CoCreateInstance(CLSID_MenuDeskBar, NULL, CLSCTX_INPROC_SERVER, IID_IUnknown, (void**)&pIUnknown);
		if(SUCCEEDED(hr))
		{
			IMenuPopup *pIMenuPopup;

			hr = pIUnknown->QueryInterface(&pIMenuPopup);	// I have a menu for IMenuPopup So DeskBar
			if(SUCCEEDED(hr))
			{
				IBandSite *pIBandSite;

				// BandSite for the menu
				hr = CoCreateInstance(CLSID_MenuBandSite, NULL, CLSCTX_INPROC_SERVER, IID_IBandSite, (void**)&pIBandSite);
				if(SUCCEEDED(hr))
				{
					hr = pIMenuPopup->SetClient(pIBandSite);			// Assign the IBandSite to DeskBar

					if(SUCCEEDED(hr))
						hr = pIBandSite->AddBand(pIDeskBand);			// Assign a IShellMenu you want to display in the IBandSite

					if(SUCCEEDED(hr))
					{
						POINTL ptl;
						RECTL rcl;

						// Menu and display them by a specified position
						ptl.x = nX;
						ptl.y = nY;

						rcl.left = nX;
						rcl.right = nX;
						rcl.top = nY;
						rcl.bottom = nY;

						hr = pIMenuPopup->Popup(&ptl, &rcl, MPPF_SETFOCUS | MPPF_BOTTOM);
						if(SUCCEEDED(hr))
							pIDeskBand->QueryInterface(&pIMenuBand);
					}

					pIBandSite->Release();
				}

				pIMenuPopup->Release();
			}

			pIUnknown->Release();
		}

		pIDeskBand->Release();
	}

	if(SUCCEEDED(hr))
		return pIMenuBand;

	return NULL;
}
