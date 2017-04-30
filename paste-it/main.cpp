#include "stdafx.h"
namespace fs = std::experimental::filesystem;

struct Open_clipboard {
  bool const opened;
  explicit Open_clipboard(HWND wnd) : opened{OpenClipboard(wnd) != 0} {
  }

  explicit operator bool() const { return opened; }

  ~Open_clipboard() {
    if (opened) CloseClipboard();
  }
};

template <class T> struct Unique_global_lock {
  HGLOBAL const handle;
  T* const mem;

  explicit Unique_global_lock(HGLOBAL handle) : handle{ handle }, mem { reinterpret_cast<T*>(GlobalLock(handle)) } {}
  operator T*() { return mem; }
  ~Unique_global_lock() { if (mem) GlobalUnlock(handle); }
};


auto disp_to_folder_view(IDispatch& disp) -> std::optional<CComPtr<IFolderView>> {
  CComPtr<IWebBrowser2> browser2;
  if (FAILED(disp.QueryInterface(&browser2))) return {};

  CComPtr<IServiceProvider> svc_provider;
  if (FAILED(disp.QueryInterface(__uuidof(svc_provider.p), reinterpret_cast<void**>(&svc_provider)))) return {};

  CComPtr<IShellBrowser> browser;
  if (FAILED(svc_provider->QueryService(SID_STopLevelBrowser, __uuidof(browser.p), reinterpret_cast<void**>(&browser)))) return {};

  CComPtr<IShellView> sh_view;
  if (FAILED(browser->QueryActiveShellView(&sh_view))) return {};

  CComPtr<IFolderView> folder_view;
  if (FAILED(sh_view.QueryInterface(&folder_view))) return {};

  return { std::move(folder_view) };
}

 auto find_folder_view(IShellWindows& sh_wnds, HWND wnd) -> std::optional<CComPtr<IFolderView>> {
  VARIANT v;
  v.vt = VT_I4;
  for (v.lVal = 0; ; ++v.lVal) {
    CComPtr<IDispatch> disp;
    if (FAILED(sh_wnds.Item(v, &disp)) || !disp) break;

    CComPtr<IWebBrowser2> browser_app;
    if (FAILED(disp.QueryInterface(&browser_app))) continue;

    HWND browser_app_wnd;
    if (FAILED(browser_app->get_HWND(reinterpret_cast<SHANDLE_PTR*>(&browser_app_wnd))) || browser_app_wnd != wnd) continue;

    return disp_to_folder_view(*disp);
  }

  return {};
}

auto folder_view_for_desktop(IShellWindows& sh_wnds) -> std::optional<CComPtr<IFolderView>> {
  CComVariant vt_loc(CSIDL_DESKTOP);
  CComVariant vt_empty;
  long wnd;
  CComPtr<IDispatch> disp;
  sh_wnds.FindWindowSW(&vt_loc, &vt_empty, SWC_DESKTOP, &wnd, SWFO_NEEDDISPATCH, &disp);

  if (disp) return disp_to_folder_view(*disp);
  return {};
}

auto find_fg_folder_view() -> std::optional<CComPtr<IFolderView>> {
  CComPtr<IShellWindows> sh_wnds;
  if (FAILED(sh_wnds.CoCreateInstance(__uuidof(ShellWindows)))) return {};

  auto const fg_wnd = GetForegroundWindow();
  if (auto fg = find_folder_view(*sh_wnds, fg_wnd)) return fg;

  wchar_t fg_cls[128];
  GetClassNameW(fg_wnd, fg_cls, static_cast<int>(std::size(fg_cls)));
  if (auto is_desktop_foreground = wcscmp(fg_cls, L"WorkerW") == 0) return folder_view_for_desktop(*sh_wnds);

  return {};
}

auto id_list_path(ITEMIDLIST& idl) -> std::optional<fs::path> {
  // TODO: handle long path
  wchar_t buffer[MAX_PATH];
  if (FAILED(SHGetPathFromIDListW(&idl, buffer))) return {};

  return { buffer };
}

auto folder_view_path(IFolderView& view) -> std::optional<fs::path> {
  CComPtr<IPersistFolder2> pf2;
  if (FAILED(view.GetFolder(__uuidof(*pf2), reinterpret_cast<void**>(&pf2)))) return {};

  CComHeapPtr<ITEMIDLIST> idl;
  if (FAILED(pf2->GetCurFolder(&idl))) return {};

  return id_list_path(*idl);
}

auto generate_filename() -> std::wstring {
  auto now = std::time(nullptr);
  auto fn = fmt::format("{:%Y-%m-%d-%H-%M-%S}", *std::localtime(&now));
  return std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>>{}.from_bytes(fn);
}

// precondition: clipboard is opened
auto read_clipboard_utf8() -> std::optional<std::string> {
  auto glb = GetClipboardData(CF_UNICODETEXT);
  if (!glb) return {};

  Unique_global_lock<wchar_t> mem(glb);
  return { std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>>{}.to_bytes(mem) };
}

auto paste_it() -> void {
  auto const view = find_fg_folder_view();
  if (!view) return;

  auto const path = folder_view_path(**view);
  if (!path) return;

  auto const file_path = (*path) / (generate_filename() + L".txt");
  {
    auto const text = [&] () -> std::optional<std::string> {
      Open_clipboard const cb{nullptr};
      if (!cb) return {};

      return read_clipboard_utf8();
    }();

    std::fstream stream{ file_path.c_str(), std::ios_base::trunc | std::ios_base::out };
    if (!stream) return;

    stream << *text;
  }

  // wait for the new file to appear in explorer window
  // TODO: find a better way 
  Sleep(1500); 

  for (auto i = 0;; ++i) {
    CComHeapPtr<ITEMID_CHILD> id;
    if (FAILED((*view)->Item(i, &id))) break;
    if (auto p = id_list_path(*id)) {
      if (p->filename() == file_path.filename()) {
        (*view)->SelectItem(i, SVSI_EDIT);
        break;
      }
    }
  }
}

int APIENTRY wWinMain(HINSTANCE, HINSTANCE, wchar_t*, int) {
  RegisterHotKey(nullptr, 24071111, MOD_CONTROL | MOD_SHIFT, VK_INSERT);
  CoInitialize(nullptr);
  

  MSG msg;
  while (GetMessage(&msg, nullptr, 0, 0)) {
    if (msg.message == WM_HOTKEY) {
      paste_it();
    }
  }

  return static_cast<int>(msg.wParam);
}
