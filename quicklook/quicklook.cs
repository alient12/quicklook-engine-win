using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections;
using System.IO;
using SHDocVw;
using System.Threading;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Net;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using WebSocketSharp.Server;
using quicklook;

public static class Globals
{
    private static string filenames = "";
    public static bool isClicked = false;
    public static WebSocketServer server;
    public static Dictionary<string, Server> sessions = new Dictionary<string, Server>();

    public static void setFilenames(string filenames)
    {
        if (filenames != "")
            if (Globals.filenames != filenames)
            {
                Globals.filenames = filenames;
                SendToAll(filenames);
            }
    }
    

    public static void SendToAll(string message)
    {
        foreach (var session in sessions.Values)
        {
            session.SendMessage(message);
        }
    }
}


public class FocusWindow
{
    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    static extern int GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    static extern int GetClassName(int hWnd, StringBuilder lpClassName, int nMaxCount);

    public static (string, IntPtr) GetActiveWindowApplicationName()
    {
        IntPtr handle = GetForegroundWindow();
        GetWindowThreadProcessId(handle, out uint processID);
        Process proc = Process.GetProcessById((int)processID);

        const int maxChars = 256;
        StringBuilder className = new StringBuilder(maxChars);
        if (GetClassName((int)handle, className, maxChars) > 0)
        {
            string cName = className.ToString();
            if (cName == "Progman" || cName == "WorkerW")
            {
                // desktop is active
                //return "desktop";
                return ("desktop", handle);
            }
        }
        //return proc.ProcessName;
        return (proc.ProcessName, handle);
    }
}

public class Desktop
{
    public static string CheckDesktop()
    {
        // we basically follow https://devblogs.microsoft.com/oldnewthing/20130318-00/?p=4933
        dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
        var windows = app.Windows;

        const int SWC_DESKTOP = 8;
        const int SWFO_NEEDDISPATCH = 1;
        var hwnd = 0;
        var disp = windows.FindWindowSW(Type.Missing, Type.Missing, SWC_DESKTOP, ref hwnd, SWFO_NEEDDISPATCH);

        var sp = (IServiceProvider)disp;
        var SID_STopLevelBrowser = new Guid("4c96be40-915c-11cf-99d3-00aa004ae837");

        var browser = (IShellBrowser)sp.QueryService(SID_STopLevelBrowser, typeof(IShellBrowser).GUID);
        var view = (IFolderView)browser.QueryActiveShellView();

        ArrayList selected = new ArrayList();
        view.Items(SVGIO.SVGIO_SELECTION, typeof(IShellItemArray).GUID, out var items);
        if (items is IShellItemArray array)
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            for (var i = 0; i < array.GetCount(); i++)
            {
                var item = array.GetItemAt(i);
                string itemName = item.GetDisplayName(SIGDN.SIGDN_NORMALDISPLAY);
                //Console.WriteLine(Path.Combine(desktopPath, itemName));
                selected.Add(Path.Combine(desktopPath, itemName));

            }
        }
        string filenames = string.Join(";", selected.Cast<string>().ToArray());
        return filenames;
    }

    [Guid("6D5140C1-7436-11CE-8034-00AA006009FA"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IServiceProvider
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object QueryService([MarshalAs(UnmanagedType.LPStruct)] Guid service, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);
    }

    // note: for the following interfaces, not all methods are defined as we don't use them here
    [Guid("000214E2-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellBrowser
    {
        void _VtblGap1_12(); // skip 12 methods https://stackoverflow.com/a/47567206/403671

        [return: MarshalAs(UnmanagedType.IUnknown)]
        object QueryActiveShellView();
    }

    [Guid("cde725b0-ccc9-4519-917e-325d72fab4ce"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFolderView
    {
        void _VtblGap1_5(); // skip 5 methods

        [PreserveSig]
        int Items(SVGIO uFlags, Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object items);
    }

    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object BindToHandler(System.Runtime.InteropServices.ComTypes.IBindCtx pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid);

        IShellItem GetParent();

        [return: MarshalAs(UnmanagedType.LPWStr)]
        string GetDisplayName(SIGDN sigdnName);

        // 2 other methods to be defined
    }

    [Guid("b63ea76d-1f85-456f-a19c-48159efa858b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItemArray
    {
        void _VtblGap1_4(); // skip 4 methods

        int GetCount();
        IShellItem GetItemAt(int dwIndex);
    }

    private enum SIGDN
    {
        SIGDN_NORMALDISPLAY,
        SIGDN_PARENTRELATIVEPARSING,
        SIGDN_DESKTOPABSOLUTEPARSING,
        SIGDN_PARENTRELATIVEEDITING,
        SIGDN_DESKTOPABSOLUTEEDITING,
        SIGDN_FILESYSPATH,
        SIGDN_URL,
        SIGDN_PARENTRELATIVEFORADDRESSBAR,
        SIGDN_PARENTRELATIVE,
        SIGDN_PARENTRELATIVEFORUI
    }

    private enum SVGIO
    {
        SVGIO_BACKGROUND,
        SVGIO_SELECTION,
        SVGIO_ALLVIEW,
        SVGIO_CHECKED,
        SVGIO_TYPE_MASK,
        SVGIO_FLAG_VIEWORDER
    }
}

class MyEplorer
{
    static public void GetListOfSelectedFilesAndFolderOfWindowsExplorer()
    {
        (string window_name, IntPtr activeHandle) = FocusWindow.GetActiveWindowApplicationName();
        if (window_name == "desktop")
        {
            Globals.setFilenames(Desktop.CheckDesktop());
        }
        else if (window_name == "explorer")
        {
            string filename;
            ArrayList selected = new ArrayList();
            var shell = new Shell32.Shell();
            //For each explorer
            foreach (SHDocVw.InternetExplorer window in new SHDocVw.ShellWindows())
            {
                filename = Path.GetFileNameWithoutExtension(window.FullName).ToLower();
                if (filename.ToLowerInvariant() == "explorer" && window.HWND == (int)activeHandle)
                {
                    Shell32.FolderItems items = ((Shell32.IShellFolderViewDual2)window.Document).SelectedItems();
                    foreach (Shell32.FolderItem item in items)
                    {
                        //Console.WriteLine(item.Path.ToString());
                        selected.Add(item.Path);
                    }
                }
            }
            Globals.setFilenames(string.Join(";", selected.Cast<string>().ToArray()));
        }
    }
}

class InterceptKeys
{
    private const int WH_KEYBOARD_LL = 13;
    private const int WM_KEYDOWN = 0x0100;
    private static LowLevelKeyboardProc _proc = HookCallback;
    private static IntPtr _hookID = IntPtr.Zero;
    
    private static string keyStr = "Space";
    public static System.Timers.Timer timer = new System.Timers.Timer(250);


    public static void Start()
    {
        timer.AutoReset = true;
        timer.Elapsed += (sender, e) => StartSTAThread(() => MyEplorer.GetListOfSelectedFilesAndFolderOfWindowsExplorer());
        _hookID = SetHook(_proc);
        Application.Run();
        UnhookWindowsHookEx(_hookID);
    }

    private static IntPtr SetHook(LowLevelKeyboardProc proc)
    {
        using (Process curProcess = Process.GetCurrentProcess())
        using (ProcessModule curModule = curProcess.MainModule)
        {
            return SetWindowsHookEx(WH_KEYBOARD_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
        }
    }

    private delegate IntPtr LowLevelKeyboardProc(
        int nCode, IntPtr wParam, IntPtr lParam);

    public static void StartSTAThread(Action action)
    {
        Thread thread = new Thread(() => action());
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        bool finished = thread.Join(100);
        if (!finished)
            thread.Abort();

    }

    private static IntPtr HookCallback(
        int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
        {
            int vkCode = Marshal.ReadInt32(lParam);

            Keys key_space;
            Enum.TryParse(keyStr, out key_space);

            (string window_name, IntPtr activeHandle) = FocusWindow.GetActiveWindowApplicationName();
            if (window_name == "desktop" ||  window_name == "explorer")
                if (!Globals.isClicked && (Keys)vkCode == key_space)
                {
                    Globals.isClicked = true;
                    timer.Start();
                }
                else if (Globals.isClicked &&  (Keys)vkCode == key_space)
                {
                    Globals.isClicked = false;
                    timer.Stop();
                }
        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook,
        LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
        IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
}
