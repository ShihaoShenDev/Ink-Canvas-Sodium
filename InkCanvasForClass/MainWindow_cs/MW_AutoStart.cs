using Ink_Canvas.Helpers;
using System;
using System.Windows;
using System.Runtime.InteropServices;

namespace Ink_Canvas {
    public partial class MainWindow : PerformanceTransparentWin {
        // IShellLink interface
        [ComImport]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IShellLinkW {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszFile, int cchMaxPath, IntPtr pfd, uint fFlags);
            void GetIDList(out IntPtr ppidl);
            void SetIDList(IntPtr pidl);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszFile, int cchMaxName);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszDir, int cchMaxPath);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszArgs, int cchMaxPath);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            void GetHotkey(out short pwHotkey);
            void SetHotkey(short wHotkey);
            void GetShowCmd(out int piShowCmd);
            void SetShowCmd(int iShowCmd);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder pszIconPath, int cchIconPath, out int piIcon);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
            void Resolve(IntPtr hwnd, uint fFlags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        // IPersistFile interface
        [ComImport]
        [Guid("0000010B-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPersistFile {
            void GetClassID(out Guid pClassID);
            void IsDirty();
            void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
            void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);
            void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
            void GetCurFile([Out, MarshalAs(UnmanagedType.LPWStr)] System.Text.StringBuilder ppszFileName);
        }

        // ShellLink CoClass
        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        [ClassInterface(ClassInterfaceType.None)]
        private class ShellLink { }

        public static bool StartAutomaticallyCreate(string exeName) {
            try {
                var shortcutPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Startup), 
                    exeName + ".lnk");
                
                // 使用 COM 接口创建快捷方式
                var shellLink = new ShellLink();
                var link = (IShellLinkW)shellLink;
                
                //设置快捷方式的目标所在的位置(源程序完整路径)
                link.SetPath(System.Windows.Forms.Application.ExecutablePath);
                //应用程序的工作目录
                //当用户没有指定一个具体的目录时,快捷方式的目标应用程序将使用该属性所指定的目录来装载或保存文件。
                link.SetWorkingDirectory(Environment.CurrentDirectory);
                //目标应用程序窗口类型(1.Normal window普通窗口,3.Maximized最大化窗口,7.Minimized最小化)
                link.SetShowCmd(1); // SW_NORMAL
                //快捷方式的描述
                link.SetDescription(exeName + "_Ink");
                //设置快捷键(如果有必要的话.)
                //link.SetHotkey(0); // 可以设置热键
                
                // 保存快捷方式
                var file = (IPersistFile)link;
                file.Save(shortcutPath, true);
                
                Marshal.ReleaseComObject(file);
                Marshal.ReleaseComObject(link);
                
                return true;
            }
            catch (Exception) { }

            return false;
        }

        public static bool StartAutomaticallyDel(string exeName) {
            try {
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + exeName +
                                      ".lnk");
                return true;
            }
            catch (Exception) { }

            return false;
        }
    }
}