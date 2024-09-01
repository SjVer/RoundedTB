using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Threading;
using Newtonsoft.Json;
using Interop.UIAutomationClient;
using System.Runtime.InteropServices;
using Windows.UI.WebUI;
using PInvoke;



namespace RoundedTB
{
    class Taskbar
    {
        private static void InitializeCOMStuffIfNeeded(ref Types.COMStuff comStuff, IntPtr hwnd)
        {
            if (comStuff.Initialized) return;

            int hr;

            // TODO: uninitialized COM when appropriate
            hr = LocalPInvoke.CoInitializeEx(IntPtr.Zero, LocalPInvoke.COINIT_APARTMENTTHREADED);
            if (hr < 0) throw Marshal.GetExceptionForHR(hr);

            // Create the UI Automation object
            hr = LocalPInvoke.CoCreateInstance(
                ref LocalPInvoke.CLSID_CUIAutomation, null, LocalPInvoke.CLSCTX_INPROC_SERVER,
                ref LocalPInvoke.IID_IUIAutomation, out object automationObj);
            if (hr < 0) throw Marshal.GetExceptionForHR(hr);
            comStuff.UIAutomation = (IUIAutomation)automationObj;

            // Create the condition for finding the taskbar root element
            IUIAutomationElement rootElement = comStuff.UIAutomation.GetRootElement();
            IUIAutomationCondition condition = comStuff.UIAutomation.CreatePropertyCondition(
                LocalPInvoke.UIA_NativeWindowHandlePropertyId, hwnd
            );

            // Create the condition for finding the task list
            comStuff.UITaskbarElement = rootElement.FindFirst(TreeScope.TreeScope_Children, condition);
            comStuff.UITaskListCondition = comStuff.UIAutomation.CreatePropertyCondition(
                LocalPInvoke.UIA_ClassNamePropertyId, "Taskbar.TaskListButtonAutomationPeer"
            );

            comStuff.Initialized = true;
        }

        private static LocalPInvoke.RECT GetTaskbarButtonRect(IUIAutomationElement button)
        {
            return new LocalPInvoke.RECT
            {
                Left = button.CurrentBoundingRectangle.left,
                Top = button.CurrentBoundingRectangle.top,
                Right = button.CurrentBoundingRectangle.right,
                Bottom = button.CurrentBoundingRectangle.bottom
            };
        }

        private static LocalPInvoke.RECT GetTaskListButtonsRect(Types.COMStuff comStuff)
        {

            // find all task list buttons
            IUIAutomationElementArray taskbarButtons = comStuff.UITaskbarElement.FindAll(
                TreeScope.TreeScope_Descendants, comStuff.UITaskListCondition
            );

            LocalPInvoke.RECT rect = new LocalPInvoke.RECT
            {
                Left = int.MaxValue,
                Top = int.MaxValue,
                Right = int.MinValue,
                Bottom = int.MinValue
            };
            // get the bounding rectangle of each button and expand the rect to include it
            for (int i = 0; i < taskbarButtons.Length; i++)
            {
                IUIAutomationElement button = taskbarButtons.GetElement(i);
                LocalPInvoke.RECT buttonRect = GetTaskbarButtonRect(button);

                rect.Left = Math.Min(rect.Left, buttonRect.Left);
                rect.Top = Math.Min(rect.Top, buttonRect.Top);
                rect.Right = Math.Max(rect.Right, buttonRect.Right);
                rect.Bottom = Math.Max(rect.Bottom, buttonRect.Bottom);
            }

            return rect;
        }

        private static LocalPInvoke.RECT GetActualTaskbarRect(Types.COMStuff comStuff, IntPtr startHwnd)
        {
            int monitor_offset = 0;

            // The left of secondary monitors is relative to the primary monitor's left
            IntPtr monitor = LocalPInvoke.MonitorFromWindow(startHwnd, 2);
            LocalPInvoke.MONITORINFO monitorInfo = new LocalPInvoke.MONITORINFO() { cbSize = LocalPInvoke.SIZEOF_MONITORINFO };
            if (LocalPInvoke.GetMonitorInfo(monitor, ref monitorInfo))
                monitor_offset = monitorInfo.MonitorRect.Left;

            // Get the rect for the start button
            LocalPInvoke.GetWindowRect(startHwnd, out LocalPInvoke.RECT startRect);

            LocalPInvoke.RECT taskListButtonsRect = GetTaskListButtonsRect(comStuff);

            // Include everything between- and including the start button and the task list
            return new LocalPInvoke.RECT
            {
                Left = startRect.Left - monitor_offset,
                Top = taskListButtonsRect.Top,
                Right = taskListButtonsRect.Right - monitor_offset,
                Bottom = taskListButtonsRect.Bottom
            };
        }



        /// <summary>
        /// Checks if the taskbar is centred.
        /// </summary>
        /// <returns>
        /// A bool indicating if the taskbar is centred.
        /// </returns>
        public static bool CheckIfCentred()
        {
            bool retVal;
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced"))
                {
                    if (key != null)
                    {
                        int val = (int)key.GetValue("TaskbarAl");

                        if (val == 1)
                        {
                            retVal = true;
                        }
                        else
                        {
                            retVal = false;
                        }
                    }
                    else
                    {
                        retVal = false;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
            return retVal;
        }

        /// <summary>
        /// Compares two taskbars' rects to see if they've changed
        /// </summary>
        /// <returns>
        /// a bool indicating if the taskbar's, applist's, and tray's rects rects have changed.
        /// </returns>
        public static bool TaskbarRefreshRequired(Types.Taskbar currentTB, Types.Taskbar newTB, bool isDynamic)
        {
            // REMINDER: newTB will only have rect & hwnd info. Everything else will be null.

            bool taskbarRectChanged = true;
            bool appListRectChanged = true;
            bool trayRectChanged = true;

            if (
                currentTB.TaskbarRect.Left == newTB.TaskbarRect.Left &&
                currentTB.TaskbarRect.Top == newTB.TaskbarRect.Top &&
                currentTB.TaskbarRect.Right == newTB.TaskbarRect.Right &&
                currentTB.TaskbarRect.Bottom == newTB.TaskbarRect.Bottom)
            {
                taskbarRectChanged = false;
            }
            if (
                currentTB.AppListRect.Left == newTB.AppListRect.Left &&
                currentTB.AppListRect.Top == newTB.AppListRect.Top &&
                currentTB.AppListRect.Right == newTB.AppListRect.Right &&
                currentTB.AppListRect.Bottom == newTB.AppListRect.Bottom)
            {
                appListRectChanged = false;
            }
            if (
                currentTB.TrayRect.Left == newTB.TrayRect.Left &&
                currentTB.TrayRect.Top == newTB.TrayRect.Top &&
                currentTB.TrayRect.Right == newTB.TrayRect.Right &&
                currentTB.TrayRect.Bottom == newTB.TrayRect.Bottom)
            {
                trayRectChanged = false;
            }

            if (isDynamic && (taskbarRectChanged || appListRectChanged || trayRectChanged))
            {
                return true;
            }
            else if (!isDynamic && taskbarRectChanged)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the rects of the three components of the taskbar from their respective handles.
        /// </summary>
        /// <returns>
        /// a partial Taskbar containing just rects and handles.
        /// </returns>
        public static Types.Taskbar GetQuickTaskbarRects(Types.Taskbar outdatedTaskbar)
        {
            InitializeCOMStuffIfNeeded(ref outdatedTaskbar.COMStuff, outdatedTaskbar.TaskbarHwnd);
            LocalPInvoke.RECT taskbarRectCheck = GetActualTaskbarRect(outdatedTaskbar.COMStuff, outdatedTaskbar.StartHwnd);

            //LocalPInvoke.GetWindowRect(taskbarHwnd, out LocalPInvoke.RECT taskbarRectCheck);
            LocalPInvoke.GetWindowRect(outdatedTaskbar.TrayHwnd, out LocalPInvoke.RECT trayRectCheck);
            LocalPInvoke.GetWindowRect(outdatedTaskbar.AppListHwnd, out LocalPInvoke.RECT appListRectCheck);

            return new Types.Taskbar()
            {
                TaskbarHwnd = outdatedTaskbar.TaskbarHwnd,
                TrayHwnd = outdatedTaskbar.TrayHwnd,
                AppListHwnd = outdatedTaskbar.AppListHwnd,
                TaskbarRect = taskbarRectCheck,
                TrayRect = trayRectCheck,
                AppListRect = appListRectCheck,
            };
        }

        public static void ResetTaskbar(Types.Taskbar taskbar, Types.Settings settings)
        {
            LocalPInvoke.SetWindowRgn(taskbar.TaskbarHwnd, IntPtr.Zero, true);
            if (settings.CompositionCompat)
            {
                Interaction.UpdateTranslucentTB(taskbar.TaskbarHwnd);
            }
        }

        /// <summary>
        /// Creates a basic region for a specific taskbar and applies it.
        /// </summary>
        /// <returns>
        /// a bool indicating success.
        /// </returns>
        public static bool UpdateSimpleTaskbar(Types.Taskbar taskbar, Types.Settings settings)
        {
            try
            {
                // If independent margins are disabled, set all four margins to the same value
                if (settings.MarginBasic != -384)
                {
                    settings.MarginLeft = settings.MarginBasic;
                    settings.MarginTop = settings.MarginBasic;
                    settings.MarginRight = settings.MarginBasic;
                    settings.MarginBottom = settings.MarginBasic;
                }

                // Create an effective region to be applied to the taskbar
                Types.EffectiveRegion taskbarEffectiveRegion = new Types.EffectiveRegion
                {
                    CornerRadius = Convert.ToInt32(settings.CornerRadius * taskbar.ScaleFactor),
                    Top = Convert.ToInt32(settings.MarginTop * taskbar.ScaleFactor),
                    Left = Convert.ToInt32(settings.MarginLeft * taskbar.ScaleFactor),
                    Width = Convert.ToInt32(taskbar.TaskbarRect.Right - taskbar.TaskbarRect.Left - (settings.MarginRight * taskbar.ScaleFactor)) + 1,
                    Height = Convert.ToInt32(taskbar.TaskbarRect.Bottom - taskbar.TaskbarRect.Top - (settings.MarginBottom * taskbar.ScaleFactor)) + 1
                };

                IntPtr region = LocalPInvoke.CreateRoundRectRgn(taskbarEffectiveRegion.Left, taskbarEffectiveRegion.Top, taskbarEffectiveRegion.Width, taskbarEffectiveRegion.Height, taskbarEffectiveRegion.CornerRadius, taskbarEffectiveRegion.CornerRadius);
                LocalPInvoke.SetWindowRgn(taskbar.TaskbarHwnd, region, true);
                if (settings.CompositionCompat)
                {
                    Interaction.UpdateTranslucentTB(taskbar.TaskbarHwnd);
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Creates a dynamic region for a specific taskbar and applies it.
        /// </summary>
        /// <returns>
        /// a bool indicating success.
        /// </returns>
        public static bool UpdateDynamicTaskbar(Types.Taskbar taskbar, Types.Settings settings)
        {
            try
            {
                // If independent margins are disabled, set all four margins to the same value
                if (settings.MarginBasic != -384)
                {
                    settings.MarginLeft = settings.MarginBasic;
                    settings.MarginTop = settings.MarginBasic;
                    settings.MarginRight = settings.MarginBasic;
                    settings.MarginBottom = settings.MarginBasic;
                }

                // TODO: fix
                settings.MarginLeft = 0;
                settings.MarginTop = 1;
                settings.MarginRight = 0;
                settings.MarginBottom = 0;

                // NOTE: It doesn't seem to agree on what is the top/bottom?
                // Maybe it considers the top of the coordinate system to be the top of the taskbar?

                int left = taskbar.TaskbarRect.Left + Convert.ToInt32(settings.MarginLeft * taskbar.ScaleFactor);
                int right = taskbar.TaskbarRect.Right - Convert.ToInt32(settings.MarginRight * taskbar.ScaleFactor);
                int top = Convert.ToInt32(settings.MarginTop * taskbar.ScaleFactor);
                int height = taskbar.TaskbarRect.Bottom - taskbar.TaskbarRect.Top;
                int bottom = top + height - Convert.ToInt32(settings.MarginBottom * taskbar.ScaleFactor);

                // Idk why
                if (!settings.IsCentred && settings.IsWindows11)
                    left += Convert.ToInt32(12 * taskbar.ScaleFactor);

                // Create an effective region to be applied to the taskbar for the applist
                int cornerRadius = Convert.ToInt32(settings.CornerRadius * taskbar.ScaleFactor);
                IntPtr mainRegion = LocalPInvoke.CreateRoundRectRgn(left, top, right, bottom, cornerRadius, cornerRadius);

                // If the user has it enabled and the tray handle isn't null, create a region for the system tray and merge it with the taskbar region
                if (settings.ShowTray && taskbar.TrayHwnd != IntPtr.Zero)
                {
                    // Create an effective region to be applied to the taskbar for the tray
                    Types.EffectiveRegion trayEffectiveRegion = new Types.EffectiveRegion
                    {
                        CornerRadius = Convert.ToInt32(settings.CornerRadius * taskbar.ScaleFactor),
                        Top = Convert.ToInt32(settings.MarginTop * taskbar.ScaleFactor),
                        Left = Convert.ToInt32(taskbar.ScaleFactor), // Disable custom margin for taskbar left as there's no "padding" provided by Windows and always looks weird as soon as you trim it.
                        Width = Convert.ToInt32(taskbar.TaskbarRect.Right - taskbar.TaskbarRect.Left - (settings.MarginLeft * taskbar.ScaleFactor)) + 1,
                        Height = Convert.ToInt32(taskbar.TaskbarRect.Bottom - taskbar.TaskbarRect.Top - (settings.MarginBottom * taskbar.ScaleFactor)) + 1
                    };

                    IntPtr trayRegion = LocalPInvoke.CreateRoundRectRgn(
                        taskbar.TrayRect.Left - trayEffectiveRegion.Left,
                        trayEffectiveRegion.Top,
                        trayEffectiveRegion.Width,
                        trayEffectiveRegion.Height,
                        trayEffectiveRegion.CornerRadius,
                        trayEffectiveRegion.CornerRadius
                        );

                    IntPtr finalRegion = LocalPInvoke.CreateRoundRectRgn(1, 1, 1, 1, 0, 0);
                    LocalPInvoke.CombineRgn(finalRegion, trayRegion, mainRegion, 2);
                    mainRegion = finalRegion;
                }

                // Apply the final region to the taskbar
                LocalPInvoke.SetWindowRgn(taskbar.TaskbarHwnd, mainRegion, true);
                if (settings.CompositionCompat)
                {
                    Interaction.UpdateTranslucentTB(taskbar.TaskbarHwnd);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        /// <summary>
        /// Checks if there are any new taskbars, or if any taskbars are no longer present.
        /// </summary>
        /// <returns>
        /// a bool indicating success.
        /// </returns>
        public static bool TaskbarCountOrHandleChanged(int taskbarCount, IntPtr mainTaskbarHandle)
        {
            List<IntPtr> currentTaskbars = new List<IntPtr>();
            bool otherTaskbarsExist = true;
            IntPtr hwndPrevious = IntPtr.Zero;
            currentTaskbars.Add(LocalPInvoke.FindWindowExA(IntPtr.Zero, hwndPrevious, "Shell_TrayWnd", null));

            if (currentTaskbars[0] == IntPtr.Zero)
            {
                return false;
            }

            if (currentTaskbars[0] != mainTaskbarHandle)
            {
                return true;
            }

            while (otherTaskbarsExist)
            {
                IntPtr hwndCurrent = LocalPInvoke.FindWindowExA(IntPtr.Zero, hwndPrevious, "Shell_SecondaryTrayWnd", null);
                hwndPrevious = hwndCurrent;

                if (hwndCurrent == IntPtr.Zero)
                {
                    otherTaskbarsExist = false;
                }
                else
                {
                    currentTaskbars.Add(hwndCurrent);
                }
            }
            if (currentTaskbars.Count != taskbarCount)
            {
                return true;
            }
            return false;
        }

        public static bool CheckDynamicUpdateIsValid(Types.Taskbar currentTB, Types.Taskbar newTB)
        {
            // REMINDER: newTB will only have rect & hwnd info. Everything else will be null.

            // Check if either of the supplied taskbars are null
            if (currentTB == null || newTB == null)
            {
                return false;
            }

            // Check if the taskbar handles are different
            if (currentTB.TaskbarHwnd != newTB.TaskbarHwnd)
            {
                return false;
            }

            // Get width of app list. Not strictly necessary as the applist is always measured from the left but doing so just in case
            int newAppListWidth = newTB.AppListRect.Right - newTB.AppListRect.Left;
            int currentAppListWidth = currentTB.AppListRect.Right - currentTB.AppListRect.Left;

            if (newTB.AppListRect.Right >= newTB.TrayRect.Left && newTB.TrayRect.Left != 0)
            {
                return false;
            }

            if (newAppListWidth == newTB.TrayRect.Left && newTB.TrayRect.Left != 0)
            {
                return false;
            }

            if (newAppListWidth <= 20 * currentTB.ScaleFactor && newAppListWidth != 0)
            {
                return false;
            }

            if (newAppListWidth >= newTB.TaskbarRect.Right - newTB.TaskbarRect.Left && newAppListWidth != 0)
            {
                return false;
            }

            Debug.WriteLine($"Old width: {currentAppListWidth}\nNew width: {newAppListWidth}");
            return true;
        }

        /// <summary>
        /// Collects information on any currently-present taskbars.
        /// </summary>
        /// <returns>
        /// A list of taskbars populated with information about their size, handles etc.
        /// </returns>
        public static List<Types.Taskbar> GenerateTaskbarInfo()
        {
            List<Types.Taskbar> retVal = new List<Types.Taskbar>();

            // main window
            {
                IntPtr hwndMain = LocalPInvoke.FindWindowExA(IntPtr.Zero, IntPtr.Zero, "Shell_TrayWnd", null); // Find main taskbar

                Types.COMStuff comStuff = new Types.COMStuff();
                InitializeCOMStuffIfNeeded(ref comStuff, hwndMain);

                IntPtr hwndStart = LocalPInvoke.FindWindowExA(hwndMain, IntPtr.Zero, "Start", null); // Find main start button
                LocalPInvoke.RECT rectMain = GetActualTaskbarRect(comStuff, hwndStart);

                IntPtr hrgnMain = IntPtr.Zero; // Set recovery region to IntPtr.Zero
                IntPtr hwndTray = LocalPInvoke.FindWindowExA(hwndMain, IntPtr.Zero, "TrayNotifyWnd", null); // Get handle to the main taskbar's tray
                LocalPInvoke.GetWindowRect(hwndTray, out LocalPInvoke.RECT rectTray); // Get the RECT for the main taskbar's tray
                IntPtr hwndAppList = LocalPInvoke.FindWindowExA(LocalPInvoke.FindWindowExA(hwndMain, IntPtr.Zero, "ReBarWindow32", null), IntPtr.Zero, "MSTaskSwWClass", null); // Get the handle to the main taskbar's app list
                LocalPInvoke.GetWindowRect(hwndAppList, out LocalPInvoke.RECT rectAppList);// Get the RECT for the main taskbar's app list

                retVal.Add(new Types.Taskbar
                {
                    TaskbarHwnd = hwndMain,
                    TrayHwnd = hwndTray,
                    AppListHwnd = hwndAppList,
                    TaskbarRect = rectMain,
                    TrayRect = rectTray,
                    AppListRect = rectAppList,
                    RecoveryHrgn = hrgnMain,
                    ScaleFactor = Convert.ToDouble(LocalPInvoke.GetDpiForWindow(hwndMain)) / 96.00,
                    TaskbarRes = $"{rectMain.Right - rectMain.Left} x {rectMain.Bottom - rectMain.Top}",
                    Ignored = false,

                    StartHwnd = hwndStart,
                    COMStuff = comStuff,
                });
            }

            // secondary windows
            {
                bool i = true;
                IntPtr hwndPrevious = IntPtr.Zero;
                while (i)
                {
                    IntPtr hwndCurrent = LocalPInvoke.FindWindowExA(IntPtr.Zero, hwndPrevious, "Shell_SecondaryTrayWnd", null);
                    hwndPrevious = hwndCurrent;

                    if (hwndCurrent == IntPtr.Zero)
                    {
                        i = false;
                    }
                    else
                    {
                        Types.COMStuff currentComStuff = new Types.COMStuff();
                        InitializeCOMStuffIfNeeded(ref currentComStuff, hwndCurrent);

                        IntPtr hwndSecStart = LocalPInvoke.FindWindowExA(hwndCurrent, IntPtr.Zero, "Start", null); // Find main start button
                        LocalPInvoke.RECT rectCurrent = GetActualTaskbarRect(currentComStuff, hwndSecStart);

                        LocalPInvoke.GetWindowRgn(hwndCurrent, out IntPtr hrgnCurrent);
                        IntPtr hwndSecTray = LocalPInvoke.FindWindowExA(hwndCurrent, IntPtr.Zero, "TrayNotifyWnd", null); // Get handle to this secondary taskbar's tray
                        LocalPInvoke.GetWindowRect(hwndSecTray, out LocalPInvoke.RECT rectSecTray); // Get the RECT for this secondary taskbar's tray
                        IntPtr hwndSecAppList = LocalPInvoke.FindWindowExA(LocalPInvoke.FindWindowExA(hwndCurrent, IntPtr.Zero, "WorkerW", null), IntPtr.Zero, "MSTaskListWClass", null); // Get the handle to the main taskbar's app list
                        LocalPInvoke.GetWindowRect(hwndSecAppList, out LocalPInvoke.RECT rectSecAppList);// Get the RECT for this secondary taskbar's app list
                        retVal.Add(new Types.Taskbar
                        {
                            TaskbarHwnd = hwndCurrent,
                            TrayHwnd = hwndSecTray,
                            AppListHwnd = hwndSecAppList,
                            TaskbarRect = rectCurrent,
                            TrayRect = rectSecTray,
                            AppListRect = rectSecAppList,
                            RecoveryHrgn = hrgnCurrent,
                            ScaleFactor = Convert.ToDouble(LocalPInvoke.GetDpiForWindow(hwndCurrent)) / 96.00,
                            TaskbarRes = $"{rectCurrent.Right - rectCurrent.Left} x {rectCurrent.Bottom - rectCurrent.Top}",
                            Ignored = false,

                            StartHwnd = hwndSecStart,
                            COMStuff = currentComStuff,
                        });
                    }
                }
            }

            //foreach (var tb in retVal)
            //{
            //    TaskbarShouldBeFilled(tb.TaskbarHwnd);
            //}
            return retVal;
        }

        public static bool TaskbarShouldBeFilled(IntPtr taskbarHwnd, Types.Settings settings)
        {
            if (settings.FillOnMaximise)
            {
                // Attempt to check for if alt+tab/task switcher is open (Windows 11 only)
                IntPtr topHwnd = LocalPInvoke.WindowFromPoint(new LocalPInvoke.POINT() { x = 0, y = 0 });
                StringBuilder windowClass = new StringBuilder(1024);
                try
                {
                    LocalPInvoke.GetClassName(topHwnd, windowClass, 1024);

                    if (windowClass.ToString() == "XamlExplorerHostIslandWindow" && settings.FillOnTaskSwitch)
                    {
                        return true;
                    }
                }
                catch (Exception) { }

                List<IntPtr> windowList = Interaction.GetTopLevelWindows();
                foreach (IntPtr windowHwnd in windowList)
                {
                    if (LocalPInvoke.IsWindowVisible(windowHwnd))
                    {
                        if (LocalPInvoke.MonitorFromWindow(taskbarHwnd, 2) == LocalPInvoke.MonitorFromWindow(windowHwnd, 2))
                        {
                            LocalPInvoke.DwmGetWindowAttribute(windowHwnd, LocalPInvoke.DWMWINDOWATTRIBUTE.Cloaked, out bool isCloaked, 0x4);
                            if (!isCloaked)
                            {
                                LocalPInvoke.WINDOWPLACEMENT lpwndpl = new LocalPInvoke.WINDOWPLACEMENT();
                                LocalPInvoke.GetWindowPlacement(windowHwnd, ref lpwndpl);
                                if (lpwndpl.ShowCmd == LocalPInvoke.ShowWindowCommands.ShowMaximized)
                                    return true;
                            }
                        }
                    }
                }
            }

            return false;
        }
    }
}
