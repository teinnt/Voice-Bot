using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Voice_Bot
{
    class Peripherals
    {
        public class Mouse
        {
            //Declare mouse events
            private const int MOUSEEVENTF_LEFTDOWN = 0x02;
            private const int MOUSEEVENTF_LEFTUP = 0x04;
            private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
            private const int MOUSEEVENTF_RIGHTUP = 0x10;

            public static void DoMouseClick()
            {
                Win32.mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            }

            public static void MoveToPoint(int X, int Y)
            {
                Win32.SetCursorPos(X, Y);
            }

            public static int X { get => Cursor.Position.X; }
            public static int Y { get => Cursor.Position.Y; }
        }

        private class Win32
        {
            //Setting location
            [DllImport("User32.Dll")]
            public static extern long SetCursorPos(int x, int y);

            [DllImport("User32.Dll")]
            public static extern bool ClientToScreen(IntPtr hWnd, ref POINT point);

            [StructLayout(LayoutKind.Sequential)]
            public struct POINT
            {
                public int x;
                public int y;
            }

            //Mouse click
            [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
            public static extern void mouse_event(uint dwFlags, int dx, int dy, uint cButtons, int dwExtraInfo);
        }
    }
}
