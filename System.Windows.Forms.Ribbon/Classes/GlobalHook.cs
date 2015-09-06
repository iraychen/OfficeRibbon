// *********************************
// Message from Original Author:
//
// 2008 Jose Menendez Poo
// Please give me credit if you use this code. It's all I ask.
// Contact me for more info: menendezpoo@gmail.com
// *********************************
//
// Original project from http://ribbon.codeplex.com/
// Continue to support and maintain by http://officeribbon.codeplex.com/


using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Reflection;
using System.ComponentModel;

namespace System.Windows.Forms.RibbonHelpers
{
    public class GlobalHook
        : IDisposable
    {
        #region Subclasses

        /// <summary>
        /// Types of available hooks
        /// </summary>
        public enum HookTypes
        {
            /// <summary>
            /// Installs a mouse hook
            /// </summary>
            Mouse,

            /// <summary>
            /// Installs a keyboard hook
            /// </summary>
            Keyboard
        }

        #endregion

        #region Fields
        private HookProcCallBack _HookProc;
        private IntPtr _hHook;
        private HookTypes _hookType;
        #endregion

        #region Events

        /// <summary>
        /// Occours when the hook captures a mouse click
        /// </summary>
        public event MouseEventHandler MouseClick;

        /// <summary>
        /// Occours when the hook captures a mouse double click
        /// </summary>
        public event MouseEventHandler MouseDoubleClick;

        /// <summary>
        /// Occours when the hook captures the mouse wheel
        /// </summary>
        public event MouseEventHandler MouseWheel;

        /// <summary>
        /// Occours when the hook captures the press of a mouse button
        /// </summary>
        public event MouseEventHandler MouseDown;

        /// <summary>
        /// Occours when the hook captures the release of a mouse button
        /// </summary>
        public event MouseEventHandler MouseUp;

        /// <summary>
        /// Occours when the hook captures the mouse moving over the screen
        /// </summary>
        public event MouseEventHandler MouseMove;

        /// <summary>
        /// Occours when a key is pressed
        /// </summary>
        public event KeyEventHandler KeyDown;

        /// <summary>
        /// Occours when a key is released
        /// </summary>
        public event KeyEventHandler KeyUp;

        /// <summary>
        /// Occours when a key is pressed
        /// </summary>
        public event KeyPressEventHandler KeyPress;

        #endregion

        #region Delegates

        /// <summary>
        /// Delegate used to recieve HookProc
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        internal delegate IntPtr HookProcCallBack(int nCode, IntPtr wParam, IntPtr lParam);

        #endregion

        #region Ctor

        /// <summary>
        /// Creates a new Hook of the specified type
        /// </summary>
        /// <param name="hookType"></param>
        public GlobalHook(HookTypes hookType)
        {
            _hookType = hookType;
            InstallHook();
        }

        ~GlobalHook()
        {
            // Finalizer calls Dispose(false)
            Dispose(false);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the type of this hook
        /// </summary>
        public HookTypes HookType
        {
            get { return _hookType; }
        }

        /// <summary>
        /// Gets the handle of the hook
        /// </summary>
        public IntPtr Handle
        {
            get { return _hHook; }
        }

        #endregion

        #region Event Triggers

        /// <summary>
        /// Raises the <see cref="MouseClick"/> event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnMouseClick(MouseEventArgs e)
        {
            if (MouseClick != null)
            {
                MouseClick(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="MouseDoubleClick"/> event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnMouseDoubleClick(MouseEventArgs e)
        {
            if (MouseDoubleClick != null)
            {
                MouseDoubleClick(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="MouseWheel"/> event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnMouseWheel(MouseEventArgs e)
        {
            if (MouseWheel != null)
            {
                MouseWheel(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="MouseDown"/> event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnMouseDown(MouseEventArgs e)
        {
            if (MouseDown != null)
            {
                MouseDown(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="MouseUp"/> event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnMouseUp(MouseEventArgs e)
        {
            if (MouseUp != null)
            {
                MouseUp(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="MouseMove"/> event
        /// </summary>
        /// <param name="e"></param>
        protected virtual void OnMouseMove(MouseEventArgs e)
        {
            if (MouseMove != null)
            {
                MouseMove(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="KeyDown"/> event
        /// </summary>
        /// <param name="e">Event Data</param>
        protected virtual void OnKeyDown(KeyEventArgs e)
        {
            if (KeyDown != null)
            {
                KeyDown(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="KeyUp"/> event
        /// </summary>
        /// <param name="e">Event Data</param>
        protected virtual void OnKeyUp(KeyEventArgs e)
        {
            if (KeyUp != null)
            {
                KeyUp(this, e);
            }
        }

        /// <summary>
        /// Raises the <see cref="KeyPress"/> event
        /// </summary>
        /// <param name="e">Event Data</param>
        protected virtual void OnKeyPress(KeyPressEventArgs e)
        {
            if (KeyPress != null)
            {
                KeyPress(this, e);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Recieves the actual unsafe mouse hook procedure
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code < 0)
            {
                return NativeMethods.CallNextHookEx(Handle, code, wParam, lParam);
            }
            else
            {
                switch (HookType)
                {
                    case HookTypes.Mouse:       
                        return MouseProc(code, wParam, lParam);
                    case HookTypes.Keyboard:    
                        return KeyboardProc(code, wParam, lParam);
                    default:
                        throw new Exception("HookType not supported");
                }
            }
        }

        /// <summary>
        /// Recieves the actual unsafe keyboard hook procedure
        /// </summary>
        /// <param name="code"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private IntPtr KeyboardProc(int code, IntPtr wParam, IntPtr lParam)
        {
            NativeMethods.KeyboardLLHookStruct hookStruct = (NativeMethods.KeyboardLLHookStruct)Marshal.PtrToStructure(lParam, typeof(NativeMethods.KeyboardLLHookStruct));

            int msg = wParam.ToInt32();
            bool handled = false;

            if (msg == NativeMethods.WM_KEYDOWN || msg == NativeMethods.WM_SYSKEYDOWN)
            {
                KeyEventArgs e = new KeyEventArgs((Keys)hookStruct.vkCode);
                OnKeyDown(e);
                handled = e.Handled;
            }
            else if (msg == NativeMethods.WM_KEYUP || msg == NativeMethods.WM_SYSKEYUP)
            {
                KeyEventArgs e = new KeyEventArgs((Keys)hookStruct.vkCode);
                OnKeyUp(e);
                handled = e.Handled;
            }

            if (msg == NativeMethods.WM_KEYDOWN && KeyPress != null)
            {
                byte[] keyState = new byte[256];
                byte[] buffer = new byte[2];
                NativeMethods.GetKeyboardState(keyState);
                int conversion = NativeMethods.ToAscii(hookStruct.vkCode, hookStruct.scanCode, keyState, buffer, hookStruct.flags);

                if (conversion == 1 || conversion == 2)
                {
                    bool shift = (NativeMethods.GetKeyState(NativeMethods.VK_SHIFT) & 0x80) == 0x80;
                    bool capital = NativeMethods.GetKeyState(NativeMethods.VK_CAPITAL) != 0;
                    char c = (char)buffer[0];
                    if ((shift ^ capital) && Char.IsLetter(c))
                    {
                        c = Char.ToUpper(c);
                    }
                    KeyPressEventArgs e = new KeyPressEventArgs(c);
                    OnKeyPress(e);
                    handled |= e.Handled;
                }
            }


            return handled ? new IntPtr( 1 ) : NativeMethods.CallNextHookEx(Handle, code, wParam, lParam);
        }

        /// <summary>
        /// Processes Mouse Procedures
        /// </summary>
        /// <param name="code"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private IntPtr MouseProc(int code, IntPtr wParam, IntPtr lParam)
        {
            NativeMethods.MouseLLHookStruct hookStruct = (NativeMethods.MouseLLHookStruct)Marshal.PtrToStructure(lParam, typeof(NativeMethods.MouseLLHookStruct));

            int msg = wParam.ToInt32();
            int x = hookStruct.pt.x;
            int y = hookStruct.pt.y;
            int delta = (short)((hookStruct.mouseData >> 16) & 0xffff);
            
            if (msg == NativeMethods.WM_MOUSEWHEEL)
            {
                OnMouseWheel(new MouseEventArgs(MouseButtons.None, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_MOUSEMOVE)
            {
                OnMouseMove(new MouseEventArgs(MouseButtons.None, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_LBUTTONDBLCLK)
            {
                OnMouseDoubleClick(new MouseEventArgs(MouseButtons.Left, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_LBUTTONDOWN)
            {
                OnMouseDown(new MouseEventArgs(MouseButtons.Left, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_LBUTTONUP)
            {
                OnMouseUp(new MouseEventArgs(MouseButtons.Left, 0, x, y, delta));
                OnMouseClick(new MouseEventArgs(MouseButtons.Left, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_MBUTTONDBLCLK)
            {
                OnMouseDoubleClick(new MouseEventArgs(MouseButtons.Middle, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_MBUTTONDOWN)
            {
                OnMouseDown(new MouseEventArgs(MouseButtons.Middle, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_MBUTTONUP)
            {
                OnMouseUp(new MouseEventArgs(MouseButtons.Middle, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_RBUTTONDBLCLK)
            {
                OnMouseDoubleClick(new MouseEventArgs(MouseButtons.Right, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_RBUTTONDOWN)
            {
                OnMouseDown(new MouseEventArgs(MouseButtons.Right, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_RBUTTONUP)
            {
                OnMouseUp(new MouseEventArgs(MouseButtons.Right, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_XBUTTONDBLCLK)
            {
                OnMouseDoubleClick(new MouseEventArgs(MouseButtons.XButton1, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_XBUTTONDOWN)
            {
                OnMouseDown(new MouseEventArgs(MouseButtons.XButton1, 0, x, y, delta));
            }
            else if (msg == NativeMethods.WM_XBUTTONUP)
            {
                OnMouseUp(new MouseEventArgs(MouseButtons.XButton1, 0, x, y, delta));
            }
            
            return NativeMethods.CallNextHookEx(Handle, code, wParam, lParam);
        }

        /// <summary>
        /// Installs the actual unsafe hook
        /// </summary>
        private void InstallHook()
        {
            /// Error check
            if (Handle != IntPtr.Zero) throw new Exception("Hook is already installed");

            #region htype
            int htype = 0;

            switch (HookType)
            {
                case HookTypes.Mouse:
                    htype = NativeMethods.WH_MOUSE_LL;
                    break;
                case HookTypes.Keyboard:
                    htype = NativeMethods.WH_KEYBOARD_LL;
                    break;
                default:
                    throw new Exception("HookType is not supported");
            } 
            #endregion

            /// Delegate to recieve message
            _HookProc = new HookProcCallBack(HookProc);

            /// Hook
            /// Ed Obeda suggestion for .net 4.0
            //_hHook = NativeMethods.SetWindowsHookEx(htype, _HookProc, Marshal.GetHINSTANCE(Assembly.GetExecutingAssembly().GetModules()[0]), 0);
            _hHook = NativeMethods.SetWindowsHookEx(htype, _HookProc, System.Diagnostics.Process.GetCurrentProcess().MainModule.BaseAddress, 0);
            
            // Error check
            int lastError = Marshal.GetLastWin32Error();
            if (IntPtr.Equals(Handle, IntPtr.Zero)) throw new Win32Exception(lastError);
        }

        /// <summary>
        /// Unhooks the hook
        /// </summary>
        private void Unhook()
        {
            if (Handle != IntPtr.Zero)
            {
                //bool ret = NativeMethods.UnhookWindowsHookEx(Handle);

                //if (ret == false)
                //    throw new Win32Exception(Marshal.GetLastWin32Error());

                //_hHook = 0; 
                try
                {
                    //Fix submitted by Simon Dallmair to handle win32 error when closing the form in vista
                    if (!NativeMethods.UnhookWindowsHookEx(Handle))
                    {
                        Win32Exception ex = new Win32Exception(Marshal.GetLastWin32Error());
                        if (ex.NativeErrorCode != 0)
                            throw ex;
                    }

                    _hHook = IntPtr.Zero;
                }
                catch (Exception)
                {
                }

            }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // free managed resources
            }
            // free native resources
            if (Handle != IntPtr.Zero)
            {
                Unhook();
            }
        }
        #endregion
    }
}
