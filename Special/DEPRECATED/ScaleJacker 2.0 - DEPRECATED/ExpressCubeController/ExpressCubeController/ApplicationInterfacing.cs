using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
namespace ScaleJacker
{
    public class ApplicationInterfacing
    {
        public static bool dimensionOn_fromStarShip = true;
        public static bool scaleOn_fromStarShip = true;

        public static IntPtr scaleCheckBox = IntPtr.Zero;
        public static IntPtr dimensionsCheckBox = IntPtr.Zero;

        public static IntPtr DataCaptureHandle = IntPtr.Zero;
        public static IntPtr ParentHandle = IntPtr.Zero;
      


        public struct WeightDimensions
        {
            public string Length;
            public string Width;
            public string Height;
            public string Weight;
        }

        public struct StarShipHandles
        {
            public IntPtr StarShip_MainHandle;
            public IntPtr StarShip_ParentHandle;
            public IntPtr StarShip_DataCaptureHandle;
            public IntPtr StarShip_LengthHandle;
            public IntPtr StarShip_WidthHandle;
            public IntPtr StarShip_HeightHandle;
            public IntPtr StarShip_WeightHandle;

            public void Update_Starship(WeightDimensions WD,bool isDimensionsOn, bool isScaleOn)
            {
                if (isDimensionsOn == true)
                {
                    NativeMethods.SendMessage(StarShip_LengthHandle, NativeMethods.WM_SETTEXT, null, WD.Length);
                    NativeMethods.SendMessage(StarShip_WidthHandle, NativeMethods.WM_SETTEXT, null, WD.Width);
                    NativeMethods.SendMessage(StarShip_HeightHandle, NativeMethods.WM_SETTEXT, null, WD.Height);                    
                }
                if (isScaleOn == true)
                {
                    NativeMethods.SendMessage(StarShip_WeightHandle, NativeMethods.WM_SETTEXT, null, WD.Weight);
                    //the code below did jack crap... got rid of it. need to find how to "set" the weight in starship.
                    //the line of code below added because starship would take forever to record an actively recorded weight...
                    //so here we shift focus back to the length box...
                    //NativeMethods.SendMessage(StarShip_LengthHandle, NativeMethods.WM_SETFOCUS, null, "");
                }
            }
            
            public void EnumerateHandles()
            {
                IntPtr StarShip_MainHandle = NativeMethods.FindWindow(null, "StarShip Parcel - New Shipment");
                //MessageBox.Show("StarShip HWND: " + int.Parse(StarShip_MainHandle.ToString()).ToString("X"));
                List<IntPtr> ParentWindows = new List<IntPtr>();
                GCHandle pwHandle = GCHandle.Alloc(ParentWindows);
                try
                {
                    NativeMethods.EnumWindowsProc parentProc = new NativeMethods.EnumWindowsProc(EnumWindow);
                    NativeMethods.EnumChildWindows(StarShip_MainHandle, parentProc, GCHandle.ToIntPtr(pwHandle));
                }
                finally
                {
                    if (pwHandle.IsAllocated)
                    {
                        pwHandle.Free();
                    }
                }

                bool hasBeenSet = false;
                for (int y = 0; y < ParentWindows.Count; y++)
                {
                    StringBuilder sb1 = new StringBuilder(256);
                    StringBuilder sb2 = new StringBuilder(256);
                    NativeMethods.GetWindowText(ParentWindows[y], sb1, sb1.Capacity);
                    NativeMethods.GetClassName(ParentWindows[y],sb2,sb2.Capacity);

                    if (sb1.ToString().ToUpper().Contains("PACKAGE") == true)
                    {
                        if (hasBeenSet == false)
                        {
                            StarShip_ParentHandle = ParentWindows[y];
                            ParentHandle = ParentWindows[y];
                        }
                        hasBeenSet = true;
                        // MessageBox.Show(sb1.ToString() + "!!!");                        
                        //break;
                    }
                    else
                    {
                        // MessageBox.Show("Skipping " + sb1.ToString() + " because it is not Package");
                    }
                    if (sb2.ToString().ToUpper().Contains("TVIRTUALSTRINGTREE") == true)
                    {
                        StarShip_DataCaptureHandle = ParentWindows[y];
                       // MessageBox.Show("The handle is " + StarShip_DataCaptureHandle.ToString());
                        DataCaptureHandle = StarShip_DataCaptureHandle;
                       
                    }
                }



                List<IntPtr> ChildWindows = new List<IntPtr>();
                GCHandle cwHandle = GCHandle.Alloc(ChildWindows);
                try
                {
                    NativeMethods.EnumWindowsProc childProc = new NativeMethods.EnumWindowsProc(EnumWindow);
                    NativeMethods.EnumChildWindows(StarShip_ParentHandle, childProc, GCHandle.ToIntPtr(cwHandle));
                }
                finally
                {
                    if (cwHandle.IsAllocated)
                    {
                        cwHandle.Free();
                    }
                }

                // MessageBox.Show("Parent ID: " + int.Parse(StarShip_Handle.ToString()).ToString("X"));
                // MessageBox.Show("Child ID: " + int.Parse(StarShip_TabPage_Handle.ToString()).ToString("X"));

                try
                {
                    StarShip_LengthHandle = ChildWindows[15];
                    StarShip_WidthHandle = ChildWindows[5];
                    StarShip_HeightHandle = ChildWindows[18];
                    StarShip_WeightHandle = ChildWindows[7];
                }
                catch
                {
                    MessageBox.Show("StarShip must be open and the \"Packaging\" tab open to enumerate handles. Call Tom if you cant figure this out.");

                }


                //for (int x = 0; x < ChildWindows.Count; x++)
               // {
                    //MessageBox.Show(int.Parse(ChildWindows[x].ToString()).ToString("X"));
               // }


            }

        }


        public struct SizeIt2Handles
        {
            public IntPtr SizeIt2_MainHandle;
            public IntPtr SizeIt2_ParentHandle;
            public IntPtr SizeIt2_LengthHandle;
            public IntPtr SizeIt2_WidthHandle;
            public IntPtr SizeIt2_HeightHandle;
            public IntPtr SizeIt2_WeightHandle;

            public WeightDimensions GetReading()
            {
                WeightDimensions wd = new WeightDimensions();
                StringBuilder sb_Length = new StringBuilder(256);
                StringBuilder sb_Width = new StringBuilder(256);
                StringBuilder sb_Height = new StringBuilder(256);
                StringBuilder sb_Weight = new StringBuilder(256);

                NativeMethods.GetWindowText(SizeIt2_LengthHandle, sb_Length, sb_Length.MaxCapacity);
                wd.Length = sb_Length.ToString();
                NativeMethods.GetWindowText(SizeIt2_WidthHandle, sb_Width, sb_Width.MaxCapacity);
                wd.Width = sb_Width.ToString();
                NativeMethods.GetWindowText(SizeIt2_HeightHandle, sb_Height, sb_Height.MaxCapacity);
                wd.Height = sb_Height.ToString();
                NativeMethods.GetWindowText(SizeIt2_WeightHandle, sb_Weight, sb_Weight.MaxCapacity);
                wd.Weight = sb_Weight.ToString();

                return wd;

            }

            public void EnumerateHandles()
            {
                IntPtr SizeIt2_MainHandle = NativeMethods.FindWindow(null, "SizeIt 2");
               // MessageBox.Show("Your SizeIt 2 Software is being detected under hwnd #" + int.Parse(SizeIt2_MainHandle.ToString()).ToString("X"));

                List<IntPtr> ParentWindows = new List<IntPtr>();
                GCHandle pwHandle = GCHandle.Alloc(ParentWindows);
                try
                {
                    NativeMethods.EnumWindowsProc parentProc = new NativeMethods.EnumWindowsProc(EnumWindow);
                    NativeMethods.EnumChildWindows(SizeIt2_MainHandle, parentProc, GCHandle.ToIntPtr(pwHandle));
                }
                finally
                {
                    if (pwHandle.IsAllocated)
                    {
                        pwHandle.Free();
                    }
                }


                for (int y = 0; y < ParentWindows.Count; y++)
                {
                    StringBuilder sb1 = new StringBuilder(256);
                    NativeMethods.GetWindowText(ParentWindows[y], sb1, sb1.Capacity);

                    if (sb1.ToString().ToUpper().Contains("MONITOR") == true)
                    {
                        // MessageBox.Show(sb1.ToString() + "!!!");
                        SizeIt2_ParentHandle = ParentWindows[y];
                        break;
                    }
                    else
                    {
                        // MessageBox.Show("Skipping " + sb1.ToString() + " because it is not Package");
                    }
                }

                //MessageBox.Show("The Parent HWND is: " + int.Parse(SizeIt2_ParentHandle.ToString()).ToString("X"));



                List<IntPtr> ChildWindows = new List<IntPtr>();
                GCHandle cwHandle = GCHandle.Alloc(ChildWindows);
                try
                {
                    NativeMethods.EnumWindowsProc childProc = new NativeMethods.EnumWindowsProc(EnumWindow);
                    NativeMethods.EnumChildWindows(SizeIt2_ParentHandle, childProc, GCHandle.ToIntPtr(cwHandle));
                }
                finally
                {
                    if (cwHandle.IsAllocated)
                    {
                        cwHandle.Free();
                    }
                }

                // MessageBox.Show("Parent ID: " + int.Parse(StarShip_Handle.ToString()).ToString("X"));
                // MessageBox.Show("Child ID: " + int.Parse(StarShip_TabPage_Handle.ToString()).ToString("X"));


                SizeIt2_LengthHandle = ChildWindows[9];
                SizeIt2_WidthHandle = ChildWindows[10];
                SizeIt2_HeightHandle = ChildWindows[11];
                SizeIt2_WeightHandle = ChildWindows[8];



                //for (int x = 0; x < ChildWindows.Count; x++)
               // {
                    //MessageBox.Show(int.Parse(ChildWindows[x].ToString()).ToString("X"));
               // }
            }


        }

        public static bool IsProcessRunning(string pFileName)
        {
            Process[] p = Process.GetProcesses();
            

            foreach (Process proc in p)
            {               
                if (proc.MainWindowTitle == pFileName)
                {
                    return true;
                }
            }
            return false;
        }




        public static Panel CreatePanel(IntPtr handle, int bX, int bY, int bHeight, int bWidth)
        {
            var panel = new Panel { Width = bWidth, Height = bHeight, Left = bX, Top = bY };
            ApplicationInterfacing.NativeMethods.SetParent(panel.Handle, handle);
            return panel;

        }
        public static Button CreateButton(IntPtr handle, int bX, int bY, int bHeight, int bWidth, string bText)
        {
            var button = new Button { Text = bText, Left = bX, Top = bY, Width = bWidth, Height = bHeight };

            button.Click += new EventHandler(ClickButton);
            ApplicationInterfacing.NativeMethods.SetParent(button.Handle, handle);
            return button;
        }

        public static CheckBox CreateCheckBox(IntPtr handle, int bX, int bY, int bHeight, int bWidth, string bText)
        {
            var checkbox = new CheckBox { Text = bText, Left = bX, Top = bY, Width = bWidth, Height = bHeight, Checked = true};
            checkbox.BackColor = Color.CornflowerBlue;
            checkbox.Click += new EventHandler(ClickCheckBox);
            
            if (bText == "Scale")
            {
                checkbox.Name = "CheckBoxScale";
                scaleCheckBox =  ApplicationInterfacing.NativeMethods.SetParent(checkbox.Handle, handle);

            }
            if (bText == "Dimensions")
            {
                checkbox.Name = "CheckBoxDimensions";
                dimensionsCheckBox = ApplicationInterfacing.NativeMethods.SetParent(checkbox.Handle, handle);
            }

            return checkbox;
            
        }
        public static void UpdateCheckBox(bool UpdateScale, bool UpdateDimension, IntPtr StarShipParentHandle)
        {
            if (UpdateScale == true)
            {
                scaleCheckBox = ApplicationInterfacing.NativeMethods.SetParent((IntPtr)scaleCheckBox, StarShipParentHandle);
            }
            if (UpdateDimension == true)
            {
                dimensionsCheckBox = ApplicationInterfacing.NativeMethods.SetParent((IntPtr)dimensionsCheckBox, StarShipParentHandle);
            }
        }

        static void checkbox_VisibleChanged(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            //MessageBox.Show("Visibility Changed.");
        }

        static void checkbox_Disposed(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            //MessageBox.Show("Control Disposed");
        }

        static void checkbox_ControlRemoved(object sender, ControlEventArgs e)
        {
            //throw new NotImplementedException();
            //MessageBox.Show("Control Removed");
        }

      
        public static void ClickCheckBox(object sender, EventArgs e)
        {
            if (((CheckBox)sender).Checked == true)
            {

                if (((CheckBox)sender).Text == "Scale")
                {
                    scaleOn_fromStarShip = true;
                    
                }
                if (((CheckBox)sender).Text == "Dimensions")
                {
                    dimensionOn_fromStarShip = true;
                }

            }
            else
            {
                if (((CheckBox)sender).Text == "Scale")
                {
                    scaleOn_fromStarShip = false;
                }
                if (((CheckBox)sender).Text == "Dimensions")
                {
                    dimensionOn_fromStarShip = false;
                }
            }
            
           
        }


        public static void ClickButton(object sender, EventArgs e)
        {
            //run this if created button gets clicked...
            UpdateWeight(DataCaptureHandle,ParentHandle);
            //MessageBox.Show("To confirm the handle, it is: " + DataCaptureHandle.ToString());


        }

        public static void UpdateWeight(IntPtr DataCaptureHandle,IntPtr ParentHandle)
        {
            Rectangle PackagingViewBox = new Rectangle();
            NativeMethods.GetWindowRect(DataCaptureHandle, ref PackagingViewBox);
            
            //MessageBox.Show("x: " + PackagingViewBox.X.ToString() + "    Y: " + PackagingViewBox.Y.ToString());
            // MessageBox.Show("New Pointer: " + newPointer.ToString());
            Point CustomBox_Point = new Point(PackagingViewBox.Left + 30, PackagingViewBox.Top + 28);
            Point ReturnPoint = Cursor.Position;           
            //NativeMethods.SetCursorPos(CustomBox_Point.X, CustomBox_Point.Y);
            Cursor.Hide();
            Cursor.Position = CustomBox_Point;
            NativeMethods.mouse_event(0x0002, 0, 0, 0, 0);
            NativeMethods.mouse_event(0x0004, 0, 0, 0, 0);
            Cursor.Position = ReturnPoint;
            Cursor.Show();

            
            //NativeMethods.mouse_event(NativeMethods.WM_LBUTTONDOWN, CustomBox_Point.X, CustomBox_Point.Y, 0, 0);

           // NativeMethods.mouse_event(NativeMethods.WM_LBUTTONUP, CustomBox_Point.X, CustomBox_Point.Y, 0, 0);
            //NativeMethods.mouse_event(NativeMethods.WM_LBUTTONDOWN, CustomBox_Point.X, CustomBox_Point.Y, 0, 0);
            //NativeMethods.mouse_event(NativeMethods.WM_LBUTTONDOWN, CustomBox_Point.X, CustomBox_Point.Y, 0, 0);
            //System.Threading.Thread.Sleep(200);
            //NativeMethods.mouse_event(NativeMethods.WM_LBUTTONUP, CustomBox_Point.X, CustomBox_Point.Y, 0, 0);
            //NativeMethods.mouse_event(0, 0, 0, 0, 0);
            //NativeMethods.mouse_event(NativeMethods.WM_LBUTTONUP, CustomBox_Point.X, CustomBox_Point.Y, 0, 0);
            //NativeMethods.SetCursorPos(CustomBox_Point.X,CustomBox_Point.Y);
            //System.Threading.Thread.Sleep(400);
            //NativeMethods.SendMessage(newPointer, NativeMethods.WM_LBUTTONDOWN, null, null);
            //System.Threading.Thread.Sleep(200);
            //NativeMethods.SendMessage(newPointer, NativeMethods.WM_LBUTTONUP, null, null);
            //System.Threading.Thread.Sleep(1000);
            //Rectangle parentRect = new Rectangle();
            //NativeMethods.GetWindowRect(ParentHandle, ref parentRect);
            

        }


        public static bool EnumWindow(IntPtr handle, IntPtr pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(pointer);
            List<IntPtr> list = gch.Target as List<IntPtr>;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(handle);
            // You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }










        internal class NativeMethods
        {
            [DllImport("User32.dll", SetLastError = true)]
            public static extern IntPtr FindWindow(string ClassN, string WindN);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

            [DllImport("user32.dll")]
            public static extern int SendMessage(
                  IntPtr hWnd,      // handle to destination window
                  uint Msg,       // message
                  string wParam,  // first message parameter
                  string lParam   // second message parameter
                  );
            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern int GetWindowTextLength(IntPtr hWnd);

            public const int WM_LBUTTONDOWN = 0x0201;
            public const int WM_LBUTTONUP = 0x0202;
            public const int MOUSEEVENTF_LEFTDOWN = 0x02;
            public const int MOUSEEVENTF_LEFTUP = 0x04;
            public const uint BM_CLICK = 0x00F5;
            public const uint WM_SETFOCUS = 0x0007;
            public const uint WM_SETTEXT = 0x000C;

            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool PostMessage(IntPtr hWnd, uint Msg, string wParam, string lParam);

            [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
            public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            [DllImport("user32.dll")]
            public static extern bool GetWindowRect(IntPtr hWnd, ref Rectangle rect);

            [DllImport("user32.dll")]
            public static extern IntPtr WindowFromPoint(Point pnt);
            
            [DllImport("User32.Dll")]
            public static extern long SetCursorPos(int x, int y);

            [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
            public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);

        }





    }
}
