using OpenCvSharp;
using System.Drawing;
using System.Runtime.InteropServices;

namespace SAP스캐너
{
    internal class KeyboardMouseBarcodeHandler
    {
        // WinAPI 선언
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern int SetCursorPos(int x, int y);
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);

        public void ProcessBarcode(IntPtr targetHwnd, string barcode, bool useMouse = false)
        {
            if (targetHwnd == IntPtr.Zero || string.IsNullOrWhiteSpace(barcode)) return;

            SetForegroundWindow(targetHwnd);
            Clipboard.SetText(barcode);
            Thread.Sleep(30);

            var pos = FindInputFieldPosition(useMouse);
            if (pos == null)
            {
                MessageBox.Show("입력 위치 이미지를 찾지 못했습니다.", "알림");
                return;
            }

            SetCursorPos(pos.Value.Item1, pos.Value.Item2);
            mouse_event(0x0002 | 0x0004, 0, 0, 0, 0); // Left Down & Up

            SendKeys.SendWait("{END}");
            SendKeys.SendWait("+{HOME}");
            SendKeys.SendWait("^V");
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(100);
        }

        public (int, int)? FindInputFieldPosition(bool useMouse)
        {
                  // ✨ 1. 검색 영역을 (0,0)부터 (1150, 810)까지의 고정된 사각형으로 설정합니다.
            var searchRect = new System.Drawing.Rectangle(0, 0, 1150, 810);

            //using (Bitmap bmp = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            //{
            //using (Graphics gr = Graphics.FromImage(bmp))
            //{
            //    gr.CopyFromScreen(0, 0, 0, 0, bmp.Size);
            //}


                //using var mat = OpenCvSharp.Extensions.BitmapConverter.ToMat(bmp);

                //string templatePath = useMouse ? "ERPdoc.png" : "Item.png";

                //// ✨ 수정된 부분: ImreadModes.Unchanged 옵션 추가
                //using var tpl = Cv2.ImRead(templatePath, ImreadModes.Unchanged);

                //if (tpl.Empty())
                //{
                //    MessageBox.Show($"템플릿 파일을 찾을 수 없습니다: {templatePath}");
                //    return null;
                //}

                //using var result = new Mat();

                //Cv2.MatchTemplate(mat, tpl, result, TemplateMatchModes.CCoeffNormed);
                //Cv2.MinMaxLoc(result, out _, out double maxval, out _, out OpenCvSharp.Point maxloc);

                //double threshold = 0.85;
                //if (maxval >= threshold)
                //{
                //    int x = maxloc.X + tpl.Width;
                //    int y = maxloc.Y + tpl.Height / 2;
                //    return (x, y);
                //}
                //else
                //{
                //    return null;
                //}

            // ✨ 2. 지정된 영역만 화면을 캡처합니다.
            using (System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(searchRect.Width, searchRect.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            {

                using (System.Drawing.Graphics gr = System.Drawing.Graphics.FromImage(bmp))
                {
                    gr.CopyFromScreen(searchRect.Left, searchRect.Top, 0, 0, bmp.Size);
                }

                using var mat = OpenCvSharp.Extensions.BitmapConverter.ToMat(bmp);
                string templatePath = useMouse ? "ERPdoc.png" : "Item.png";
                using var tpl = OpenCvSharp.Cv2.ImRead(templatePath, OpenCvSharp.ImreadModes.Unchanged);

                if (tpl.Empty())
                {
                    System.Windows.Forms.MessageBox.Show($"템플릿 파일을 찾을 수 없습니다: {templatePath}");
                    return null;
                }

                using var result = new OpenCvSharp.Mat();
                OpenCvSharp.Cv2.MatchTemplate(mat, tpl, result, OpenCvSharp.TemplateMatchModes.CCoeffNormed);
                OpenCvSharp.Cv2.MinMaxLoc(result, out _, out double maxval, out _, out OpenCvSharp.Point maxloc);

                double threshold = 0.85;
                if (maxval >= threshold)
                {
                            // ✨ 3. 찾은 위치를 실제 화면 좌표로 변환합니다.
                            //    (검색 시작점이 0,0 이므로 Left, Top 값을 더할 필요가 없습니다.)
                    int x = maxloc.X + tpl.Width;
                    int y = maxloc.Y + tpl.Height / 2;
                    return (x, y);
                }
                else
                {
                    return null;
                }
            }
        }
    }
}