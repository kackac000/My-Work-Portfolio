using OpenCvSharp;
using System.Diagnostics;
//using System.Drawing;
using System.Runtime.InteropServices;
//using System.Windows.Forms;

namespace SAP스캐너
{
    // ✨✨✨
    // ✨ 여기, Form1 클래스 바깥쪽에 아래 코드를 추가해주세요.
    // ✨✨✨
    //public class EnumWindowInfo
    //{
    //    public List<string> TitlesToFind { get; set; } = new List<string>();
    //    public IntPtr FoundHwnd { get; set; } = IntPtr.Zero;
    //}
    public partial class Form1 : Form
    {
        // WinAPI 선언
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        static extern int SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);


        //// Form1.cs 파일에 추가할 WinAPI 및 델리게이트 선언
        //[DllImport("user32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        //[DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        //private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        //[DllImport("user32.dll")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //private static extern bool IsWindowVisible(IntPtr hWnd);

        //// 기존 DllImport 선언들 아래에 추가
        //[DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        //private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        //private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        //// ✨ 1. GetForegroundWindow API 선언 추가
        //[DllImport("user32.dll")]
        //static extern IntPtr GetForegroundWindow();

        // 멤버 변수
        private IntPtr hWndSCAN = IntPtr.Zero;
        private IntPtr targetHwnd = IntPtr.Zero;
        //private bool wasPopupOpen = false;
        private bool isScriptMode = true;
        private ScriptBarcodeHandler scriptHandler = new ScriptBarcodeHandler();
        private KeyboardMouseBarcodeHandler keyMouseHandler = new KeyboardMouseBarcodeHandler();
        private string selectedBranch = "";
        private string Title = "";
        private string TitleIN = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> branches = new List<string> { "3125", "3126", "3127", "3128" };
            string branch = ShowBranchSelectDialog(branches);

            if (string.IsNullOrEmpty(branch))
            {
                Application.Exit();
                return;
            }
            selectedBranch = branch.Trim();
            Title = $"Maintain Outb. Deliv. Order - Warehouse No. {selectedBranch} (Time Zone UTC+9)";
            TitleIN = $"Maintain Inbound Delivery - Warehouse Number {selectedBranch} (Time Zone UTC+9)";

            var bounds = Screen.AllScreens[0].Bounds;
            int screenWidth = bounds.Width;
            int screenHeight = bounds.Height;
            int posX = (int)(screenWidth * 0.7);
            int posY = (int)(screenHeight * 0.2);

            hWndSCAN = this.Handle; // FindWindow보다 this.Handle이 더 정확합니다.
            if (hWndSCAN != IntPtr.Zero)
            {
                ShowWindowAsync(hWndSCAN, 1);
                SetWindowPos(hWndSCAN, 0, posX, posY, 0, 0, 1);
            }

            string processName = Process.GetCurrentProcess().ProcessName;
            if (Process.GetProcessesByName(processName).Length > 1)
            {
                MessageBox.Show("이미 SAP스캐너가 실행 중입니다.");
                Application.Exit();
                return;
            }

            timer1.Interval = 500;
            timer1.Enabled = true;
        }

        private void TextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            // ✨ result 변수를 try 바깥에 선언하여 finally에서도 접근 가능하게 합니다.
            BarcodeParseResult result = null;

            try
            {
                if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
                {
                    if (!string.IsNullOrWhiteSpace(textBox1.Text))
                    {
                        IntPtr hWndOUT = FindWindow("SAP_FRONTEND_SESSION", Title);
                        IntPtr hWndIN = FindWindow(null, TitleIN);

                        if (hWndOUT == IntPtr.Zero && hWndIN == IntPtr.Zero) { MessageBox.Show("검수창이 열려있지 않습니다."); return; }

                        targetHwnd = (hWndOUT != IntPtr.Zero) ? hWndOUT : hWndIN;
                        result = BarcodeParser.Parse(textBox1.Text);

                        if (isScriptMode && scriptHandler.TryProcessBarcode(targetHwnd, result.Value ?? "")) { }
                        else
                        {
                            if (isScriptMode) isScriptMode = false;
                            keyMouseHandler.ProcessBarcode(targetHwnd, result.Value ?? "", result.Type == "OBD");
                            Thread.Sleep(30);
                            TryAutoF11AfterScan(Title, hWndSCAN);
                        }
                    }
                    else
                    {
                        if (targetHwnd != IntPtr.Zero)
                        {
                            SetForegroundWindow(targetHwnd);
                            Thread.Sleep(100);
                            SendKeys.Send("{ENTER}");
                            if (hWndSCAN != IntPtr.Zero)
                            {
                                SetForegroundWindow(hWndSCAN);
                                this.Activate();
                            }
                        }
                    }
                }
                //else if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return) && string.IsNullOrWhiteSpace(textBox1.Text))
                //{
                //    if (targetHwnd != IntPtr.Zero)
                //    {
                //        SetForegroundWindow(targetHwnd);
                //        Thread.Sleep(100);
                //        SendKeys.Send("{ENTER}");
                //        if (hWndSCAN != IntPtr.Zero)
                //        {
                //            SetForegroundWindow(hWndSCAN);
                //            this.Activate();
                //        }
                //    }
                //}
                else if (e.KeyCode == Keys.Escape && string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    if (targetHwnd != IntPtr.Zero)
                    {
                        SetForegroundWindow(targetHwnd);
                        Thread.Sleep(100);
                        SendKeys.Send("{ESC}");
                        if (hWndSCAN != IntPtr.Zero)
                        {
                            SetForegroundWindow(hWndSCAN);
                            this.Activate();
                        }
                    }
                }
                else if (e.KeyCode == Keys.F11 && string.IsNullOrWhiteSpace(textBox1.Text))
                {
                    if (targetHwnd != IntPtr.Zero)
                    {
                        SetForegroundWindow(targetHwnd);
                        Thread.Sleep(150);
                        SendKeys.Send("{F11}");
                        Thread.Sleep(150);
                        SendKeys.Send("{F11}");
                        if (hWndSCAN != IntPtr.Zero)
                        {
                            SetForegroundWindow(hWndSCAN);
                            this.Activate();
                        }
                    }
                }
            }
            finally
            {
                if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Return)
                {
                    textBox1.Text = "";
                    Clipboard.Clear();
                }
                // ✨ 파서가 넘겨준 타입만으로 간단하게 포커스 여부를 결정합니다.
                //    (기존 코드에 흩어져 있던 포커스 반환 로직을 이곳으로 통합하여 정리했습니다.)
                if (result == null || result.Type != "GENERAL_SPECIAL")
                {
                    if (hWndSCAN != IntPtr.Zero)
                    {
                        SetForegroundWindow(hWndSCAN);
                        this.Activate();
                    }
                }
            }
        }

        // 모든 창을 검사하는 콜백 메서드
        //private bool EnumWindowsCallback(IntPtr hWnd, IntPtr lParam)
        //{
        //    var handle = System.Runtime.InteropServices.GCHandle.FromIntPtr(lParam);
        //    var info = handle.Target as EnumWindowInfo;
        //    if (info == null) return false;

        //    if (!IsWindowVisible(hWnd))
        //        return true;

        //    // ✨ 1. 창의 ClassName을 먼저 확인합니다.
        //    var classNameBuilder = new System.Text.StringBuilder(256);
        //    GetClassName(hWnd, classNameBuilder, classNameBuilder.Capacity);
        //    string windowClass = classNameBuilder.ToString();

        //    // ✨ 2. ClassName이 '#32770'(팝업)이 아니면, 바로 건너뜁니다.
        //    if (windowClass != "#32770")
        //        return true;

        //    // ✨ 3. ClassName이 일치하는 경우에만 제목을 확인합니다.
        //    var titleBuilder = new System.Text.StringBuilder(256);
        //    GetWindowText(hWnd, titleBuilder, titleBuilder.Capacity);
        //    string windowTitle = titleBuilder.ToString();

        //    if (info.TitlesToFind.Contains(windowTitle))
        //    {
        //        info.FoundHwnd = hWnd;
        //        return false; // 진짜 팝업을 찾았으므로 중단
        //    }

        //    return true;
        //}

        //// 기존 GetActivePopupHwnd를 대체할 새로운 메서드
        //private IntPtr GetActivePopupHwnd_FindAll()
        //{
        //    targetHwnd = IntPtr.Zero;

        //    // ✨ 여기에 우리가 찾아야 할 모든 팝업창의 제목을 넣습니다.
        //    var popupTitlesToFind = new List<string> { "Batch Maintenance" };

        //    // 팝업창 제목 목록을 콜백 함수로 전달하기 위한 핸들 생성
        //    var gcHandle = System.Runtime.InteropServices.GCHandle.Alloc(popupTitlesToFind);

        //    try
        //    {
        //        EnumWindows(EnumWindowsCallback, System.Runtime.InteropServices.GCHandle.ToIntPtr(gcHandle));
        //    }
        //    finally
        //    {
        //        gcHandle.Free();
        //    }

        //    return targetHwnd;
        //}

        //// ✨ 2. IsPopupActive를 GetActivePopupHwnd로 대체
        //private IntPtr GetActivePopupHwnd()
        //{
        //    IntPtr hWndPopup1 = FindWindow("#32770", Title);                //유효기한 임박 계속하시겠습니까? 확인 창
        //    if (hWndPopup1 != IntPtr.Zero) return hWndPopup1;
        //    IntPtr hWndPopup2 = FindWindow("#32770", "Information");        //검수완료 후 확인 창
        //    if (hWndPopup2 != IntPtr.Zero) return hWndPopup2;
        //    IntPtr hWndPopup3 = FindWindow("#32770", "Display logs");       //다른제품 스캔시 확인 창
        //    if (hWndPopup3 != IntPtr.Zero) return hWndPopup3;
        //    IntPtr hWndPopup4 = FindWindow("#32770", "Batch Maintenance");  //배치 입력 창
        //    if (hWndPopup4 != IntPtr.Zero) return hWndPopup4;
        //    return IntPtr.Zero;
        //}

        //private void timer1_Tick(object sender, EventArgs e)
        //{
        //    IntPtr popupHwnd = GetActivePopupHwnd_FindAll();

        //    // 1. 팝업창이 감지되었을 경우
        //    if (popupHwnd != IntPtr.Zero)
        //    {
        //        // ✨ 수정된 부분: 만약 팝업이 있는데도 포커스가 스캐너에 있다면, 팝업으로 포커스를 강제 이동
        //        IntPtr currentForeground = GetForegroundWindow();
        //        if (currentForeground == this.hWndSCAN)
        //        {
        //            SetForegroundWindow(popupHwnd);
        //        }
        //        wasPopupOpen = true; // 팝업이 열려있음을 기록
        //    }
        //    // 2. 이전에 팝업이 열려있다가 지금은 닫혔을 경우
        //    else if (wasPopupOpen)
        //    {
        //        // 팝업이 닫혔으므로 스캐너 창으로 포커스를 가져옵니다.
        //        if (hWndSCAN != IntPtr.Zero)
        //        {
        //            SetForegroundWindow(hWndSCAN);
        //            this.Activate();
        //        }
        //        wasPopupOpen = false; // 상태를 초기화
        //    }
        //}

        // --- 아래는 수정할 필요 없는 기존 코드들 ---
        private string ShowBranchSelectDialog(List<string> branches)
        {
            using (Form dlg = new Form())
            {
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.Text = "지점 선택";
                dlg.ClientSize = new System.Drawing.Size(140, 160);
                ListBox lb = new ListBox { Left = 20, Top = 20, Width = 100, Height = 90 };
                lb.Items.AddRange(branches.ToArray());
                dlg.Controls.Add(lb);
                Button btn = new Button { Text = "확인", Left = 30, Top = 120, Width = 80, DialogResult = DialogResult.OK };
                dlg.Controls.Add(btn);
                dlg.AcceptButton = btn;
                if (dlg.ShowDialog() == DialogResult.OK && lb.SelectedItem != null)
                {
                    return lb.SelectedItem.ToString();
                }
                return null;
            }
        }

        private bool TryAutoF11AfterScan(string title, IntPtr hWndSCAN, string templatePath = "complete.png", double threshold = 0.85)
        {
            IntPtr hWndSAP = FindWindow("SAP_FRONTEND_SESSION", title);
            if (hWndSAP == IntPtr.Zero) return false;
            Thread.Sleep(150);
            try
            {
                using (Bitmap screenBmp = CaptureFullScreenBitmap())
                using (Bitmap templateBmp = (Bitmap)Image.FromFile(templatePath))
                {
                    double similarity = CalcTemplateMatch(screenBmp, templateBmp);
                    if (similarity >= threshold)
                    {
                        SetForegroundWindow(hWndSAP);
                        Thread.Sleep(100);
                        SendKeys.SendWait("{F11}");
                        Thread.Sleep(100);
                        SendKeys.SendWait("{F11}");
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"자동 F11 처리 중 오류가 발생했습니다:\n{ex.Message}", "오류");
            }
            return false;
        }

        private Bitmap CaptureFullScreenBitmap()
        {
            var bounds = Screen.PrimaryScreen.Bounds;
            var bmp = new Bitmap(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var gr = Graphics.FromImage(bmp))
            {
                gr.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size);
            }
            return bmp;
        }

        private double CalcTemplateMatch(Bitmap screen, Bitmap template)
        {
            using var matScreen = OpenCvSharp.Extensions.BitmapConverter.ToMat(screen);
            using var matTemplate = OpenCvSharp.Extensions.BitmapConverter.ToMat(template);
            if (matScreen.Channels() == 4 && matTemplate.Channels() == 3)
            {
                Cv2.CvtColor(matTemplate, matTemplate, ColorConversionCodes.BGR2BGRA);
            }
            using var result = new Mat();
            Cv2.MatchTemplate(matScreen, matTemplate, result, TemplateMatchModes.CCoeffNormed);
            Cv2.MinMaxLoc(result, out _, out double maxval, out _, out _);
            return maxval;
        }
    }
}
