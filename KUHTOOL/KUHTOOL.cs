using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Schema;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Runtime;
using System.Globalization;

namespace KUHTOOL
{
    public partial class KUHTOOL : Form
    {
        static Boolean CloseChk = false;
        public KUHTOOL()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, KeyModifiers fsModifiers, Keys vk);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")]
        public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowDC(IntPtr window);
        [DllImport("user32.dll")]
        public static extern int ReleaseDC(IntPtr window, IntPtr dc);
        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFalgs, int dx, int dy, int dwData, int dwExtraInfo);

        [DllImport("gdi32.dll")]
        public static extern uint GetPixel(IntPtr dc, int x, int y);


        const int HOTKET_ID = 31197;
        const int WM_HOTKEY = 0x0312;
        const int WM_LBUTTON = 0x0021;
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        const int MOUSEEVENTF_LEFTUP = 0x0004;

        int Tick = 0;
        bool time = true;
        int iSearchCount = 0;
        String sSearchText = "";
        
        public enum KeyModifiers
        {
            None = 0,
            Alt = 1,
            Control = 2,
            Shift = 4,
            Windows = 8
        }

        protected override void WndProc(ref Message m)
        {
            switch(m.Msg)
            {
                case WM_HOTKEY :
                    Keys key = (Keys)(((int)m.LParam >> 16) & 0xFFFF);
                    KeyModifiers modifier = (KeyModifiers)((int)m.LParam & 0xFFFF);

                    if (KeyModifiers.Control == modifier && Keys.D1 == key)
                    {
                        this.TopMost = true;
                        this.TopMost = false;
                        this.Show();
                        timer1.Enabled = false;

                        Screen[] Monitor = Screen.AllScreens;
                        if (Monitor.Length >= 1)
                        {
                            Screen screen = Monitor[0];
                            Rectangle workingArea = screen.WorkingArea;
                            this.Location = new Point() {
                                X = Math.Max(workingArea.X, workingArea.X + (workingArea.Width - this.Width) / 2),
                                Y = Math.Max(workingArea.Y, workingArea.Y + (workingArea.Height - this.Height) / 2)
                            };
                        }
                    }

                    if (KeyModifiers.Control == modifier && Keys.D2 == key)
                    {
                        this.TopMost = true;
                        this.TopMost = false;
                        this.Show();
                        timer1.Enabled = false;

                        Screen[] Monitor = Screen.AllScreens;
                        if (Monitor.Length >= 2)
                        {
                            Screen screen = Monitor[1];
                            Rectangle workingArea = screen.WorkingArea;
                            this.Location = new Point() {
                                X = Math.Max(workingArea.X, workingArea.X + (workingArea.Width - this.Width) / 2),
                                Y = Math.Max(workingArea.Y, workingArea.Y + (workingArea.Height - this.Height) / 2)
                            };
                        }
                        else
                        {
                            Screen screen = Monitor[0];
                            Rectangle workingArea = screen.WorkingArea;
                            this.Location = new Point() {
                                X = Math.Max(workingArea.X, workingArea.X + (workingArea.Width - this.Width) / 2),
                                Y = Math.Max(workingArea.Y, workingArea.Y + (workingArea.Height - this.Height) / 2)
                            };
                        }
                    }
                    //else if (KeyModifiers.Control == modifier && Keys.K == key) // 숨김 단축키
                    //{
                        //kuhIcon.Visible = true;
                        //this.Visible = false;
                        //this.ShowInTaskbar = false;
                    //}
                break;
            }
            base.WndProc(ref m);
        }

        private void ultraButton1_Click(object sender, EventArgs e) // KIS 3.0 DEV LOCAL
        {
            try {

                Process.Start("C:\\Program Files (x86)\\Internet Explorer\\iexplore.exe", "http://127.0.0.1:8180/kisframe.jsp");
            }catch (Exception ex){
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton2_Click(object sender, EventArgs e) // KIS 3.0 DEV
        {
            try
            {
                Process.Start("C:\\Program Files (x86)\\Internet Explorer\\iexplore.exe", "http://kisnewdev.kuh.ac.kr/kisframe.jsp");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void kis_exe3_Click(object sender, EventArgs e) // KIS 3.0 OPER LOCAL
        {
            try
            {
                Process.Start("C:\\Program Files (x86)\\Internet Explorer\\iexplore.exe", "http://127.0.0.1:8280/kisframe.jsp");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void kis_exe4_Click(object sender, EventArgs e) // KIS 3.0 OPER
        {
            try
            {
                Process.Start("C:\\Program Files (x86)\\Internet Explorer\\iexplore.exe", "http://kis.kuh.ac.kr/kisframe.jsp");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        

        private void close_Click(object sender, EventArgs e)
        {
            kuhIcon.Visible = true;
            this.Visible = false;
            this.ShowInTaskbar = false;
        }

        private void KUHTOOL_Load(object sender, EventArgs e)
        {
            RegisterHotKey(this.Handle, HOTKET_ID, KeyModifiers.Control, Keys.D1);
            RegisterHotKey(this.Handle, HOTKET_ID, KeyModifiers.Control, Keys.D2);
            this.MaximizeBox = false;

            try
            {
                String runKey = @"SOFTWARE\KUHTOOL";
                RegistryKey strUpKey = Registry.LocalMachine.OpenSubKey(runKey, true);

                if (strUpKey == null)
                {
                    Registry.LocalMachine.CreateSubKey(runKey);
                    strUpKey = Registry.LocalMachine.OpenSubKey(runKey, true);
                }

                if (strUpKey.GetValue("LOGINUSER") != null)
                {
                    this.loginUser.Text = strUpKey.GetValue("LOGINUSER").ToString();
                    strUpKey.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            if (File.Exists(@"C:\HIS\temp\KUHTOOL.txt")) {
                richTextBox1.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL2.txt"))
            {
                richTextBox2.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL2.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL3.txt"))
            {
                richTextBox3.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL3.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL4.txt"))
            {
                richTextBox4.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL4.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL5.txt"))
            {
                richTextBox5.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL5.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL6.txt"))
            {
                richTextBox6.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL6.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL7.txt"))
            {
                richTextBox7.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL7.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL8.txt"))
            {
                richTextBox8.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL8.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL9.txt"))
            {
                richTextBox9.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL9.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL10.txt"))
            {
                richTextBox10.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL10.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL11.txt"))
            {
                richTextBox11.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL11.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL12.txt"))
            {
                richTextBox12.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL12.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL13.txt"))
            {
                richTextBox13.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL13.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL14.txt"))
            {
                richTextBox14.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL14.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\KUHTOOL15.txt"))
            {
                richTextBox15.Text = File.ReadAllText(@"C:\HIS\temp\KUHTOOL15.txt", Encoding.UTF8);
            }

            if (File.Exists(@"C:\HIS\temp\TABName1.txt"))
            {
                tabPage1.Text = File.ReadAllText(@"C:\HIS\temp\TABName1.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName2.txt"))
            {
                tabPage2.Text = File.ReadAllText(@"C:\HIS\temp\TABName2.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName3.txt"))
            {
                tabPage3.Text = File.ReadAllText(@"C:\HIS\temp\TABName3.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName4.txt"))
            {
                tabPage4.Text = File.ReadAllText(@"C:\HIS\temp\TABName4.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName5.txt"))
            {
                tabPage5.Text = File.ReadAllText(@"C:\HIS\temp\TABName5.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName6.txt"))
            {
                tabPage6.Text = File.ReadAllText(@"C:\HIS\temp\TABName6.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName7.txt"))
            {
                tabPage7.Text = File.ReadAllText(@"C:\HIS\temp\TABName7.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName8.txt"))
            {
                tabPage8.Text = File.ReadAllText(@"C:\HIS\temp\TABName8.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName9.txt"))
            {
                tabPage9.Text = File.ReadAllText(@"C:\HIS\temp\TABName9.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName10.txt"))
            {
                tabPage10.Text = File.ReadAllText(@"C:\HIS\temp\TABName10.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName11.txt"))
            {
                tabPage11.Text = File.ReadAllText(@"C:\HIS\temp\TABName11.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName12.txt"))
            {
                tabPage12.Text = File.ReadAllText(@"C:\HIS\temp\TABName12.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName13.txt"))
            {
                tabPage13.Text = File.ReadAllText(@"C:\HIS\temp\TABName13.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName14.txt"))
            {
                tabPage14.Text = File.ReadAllText(@"C:\HIS\temp\TABName14.txt", Encoding.UTF8);
            }
            if (File.Exists(@"C:\HIS\temp\TABName15.txt"))
            {
                tabPage15.Text = File.ReadAllText(@"C:\HIS\temp\TABName15.txt", Encoding.UTF8);
            }

            //발열체크 중지 Mod by jykim 2022.07.21
            //timer2.Enabled = true;

            toolTip1.SetToolTip(checkBox1, "Fixes: 이슈 수정중");
            toolTip1.SetToolTip(checkBox2, "Resolves: 이슈가 해결되었을 때");
            toolTip1.SetToolTip(checkBox3, "Ref: 참고할 이슈가 있을 때");
            toolTip1.SetToolTip(checkBox4, "Related to: 해당 커밋에 관련된 이슈번호 (아직 해결되지 않은 경우)");
            toolTip1.SetToolTip(comboBox2, "feat: 새로운 기능 추가\n" +
                                           "fix: 버그 수정\n" +
                                           "docs: 문서 수정\n" +
                                           "style: 코드 포맷 변경, 세미콜론 누락, 소스코드 내용 변경 없는 경우\n" +
                                           "refactor: 코드 refactoring\n" +
                                           "test: 테스트 코드, refactoring 테스트 코드 추가\n" +
                                           "chore: 빌드 업무 수정, Package Manager 수정, 프로덕션 코드 변경없는 수정\n" +
                                           "etc: 그 외 모든 수정");
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            this.Show();
        }

        private void key_Down_Event(object sender, KeyEventArgs inputkey)
        {
            
        }

        private void SMGwanRi_Click(object sender, EventArgs e) // 협력업체관리 OPEN
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리");
            }
        }

        private void fileserver_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\30 파일서버"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\30 파일서버");
            }
        }

        private void dbchange_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\10 보고서\\KUH-PM-S01-DB 변경요청서"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\10 보고서\\KUH-PM-S01-DB 변경요청서");
            }
        }

        private void dbtuning_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\01.근태신청"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\01.근태신청");
            }
        }

        private void silsuk_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\02.실적평가"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\02.실적평가");
            }
        }

        private void business_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\03.업무실적"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\03.업무실적");
            }
        }

        private void etc_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\07.기타"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\07.기타");
            }
        }

        private void want_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\05.요구사항(한양)"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\05.요구사항(한양)");
            }
        }

        private void DBSafer_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Program Files (x86)\\PNPSECURE\\DBSAFER AGENT\\dsastlch.exe"))
            {
                Process.Start("C:\\Program Files (x86)\\PNPSECURE\\DBSAFER AGENT\\dsastlch.exe");
            }
        }

        private void ultraButton1_Click_1(object sender, EventArgs e) // KIS 최적화
        {
            if (MessageBox.Show("KIS 최적화를 실행 하시겠습니까?", "KIS 최적화", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Process.Start("C:\\HIS\\kuhexp\\InitEnv.cmd");
            }    
        }
        
        private void ezCloudTalk_Click_2(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Users\\KUH\\AppData\\Local\\ezCloudTalk\\bin\\ezCloudTalk.exe"))
            {
                Process.Start("C:\\Users\\KUH\\AppData\\Local\\ezCloudTalk\\bin\\ezCloudTalk.exe");
            }
        }

        private void tf3_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Program Files (x86)\\Comsquare\\TrustForm3\\TFDesigner\\FormDesigner.exe"))
            {
                Process.Start("C:\\Program Files (x86)\\Comsquare\\TrustForm3\\TFDesigner\\FormDesigner.exe");
            }
        }

        private void tf4_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Program Files (x86)\\Comsquare\\TrustForm4\\TFDesigner4.exe"))
            {
                Process.Start("C:\\Program Files (x86)\\Comsquare\\TrustForm4\\TFDesigner4.exe");
            }
        }

        private void tf5_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Program Files (x86)\\Comsquare\\Trustform Designer 5\\TF.Designer.5.exe"))
            {
                Process.Start("C:\\Program Files (x86)\\Comsquare\\Trustform Designer 5\\TF.Designer.5.exe");
            }
        }

        private void recordcopy_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Recordcopy\\KUHIS.Update.exe"))
            {
                Process.Start("C:\\Recordcopy\\KUHIS.Update.exe");
            }
            else 
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\30 파일서버\\표준 및 지침\\개발 환경 설치 자료\\개발환경 설치[최종] for Windows7\\Recordcopy_setup.exe");
            }
        }

        private void addstartup_Click(object sender, EventArgs e)
        {
            try
            {
                //MessageBox.Show(Application.ExecutablePath);
                String runKey = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";
                RegistryKey strUpKey = Registry.LocalMachine.OpenSubKey(runKey,true);
                if (strUpKey.GetValue("KUHTOOL") == null)
                {
                    strUpKey.SetValue("KUHTOOL", Application.ExecutablePath);
                    strUpKey.Close();
                }
                else 
                {
                    strUpKey.DeleteValue("KUHTOOL");
                    strUpKey.SetValue("KUHTOOL", Application.ExecutablePath);
                    strUpKey.Close();
                }

                MessageBox.Show("시작메뉴 등록을 완료했습니다");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void KUHTOOL_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Visible = false;
            kuhIcon.Visible = true;
            if (!CloseChk)             
            {
                e.Cancel = true;
            }
        }

        private void exit_Click(object sender, EventArgs e)
        {
            CloseChk = true;
            kuhIcon.Visible = false;
            UnregisterHotKey(this.Handle, HOTKET_ID);
            this.Close();
        }

        private void show_Click(object sender, EventArgs e)
        {
            this.Show();
        }

        private void KUHTOOL_Deactivate(object sender, EventArgs e)
        {
            this.TopMost = false;
        }

        private void KUHTOOL_Activated(object sender, EventArgs e)
        {
            this.TopMost = false;
        }

        private void rexpert_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Program Files (x86)\\clipsoft\\rexpert30\\bin\\Designer\\Rexpert30Designer.exe"))
            {
                Process.Start("C:\\Program Files (x86)\\clipsoft\\rexpert30\\bin\\Designer\\Rexpert30Designer.exe");
            }
        }

        private void toad_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Quest Software\\Toad for Oracle 11.6\\Toad.exe")) 
            {
                Process.Start("C:\\Quest Software\\Toad for Oracle 11.6\\Toad.exe");
            }
            else if (File.Exists("C:\\Program Files (x86)\\Quest Software\\Toad for Oracle\\toad.exe"))
            {
                Process.Start("C:\\Program Files (x86)\\Quest Software\\Toad for Oracle\\toad.exe");
            }
            else
            {
                MessageBox.Show("Toad 설치 확인 해주세요.");
            }
        }

        private void golden_Click(object sender, EventArgs e)
        {
            if (File.Exists("C:\\Program Files\\Benthic\\Golden6.exe")) 
            {
                Process.Start("C:\\Program Files\\Benthic\\Golden6.exe");
            }
            else if (File.Exists("C:\\Program Files (x86)\\Benthic\\Golden6.exe"))
            {
                Process.Start("C:\\Program Files (x86)\\Benthic\\Golden6.exe");
            }
            else
            {
                MessageBox.Show("Golden 설치 확인 해주세요.");
            }
        }

        private void ultraButton9_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Google\Chrome\Application\chrome.exe", "https://doc.kuh.ac.kr");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton8_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Google\Chrome\Application\chrome.exe", "https://git.kuh.ac.kr");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton7_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Google\Chrome\Application\chrome.exe", "https://ci.kuh.ac.kr");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton6_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files\Google\Chrome\Application\chrome.exe", "https://jira.kuh.ac.kr/");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void KIS_EXE_Click(object sender, EventArgs e)
        {

        }

        private void ultraButton4_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", @"--new-window --app=http://127.0.0.1:8180");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton5_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", @"--new-window --app=http://kisnewdev.kuh.ac.kr");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton3_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", @"--new-window --app=http://127.0.0.1:8280");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton2_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process.Start(@"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe", @"--new-window --app=http://kis.kuh.ac.kr");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton1_Click_2(object sender, EventArgs e)
        {
            try
            {
                String runKey = @"SOFTWARE\KUHTOOL";
                RegistryKey strUpKey = Registry.LocalMachine.OpenSubKey(runKey, true);

                if (strUpKey == null) {
                    Registry.LocalMachine.CreateSubKey(runKey);
                    strUpKey = Registry.LocalMachine.OpenSubKey(runKey, true);
                }

                if (strUpKey.GetValue("LOGINUSER") == null)
                {
                    strUpKey.SetValue("LOGINUSER", this.loginUser.Text);
                    strUpKey.Close();
                }
                else
                {
                    strUpKey.DeleteValue("LOGINUSER");
                    strUpKey.SetValue("LOGINUSER", this.loginUser.Text);
                    strUpKey.Close();
                }

                MessageBox.Show("로그인ID 등록을 완료했습니다");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton11_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.loginUser.Text != "")
                {
                    Process.Start(@"C:\HIS\kuhexp\TFContainerForKIS3.exe", "http://kis.kuh.ac.kr /kis/sysinfo_reweb/xrw/MIT_MERSChecker.xrw " + this.loginUser.Text + " 1▦1520▦911▦50▦50");
                }
                else
                {
                    MessageBox.Show("로그인할 ID를 입력해주세요");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton10_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.loginUser.Text != "")
                {
                    Process.Start(@"C:\HIS\kuhexp\TFContainerForKIS3.exe", "http://kis.kuh.ac.kr /ast/yeongyangsil_reweb/xrw/OAN_DWSchedule.xrw " + this.loginUser.Text + " 1▦1520▦911▦50▦50");
                }
                else
                {
                    MessageBox.Show("로그인할 ID를 입력해주세요");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        public void GetColorAt(Point cursor)
        {
            int x = (int)cursor.X;
            int y = (int)cursor.Y;

            IntPtr desk = GetDesktopWindow();
            IntPtr dc = GetWindowDC(desk);
            int a = (int)GetPixel(dc, x, y);
            ReleaseDC(desk, dc);

            ColorPicker.Color = Color.FromArgb((a >> 0 & 0xff), (a >> 8 & 0xff), (a >> 16 & 0xff));
            Color myColor = Color.FromArgb((a >> 0 & 0xff), (a >> 8 & 0xff), (a >> 16 & 0xff));
            string hex = myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");
            HexColor.Text = "#" + hex;
        }

        private void colorpoint_Click(object sender, EventArgs e)
        {
            if (timer1.Enabled == true)
            {
                timer1.Enabled = false;
            }
            else
            {
                timer1.Enabled = true;
            }            

            //GetColorAt(Cursor.Position);
            //colorDialog1.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Tick++;
            GetColorAt(Cursor.Position);

            //mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            //mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

            if (Tick == 1000)
            {
                timer1.Enabled = false;
            }
        }

        private void ultraButton12_Click(object sender, EventArgs e)
        {
            colorDialog1.ShowDialog();
        }

        private void ultraButton13_Click(object sender, EventArgs e)
        {
            try
            {
                String windir = Environment.GetEnvironmentVariable("windir");
                Process.Start(windir + @"\system32\SnippingTool.exe");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton14_Click(object sender, EventArgs e)
        {

        }

        private void loginUser_TextChanged(object sender, EventArgs e)
        {

        }
        
        private void mainpanel_PaintClient(object sender, PaintEventArgs e)
        {

        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            int H, M;
            H = DateTime.Now.Hour;
            M = DateTime.Now.Minute;
            DateTime nowDt = DateTime.Now;

            if (H == 13 && M == 00 && nowDt.DayOfWeek == DayOfWeek.Wednesday)
            {
                if (time == true)
                {
                    time = false;
                    this.Show();
                    MessageBox.Show("금일 퇴근전까지 주간보고 작성!");

                    try
                    {
                        if (this.loginUser.Text != "")
                        {
                            Process.Start(@"C:\HIS\kuhexp\TFContainerForKIS3.exe", "http://kis.kuh.ac.kr /ocs/ocsservice_reweb/xrw/OSV_SRGwanRi.xrw " + this.loginUser.Text + " 1▦1520▦911▦50▦50");
                        }
                        else
                        {
                            MessageBox.Show("로그인할 ID를 입력해주세요");
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.ToString());
                    }
                }
            }
            else
            {
                time = true;
            }
        }

        private void ultraButton2_Click_2(object sender, EventArgs e)
        {
            try
            {
                if (this.loginUser.Text != "")
                {
                    Process.Start(@"C:\HIS\kuhexp\TFContainerForKIS3.exe", "http://kis.kuh.ac.kr /ocs/ocsservice_reweb/xrw/OSV_SRGwanRi.xrw " + this.loginUser.Text + " 1▦1520▦911▦50▦50");
                }
                else
                {
                    MessageBox.Show("로그인할 ID를 입력해주세요");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton3_Click_1(object sender, EventArgs e)
        {
            try
            {
                String windir = Environment.GetEnvironmentVariable("windir");
                Process.Start(windir + @"\system32\win32calc.exe");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL.txt", richTextBox1.Text, Encoding.UTF8);
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL2.txt", richTextBox2.Text, Encoding.UTF8);
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL3.txt", richTextBox3.Text, Encoding.UTF8);
        }

        private void richTextBox4_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL4.txt", richTextBox4.Text, Encoding.UTF8);
        }

        private void richTextBox5_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL5.txt", richTextBox5.Text, Encoding.UTF8);
        }

        private void richTextBox6_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL6.txt", richTextBox6.Text, Encoding.UTF8);
        }

        private void richTextBox7_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL7.txt", richTextBox7.Text, Encoding.UTF8);
        }

        private void richTextBox8_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL8.txt", richTextBox8.Text, Encoding.UTF8);
        }

        private void richTextBox9_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL9.txt", richTextBox9.Text, Encoding.UTF8);
        }

        private void richTextBox10_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL10.txt", richTextBox10.Text, Encoding.UTF8);
        }

        private void richTextBox11_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL11.txt", richTextBox11.Text, Encoding.UTF8);
        }

        private void richTextBox12_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL12.txt", richTextBox12.Text, Encoding.UTF8);
        }

        private void richTextBox13_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL13.txt", richTextBox13.Text, Encoding.UTF8);
        }

        private void richTextBox14_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL14.txt", richTextBox14.Text, Encoding.UTF8);
        }

        private void richTextBox15_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(@"C:\HIS\temp\KUHTOOL15.txt", richTextBox15.Text, Encoding.UTF8);
        }
        
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            iSearchCount = -1;
            String sFindControl = "richTextBox" + (tabControl1.SelectedIndex + 1).ToString();
            RichTextBox ribox = (RichTextBox)this.Controls.Find(sFindControl, true)[0];

            ribox.Select(0, ribox.Text.Length);
            ribox.SelectionColor = Color.Black;
            ribox.SelectionBackColor = Color.White;
            ribox.Select(0, 0);
            ribox.ScrollToCaret();
        }

        private void ultraButton4_Click_1(object sender, EventArgs e)
        {
            fSearchText();
        }

        private void searchtext_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                fSearchText();
            }
        }   

        private void fSearchText()
        {
            String sFindControl = "richTextBox" + (tabControl1.SelectedIndex + 1).ToString();
            RichTextBox ribox = (RichTextBox)this.Controls.Find(sFindControl, true)[0];

            Regex regex = new Regex(searchtext.Text, RegexOptions.IgnoreCase);
            MatchCollection mc = regex.Matches(ribox.Text);

            ribox.Select(0, ribox.Text.Length);
            ribox.SelectionColor = Color.Black;
            ribox.SelectionBackColor = Color.White;

            int i = 0;
            int iSearchIndex = 0;
            int iCursorPosition = ribox.SelectionStart;

            if (mc.Count - 1 == iSearchCount)
            {
                iSearchCount = 0;
            }
            else if (sSearchText.Equals(searchtext.Text))
            {
                iSearchCount++;
            }
            else
            {
                iSearchCount = 0;
            }

            foreach (Match m in mc)
            {
                int iStartIdx = m.Index;
                int iStopIdx = m.Length;
                                
                if (i == iSearchCount)
                {
                    iSearchIndex = iStartIdx;
                }

                ribox.Select(iStartIdx, iStopIdx);
                ribox.SelectionColor = Color.Black;
                ribox.SelectionBackColor = Color.Yellow;
                ribox.SelectionStart = iCursorPosition;

                i++;
            }

            sSearchText = searchtext.Text;
            ribox.Select(iSearchIndex, 0);
            ribox.ScrollToCaret();
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            String tabIndex = (tabControl1.SelectedIndex + 1).ToString();
            TabNameChange tabch = new TabNameChange(tabIndex);
            tabch.Owner = this;
            tabch.ShowDialog();
        }

        private void HexColor_TextChanged(object sender, EventArgs e)
        {
            string hexColor = HexColor.Text;
            hexColor = hexColor.TrimStart('#');
            Color clr;
            if (hexColor.Length == 6)
            {
                clr = Color.FromArgb(255, // hardcoded opaque
                            int.Parse(hexColor.Substring(0, 2), NumberStyles.HexNumber),
                            int.Parse(hexColor.Substring(2, 2), NumberStyles.HexNumber),
                            int.Parse(hexColor.Substring(4, 2), NumberStyles.HexNumber));
            }
            else // assuming length of 8
            {
                clr = Color.FromArgb(
                            int.Parse(hexColor.Substring(0, 2), NumberStyles.HexNumber),
                            int.Parse(hexColor.Substring(2, 2), NumberStyles.HexNumber),
                            int.Parse(hexColor.Substring(4, 2), NumberStyles.HexNumber),
                            int.Parse(hexColor.Substring(6, 2), NumberStyles.HexNumber));
            }
            ColorPicker.Color = clr;
        }

        private void ultraButton5_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.WorkingDirectory = "C:\\Program Files (x86)\\kumc\\KIS 3.0 App";
                process.StartInfo.FileName = "C:\\Program Files (x86)\\kumc\\KIS 3.0 App\\KISApp.exe";
                process.StartInfo.Arguments = "http://127.0.0.1:8180";
                process.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ultraButton14_Click_1(object sender, EventArgs e)
        {
            try
            {
                Process process = new Process();
                process.StartInfo.WorkingDirectory = "C:\\Program Files (x86)\\kumc\\KIS 3.0 App";
                process.StartInfo.FileName = "C:\\Program Files (x86)\\kumc\\KIS 3.0 App\\KISApp.exe";
                process.StartInfo.Arguments = "http://127.0.0.1:8280";
                process.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        //Commit
        private void ultraButton15_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text.Equals("") || comboBox2.Text.Equals(""))
            {
                MessageBox.Show("업무구분 또는 커밋타입을 선택해주세요");
            }
            else
            {
                String sClipboardText = "[" + comboBox1.Text + "] " + comboBox2.Text + ": \n\n"
                                  + "#상세내용#";

                if (checkBox1.Checked == true || checkBox2.Checked == true || checkBox3.Checked == true || checkBox4.Checked == true)
                {
                    sClipboardText += "\n";
                }

                if (checkBox1.Checked == true)
                {
                    sClipboardText += "\nFixes: #KUMCSR-";
                }
                if (checkBox2.Checked == true)
                {
                    sClipboardText += "\nResolves: #KUMCSR-";
                }
                if (checkBox3.Checked == true)
                {
                    sClipboardText += "\nRef: #KUMCSR-";
                }
                if (checkBox4.Checked == true)
                {
                    sClipboardText += "\nRelated to: #KUMCSR-";
                }
                
                Clipboard.SetText(sClipboardText);
            }
        }

        private void ultraButton16_Click(object sender, EventArgs e)
        {
            if (Directory.Exists("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\00.DB_SCHEMA"))
            {
                Process.Start("\\\\10.20.210.11\\건대병원 프로젝트\\40 협력업체관리\\00.DB_SCHEMA");
            }
        }
    }
}
