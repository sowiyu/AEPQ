using AEPQ.Services;
using EasyModbus;
using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AEPQ
{
    public partial class MainForm : MaterialForm
    {
        private readonly MaterialSkinManager materialSkinManager;

        // --- 통신 서비스 ---
        private Rs485Service rs485Service;
        private ModbusService modbusService;
        private TcpService tcpService;

        // --- UI 컨트롤 ---
        private MaterialTabControl tabControl;
        private TabPage tabPageStart, tabPageManual;

        // '시작' 탭 컨트롤
        private GroupBox groupModbus, groupTcp, groupRs485Start;
        private TextBox txtModbusIp, txtModbusPort;
        private Button btnModbusConnect, btnModbusDisconnect;
        private TextBox txtTcpIp, txtTcpPort;
        private Button btnTcpConnect, btnTcpDisconnect;
        private ComboBox cmbPortStart;
        private Button btnRs485ConnectStart, btnRs485DisconnectStart, btnRs485RefreshStart;
        private Button btnStartOperation1, btnStartOperation2;
        private RichTextBox txtLog;

        // '단동 테스트' 탭 컨트롤
        private GroupBox groupRS485Manual;
        private ComboBox cmbPortManual;
        private Button btnConnectManual, btnDisconnectManual, btnRefreshManual;
        private DataGridView dgvCommands;
        private GroupBox autoModeGroup, sendGroup, simulationGroup;
        private Button btnStartAutoMode, btnStopAutoMode;
        private TextBox txtCustomPacket, txtSimulatedReceive;
        private Button btnSendCustomPacket, btnProcessSimulatedReceive;

        public MainForm()
        {
            InitializeComponent();

            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.BlueGrey800, Primary.BlueGrey900,
                Primary.BlueGrey500, Accent.LightBlue200,
                TextShade.WHITE);
        }

        private void InitializeComponent()
        {
            this.Text = "AEPQ 통합 제어 프로그램";
            this.ClientSize = new System.Drawing.Size(940, 800);

            tabControl = new MaterialTabControl { Dock = DockStyle.Fill };
            tabPageStart = new TabPage { Text = "시작" };
            tabPageManual = new TabPage { Text = "단동 테스트" };
            tabControl.TabPages.Add(tabPageStart);
            tabControl.TabPages.Add(tabPageManual); // <-- 빠뜨렸던 코드 추가!

            txtLog = new RichTextBox
            {
                Dock = DockStyle.Bottom,
                Height = 250,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = Color.White
            };

            this.Controls.Add(tabControl);
            this.Controls.Add(txtLog);

            InitializeStartTab();
            InitializeManualTestTab();
        }

        #region UI 초기화
        private void InitializeStartTab()
        {
            groupRs485Start = new GroupBox { Text = "RS-485 연결", Location = new Point(20, 20), Width = 300, Height = 120 };
            groupRs485Start.Controls.Add(new Label { Text = "COM 포트:", Location = new Point(15, 35), AutoSize = true });
            cmbPortStart = new ComboBox { Location = new Point(95, 32), Width = 120 };
            btnRs485RefreshStart = new Button { Text = "새로고침", Location = new Point(225, 30), Width = 60 };
            btnRs485RefreshStart.Click += (s, e) => LoadPorts();
            btnRs485ConnectStart = new Button { Text = "연결", Location = new Point(15, 70), Width = 130 };
            btnRs485ConnectStart.Click += BtnRs485Connect_Click;
            btnRs485DisconnectStart = new Button { Text = "해제", Location = new Point(155, 70), Width = 130, Enabled = false };
            btnRs485DisconnectStart.Click += BtnRs485Disconnect_Click;
            groupRs485Start.Controls.Add(cmbPortStart);
            groupRs485Start.Controls.Add(btnRs485RefreshStart);
            groupRs485Start.Controls.Add(btnRs485ConnectStart);
            groupRs485Start.Controls.Add(btnRs485DisconnectStart);
            tabPageStart.Controls.Add(groupRs485Start);

            groupModbus = new GroupBox { Text = "Modbus TCP/IP 연결", Location = new Point(330, 20), Width = 300, Height = 120 };
            groupModbus.Controls.Add(new Label { Text = "IP 주소:", Location = new Point(15, 35), AutoSize = true });
            txtModbusIp = new TextBox { Text = "127.0.0.1", Location = new Point(80, 32), Width = 120 };
            groupModbus.Controls.Add(new Label { Text = "포트:", Location = new Point(210, 35), AutoSize = true });
            txtModbusPort = new TextBox { Text = "502", Location = new Point(250, 32), Width = 40 };
            btnModbusConnect = new Button { Text = "연결", Location = new Point(15, 70), Width = 130 };
            btnModbusConnect.Click += BtnModbusConnect_Click;
            btnModbusDisconnect = new Button { Text = "해제", Location = new Point(155, 70), Width = 130, Enabled = false };
            btnModbusDisconnect.Click += BtnModbusDisconnect_Click;
            groupModbus.Controls.Add(txtModbusIp);
            groupModbus.Controls.Add(txtModbusPort);
            groupModbus.Controls.Add(btnModbusConnect);
            groupModbus.Controls.Add(btnModbusDisconnect);
            tabPageStart.Controls.Add(groupModbus);

            groupTcp = new GroupBox { Text = "TCP/IP 연결 (좌표)", Location = new Point(640, 20), Width = 280, Height = 120 };
            groupTcp.Controls.Add(new Label { Text = "IP:Port", Location = new Point(15, 35), AutoSize = true });
            txtTcpIp = new TextBox { Text = "127.0.0.1", Location = new Point(65, 32), Width = 120 };
            txtTcpPort = new TextBox { Text = "8080", Location = new Point(195, 32), Width = 60 };
            btnTcpConnect = new Button { Text = "연결", Location = new Point(15, 70), Width = 120 };
            btnTcpConnect.Click += BtnTcpConnect_Click;
            btnTcpDisconnect = new Button { Text = "해제", Location = new Point(145, 70), Width = 120, Enabled = false };
            btnTcpDisconnect.Click += BtnTcpDisconnect_Click;
            groupTcp.Controls.Add(txtTcpIp);
            groupTcp.Controls.Add(txtTcpPort);
            groupTcp.Controls.Add(btnTcpConnect);
            groupTcp.Controls.Add(btnTcpDisconnect);
            tabPageStart.Controls.Add(groupTcp);

            btnStartOperation1 = new Button { Text = "동작 1 시작", Location = new Point(50, 160), Width = 400, Height = 300, Font = new Font(this.Font.FontFamily, 24, FontStyle.Bold) };
            btnStartOperation1.Click += BtnStartOperation1_Click;
            btnStartOperation2 = new Button { Text = "동작 2 시작", Location = new Point(480, 160), Width = 400, Height = 300, Font = new Font(this.Font.FontFamily, 24, FontStyle.Bold) };
            btnStartOperation2.Click += BtnStartOperation2_Click;
            tabPageStart.Controls.Add(btnStartOperation1);
            tabPageStart.Controls.Add(btnStartOperation2);
        }

        private void InitializeManualTestTab()
        {
            groupRS485Manual = new GroupBox { Text = "RS-485 연결", Location = new Point(20, 20), Width = 880, Height = 60 };
            groupRS485Manual.Controls.Add(new Label { Text = "COM 포트:", Location = new Point(15, 25), AutoSize = true });
            cmbPortManual = new ComboBox { Location = new Point(95, 22), Width = 120 };
            btnRefreshManual = new Button { Text = "새로고침", Location = new Point(225, 20), Width = 80 };
            btnRefreshManual.Click += (s, e) => LoadPorts();
            btnConnectManual = new Button { Text = "연결", Location = new Point(315, 20), Width = 100 };
            btnConnectManual.Click += BtnRs485Connect_Click;
            btnDisconnectManual = new Button { Text = "해제", Location = new Point(425, 20), Width = 100, Enabled = false };
            btnDisconnectManual.Click += BtnRs485Disconnect_Click;
            groupRS485Manual.Controls.Add(cmbPortManual);
            groupRS485Manual.Controls.Add(btnRefreshManual);
            groupRS485Manual.Controls.Add(btnConnectManual);
            groupRS485Manual.Controls.Add(btnDisconnectManual);
            tabPageManual.Controls.Add(groupRS485Manual);

            dgvCommands = new DataGridView
            {
                Location = new Point(20, 90),
                Width = 880,
                Height = 250,
                AllowUserToAddRows = false,
                ReadOnly = true,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            };
            dgvCommands.CellContentClick += DgvCommands_CellContentClick;
            tabPageManual.Controls.Add(dgvCommands);

            autoModeGroup = new GroupBox { Text = "자동 모드 (핸드툴 1 & 2)", Location = new Point(20, 350), Width = 880, Height = 55 };
            btnStartAutoMode = new Button { Text = "시작", Location = new Point(15, 20), Width = 150 };
            btnStartAutoMode.Click += BtnStartAutoMode_Click;
            btnStopAutoMode = new Button { Text = "중지", Location = new Point(180, 20), Width = 150, Enabled = false };
            btnStopAutoMode.Click += BtnStopAutoMode_Click;
            autoModeGroup.Controls.Add(btnStartAutoMode);
            autoModeGroup.Controls.Add(btnStopAutoMode);
            tabPageManual.Controls.Add(autoModeGroup);

            sendGroup = new GroupBox { Text = "패킷 직접 전송", Location = new Point(20, 410), Width = 880, Height = 75 };
            sendGroup.Controls.Add(new Label { Text = "16-byte hex:", Location = new Point(15, 35), AutoSize = true });
            txtCustomPacket = new TextBox { Location = new Point(100, 32), Width = 650, Font = new Font("Consolas", 9) };
            btnSendCustomPacket = new Button { Text = "전송", Location = new Point(760, 30), Width = 100 };
            btnSendCustomPacket.Click += BtnSendCustomPacket_Click;
            sendGroup.Controls.Add(txtCustomPacket);
            sendGroup.Controls.Add(btnSendCustomPacket);
            tabPageManual.Controls.Add(sendGroup);

            simulationGroup = new GroupBox { Text = "수신 시뮬레이션", Location = new Point(20, 490), Width = 880, Height = 75 };
            simulationGroup.Controls.Add(new Label { Text = "16-byte hex:", Location = new Point(15, 35), AutoSize = true });
            txtSimulatedReceive = new TextBox { Location = new Point(100, 32), Width = 650, Font = new Font("Consolas", 9) };
            btnProcessSimulatedReceive = new Button { Text = "수신 처리", Location = new Point(760, 30), Width = 100 };
            btnProcessSimulatedReceive.Click += BtnProcessSimulatedReceive_Click;
            simulationGroup.Controls.Add(txtSimulatedReceive);
            simulationGroup.Controls.Add(btnProcessSimulatedReceive);
            tabPageManual.Controls.Add(simulationGroup);

            LoadPorts();
            InitializeCommandGrid();
        }
        #endregion

        #region 시작 탭 이벤트
        private void BtnModbusConnect_Click(object sender, EventArgs e)
        {
            modbusService = new ModbusService(txtModbusIp.Text, int.Parse(txtModbusPort.Text), (msg, color) => Log(msg, color));
            if (modbusService.Connect())
            {
                btnModbusConnect.Enabled = false;
                btnModbusDisconnect.Enabled = true;
            }
        }

        private void BtnModbusDisconnect_Click(object sender, EventArgs e)
        {
            modbusService?.Disconnect();
            btnModbusConnect.Enabled = true;
            btnModbusDisconnect.Enabled = false;
        }

        private async void BtnTcpConnect_Click(object sender, EventArgs e)
        {
            tcpService = new TcpService((msg, color) => Log(msg, color));
            bool success = await tcpService.ConnectAsync(txtTcpIp.Text, int.Parse(txtTcpPort.Text));
            if (success)
            {
                btnTcpConnect.Enabled = false;
                btnTcpDisconnect.Enabled = true;
            }
        }

        private void BtnTcpDisconnect_Click(object sender, EventArgs e)
        {
            tcpService?.Disconnect();
            btnTcpConnect.Enabled = true;
            btnTcpDisconnect.Enabled = false;
        }

        private async void BtnStartOperation1_Click(object sender, EventArgs e)
        {
            Log("▶️ 동작 1 시작...", Color.Blue);

            if (rs485Service == null || !rs485Service.IsOpen) { Log("  - ⚠ RS-485 포트가 연결되지 않았습니다.", Color.Orange); return; }
            if (tcpService == null || !tcpService.IsConnected) { Log("  - ⚠ TCP/IP (좌표)가 연결되지 않았습니다.", Color.Orange); return; }
            if (modbusService == null || !modbusService.IsConnected) { Log("  - ⚠ Modbus TCP/IP가 연결되지 않았습니다.", Color.Orange); return; }

            var cmd = new CommandData { Description = "케이스 실린더 전진", Data = new byte[] { 0, 0, 0x20, 0, 0, 0, 0, 0, 0 } };
            rs485Service.SendPacket(cmd);
            await Task.Delay(500);

            string coordinate = await tcpService.RequestCoordinate();
            if (coordinate == "ERROR") return;

            bool success = modbusService.WriteRegister(130, 1);
            if (!success) return;
            await Task.Delay(200);

            int[] readValue = modbusService.ReadHoldingRegisters(150, 1);
            if (readValue != null && readValue.Length > 0)
            {
                Log($"  - Modbus: 주소 150에서 값 '{readValue[0]}' 읽기 완료", Color.Blue);
            }

            Log("✅ 동작 1의 초기 단계가 완료되었습니다.", Color.Green);
        }

        private void BtnStartOperation2_Click(object sender, EventArgs e)
        {
            Log("▶️ 동작 2 시작...", Color.DarkCyan);
        }
        #endregion

        #region 단동 테스트 탭 이벤트
        private void InitializeCommandGrid()
        {
            var commands = new System.Collections.Generic.List<CommandData>
            {
                new CommandData { Description = "🚨 전체 정지", Data = new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 1 진공", Data = new byte[] { 0, 0x01, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 1 파기", Data = new byte[] { 0, 0x02, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 2 진공", Data = new byte[] { 0, 0x04, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 2 파기", Data = new byte[] { 0, 0x08, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3 진공", Data = new byte[] { 0, 0x10, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3 파기", Data = new byte[] { 0, 0x20, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 4 진공", Data = new byte[] { 0, 0x40, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 4 파기", Data = new byte[] { 0, 0x80, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3+4 동시 진공", Data = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 3+4 동시 파기", Data = new byte[] { 0, 0xA0, 0, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5 진공", Data = new byte[] { 0, 0, 0x01, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5 파기", Data = new byte[] { 0, 0, 0x02, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 6 진공", Data = new byte[] { 0, 0, 0x04, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 6 파기", Data = new byte[] { 0, 0, 0x08, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5+6 동시 진공", Data = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "이젝터 5+6 동시 파기", Data = new byte[] { 0, 0, 0x0A, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 1 후진", Data = new byte[] { 0, 0, 0x10, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 1 전진", Data = new byte[] { 0, 0, 0x20, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 2 후진", Data = new byte[] { 0, 0, 0x40, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 2 전진", Data = new byte[] { 0, 0, 0x80, 0, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 3 후진", Data = new byte[] { 0, 0, 0, 0x01, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 3 전진", Data = new byte[] { 0, 0, 0, 0x02, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 4 후진", Data = new byte[] { 0, 0, 0, 0x04, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 4 전진", Data = new byte[] { 0, 0, 0, 0x08, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 5 UP",   Data = new byte[] { 0, 0, 0, 0x10, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 5 DOWN", Data = new byte[] { 0, 0, 0, 0x20, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 6 후진", Data = new byte[] { 0, 0, 0, 0x40, 0, 0, 0, 0, 0 } },
                new CommandData { Description = "실린더 6 전진", Data = new byte[] { 0, 0, 0, 0x80, 0, 0, 0, 0, 0 } },
            };
            dgvCommands.DataSource = commands;
            dgvCommands.Columns["Description"].HeaderText = "명령 설명";
            dgvCommands.Columns["Data"].Visible = false;
            var sendButtonColumn = new DataGridViewButtonColumn { Name = "SendButton", HeaderText = "동작", Text = "전송", UseColumnTextForButtonValue = true };
            dgvCommands.Columns.Add(sendButtonColumn);
            dgvCommands.Columns["SendButton"].Width = 80;
        }

        private void LoadPorts()
        {
            string[] portNames = System.IO.Ports.SerialPort.GetPortNames();
            cmbPortStart.Items.Clear();
            cmbPortManual.Items.Clear();
            cmbPortStart.Items.AddRange(portNames);
            cmbPortManual.Items.AddRange(portNames);

            if (portNames.Length > 0)
            {
                cmbPortStart.SelectedIndex = 0;
                cmbPortManual.SelectedIndex = 0;
            }
            else
            {
                cmbPortStart.Text = "포트 없음";
                cmbPortManual.Text = "포트 없음";
            }
        }

        private void BtnRs485Connect_Click(object sender, EventArgs e)
        {
            bool isStartTab = sender == btnRs485ConnectStart;
            ComboBox currentCmb = isStartTab ? cmbPortStart : cmbPortManual;

            if (currentCmb.SelectedItem == null) { Log("COM 포트를 선택하세요.", Color.Red); return; }

            rs485Service?.Disconnect();

            rs485Service = new Rs485Service((msg, color) => Log(msg, color));
            if (rs485Service.Connect(currentCmb.SelectedItem.ToString(), 115200))
            {
                btnRs485ConnectStart.Enabled = false;
                btnRs485DisconnectStart.Enabled = true;
                btnConnectManual.Enabled = false;
                btnDisconnectManual.Enabled = true;
            }
        }

        private void BtnRs485Disconnect_Click(object sender, EventArgs e)
        {
            rs485Service?.Disconnect();
            btnRs485ConnectStart.Enabled = true;
            btnRs485DisconnectStart.Enabled = false;
            btnConnectManual.Enabled = true;
            btnDisconnectManual.Enabled = false;
        }

        private void DgvCommands_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvCommands.Columns["SendButton"].Index && e.RowIndex >= 0)
            {
                if (rs485Service != null && rs485Service.IsOpen)
                {
                    var selectedCommand = (dgvCommands.Rows[e.RowIndex].DataBoundItem as CommandData);
                    if (selectedCommand != null) { rs485Service.SendPacket(selectedCommand); }
                }
                else { Log("⚠ RS-485 포트를 먼저 연결하세요.", Color.Red); }
            }
        }

        private void BtnStartAutoMode_Click(object sender, EventArgs e)
        {
            if (rs485Service != null && rs485Service.IsOpen)
            {
                rs485Service.StartAutoMode();
                btnStartAutoMode.Enabled = false;
                btnStopAutoMode.Enabled = true;
            }
            else { Log("⚠ RS-485 포트를 먼저 연결하세요.", Color.Red); }
        }

        private void BtnStopAutoMode_Click(object sender, EventArgs e)
        {
            if (rs485Service != null)
            {
                rs485Service.StopAutoMode();
                btnStartAutoMode.Enabled = true;
                btnStopAutoMode.Enabled = false;
            }
        }

        private void BtnSendCustomPacket_Click(object sender, EventArgs e)
        {
            if (rs485Service == null || !rs485Service.IsOpen) { Log("⚠ RS-485 포트를 먼저 연결하세요.", Color.Red); return; }

            string inputText = txtCustomPacket.Text.Trim();
            string[] hexValues = inputText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (hexValues.Length != 16) { Log($"❌ 데이터는 반드시 16-byte여야 합니다.", Color.Red); return; }
            try
            {
                byte[] fullPacket = hexValues.Select(hex => Convert.ToByte(hex, 16)).ToArray();
                rs485Service.SendRawPacket(fullPacket, "사용자 정의 패킷");
            }
            catch (Exception ex) { Log($"❌ 데이터 변환 오류: {ex.Message}", Color.Red); }
        }

        private void BtnProcessSimulatedReceive_Click(object sender, EventArgs e)
        {
            string inputText = txtSimulatedReceive.Text.Trim();
            string[] hexValues = inputText.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (hexValues.Length != 16) { Log($"❌ 데이터는 반드시 16-byte여야 합니다.", Color.Red); return; }
            try
            {
                byte[] fullPacket = hexValues.Select(hex => Convert.ToByte(hex, 16)).ToArray();
                rs485Service?.ProcessSimulatedPacket(fullPacket);
            }
            catch (Exception ex) { Log($"❌ 데이터 변환 오류: {ex.Message}", Color.Red); }
        }

        #endregion

        #region 공용 메서드
        public void Log(string message, Color color)
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (txtLog.InvokeRequired)
            {
                try { txtLog.Invoke((MethodInvoker)delegate { Log(message, color); }); }
                catch (ObjectDisposedException) { /* 폼 닫을 때 무시 */ }
            }
            else
            {
                if (txtLog.TextLength > 50000) txtLog.Clear();
                txtLog.SelectionStart = txtLog.TextLength;
                txtLog.SelectionLength = 0;
                txtLog.SelectionColor = color;
                txtLog.AppendText($"{DateTime.Now:HH:mm:ss} - {message}\r\n");
                txtLog.SelectionColor = txtLog.ForeColor;
                txtLog.ScrollToCaret();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            rs485Service?.Disconnect();
            modbusService?.Disconnect();
            tcpService?.Disconnect();
            base.OnFormClosing(e);
        }
        #endregion
    }
}

