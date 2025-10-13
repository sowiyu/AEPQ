using AEPQ.Services;
using log4net.Repository.Hierarchy;
using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Modbus.Device;
using System.Drawing;

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
        private MaterialTabSelector tabSelector;
        private Panel mainPanel;
        private TabPage tabPageStart, tabPageManual;
        public event Action<bool, bool> OnButtonEnableChanged;


        // '시작' 탭 컨트롤
        private GroupBox groupModbus, groupTcp, groupRs485Start;
        private TextBox txtModbusIp, txtModbusPort;
        private Button btnModbusConnect, btnModbusDisconnect;
        private TextBox txtTcpIp, txtTcpPort;
        private Button btnTcpConnect, btnTcpDisconnect;
        private ComboBox cmbPortStart;
        private Button btnRs485ConnectStart, btnRs485DisconnectStart, btnRs485RefreshStart;
        private Button btnStartOperation1, btnStartOperation2;
        private Button btnStartPosition1, btnStartPosition2; // <-- 이 줄 추가
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
        private Panel manualPagePanel;
        public MainForm()
        {
            InitializeComponent();

            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.EnforceBackcolorOnAllComponents = true;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.LightBlue500, Primary.LightBlue500,
                Primary.LightBlue200, Accent.LightBlue700,
                TextShade.WHITE);
        }

        private void InitializeComponent()
        {
            this.Text = "AEPQ 통합 제어 프로그램";
            this.ClientSize = new System.Drawing.Size(940, 800);

            // 1. 탭 컨트롤들 생성
            tabControl = new MaterialTabControl { Dock = DockStyle.Fill };
            tabPageStart = new TabPage { Text = "시작" };
            tabPageManual = new TabPage { Text = "단동 테스트" };
            tabControl.TabPages.Add(tabPageStart);
            tabControl.TabPages.Add(tabPageManual);

            tabSelector = new MaterialTabSelector
            {
                BaseTabControl = tabControl,
                Depth = 0,
                Dock = DockStyle.Top,
                // 폰트를 굵게 설정
                Font = new Font("Roboto", 13f, FontStyle.Bold, GraphicsUnit.Point)
            };

            // 2. 로그 박스 생성 (스크롤바 및 줄바꿈 속성 추가)
            txtLog = new RichTextBox
            {
                Dock = DockStyle.Bottom,
                Height = 200,
                ReadOnly = true,
                Font = new Font("Consolas", 9),
                BackColor = Color.FromArgb(50, 50, 50),
                ForeColor = Color.White,
                WordWrap = false, // 자동 줄바꿈 끄기
                ScrollBars = RichTextBoxScrollBars.ForcedBoth // 가로/세로 스크롤바 항상 표시
            };

            // 3. 메인 컨테이너 Panel 생성
            mainPanel = new Panel
            {
                Dock = DockStyle.Fill
            };

            // 4. Panel에 tabControl과 txtLog를 추가
            mainPanel.Controls.Add(tabControl);
            mainPanel.Controls.Add(txtLog);

            // 5. 폼(this)에 mainPanel과 tabSelector를 추가
            this.Controls.Add(mainPanel);
            this.Controls.Add(tabSelector);

            // 6. 탭 페이지 초기화
            InitializeStartTab();
            InitializeManualTestTab();
            this.Load += MainForm_Load;
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
            txtModbusIp = new TextBox { Text = "192.168.1.10", Location = new Point(80, 32), Width = 120 };
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

            txtTcpPort = new TextBox { Text = "20001", Location = new Point(195, 32), Width = 60 };
            btnTcpConnect = new Button { Text = "연결", Location = new Point(15, 70), Width = 120 };
            btnTcpConnect.Click += BtnTcpConnect_Click;
            btnTcpDisconnect = new Button { Text = "해제", Location = new Point(145, 70), Width = 120, Enabled = false };
            btnTcpDisconnect.Click += BtnTcpDisconnect_Click;
            groupTcp.Controls.Add(txtTcpIp);
            groupTcp.Controls.Add(txtTcpPort);
            groupTcp.Controls.Add(btnTcpConnect);
            groupTcp.Controls.Add(btnTcpDisconnect);
            tabPageStart.Controls.Add(groupTcp);

            // 1. '위치 시작' 버튼 2개 생성
            btnStartPosition1 = new Button { Text = "1 위치 시작", Location = new Point(50, 160), Width = 195, Height = 60, Font = new Font(this.Font.FontFamily, 16, FontStyle.Bold) };
            btnStartPosition1.Click += (s, e) => StartOperation1WithPosition(1);

            btnStartPosition2 = new Button { Text = "2 위치 시작", Location = new Point(255, 160), Width = 195, Height = 60, Font = new Font(this.Font.FontFamily, 16, FontStyle.Bold) };
            btnStartPosition2.Click += (s, e) => StartOperation1WithPosition(2);

            // 2. '위치 지정 안함' 버튼 생성 (기존 btnStartOperation1 역할)
            btnStartOperation1 = new Button { Text = "동작 1 (위치 미지정)", Location = new Point(50, 230), Width = 400, Height = 230, Font = new Font(this.Font.FontFamily, 24, FontStyle.Bold) };
            btnStartOperation1.Click += (s, e) => StartOperation1WithPosition(0);

            // 3. '동작 2 시작' 버튼 생성
            btnStartOperation2 = new Button { Text = "동작 2 시작", Location = new Point(480, 160), Width = 400, Height = 300, Font = new Font(this.Font.FontFamily, 24, FontStyle.Bold) };
            btnStartOperation2.Click += BtnStartOperation2_Click;

            // 4. 폼에 최종 버튼들 추가
            tabPageStart.Controls.Add(btnStartPosition1);
            tabPageStart.Controls.Add(btnStartPosition2);
            tabPageStart.Controls.Add(btnStartOperation1);
            tabPageStart.Controls.Add(btnStartOperation2);
        }

        private void InitializeManualTestTab()
        {
            // 1. 스크롤을 담당할 패널을 생성합니다.
            // 이 패널이 탭 페이지 전체를 덮고, 내용이 길어지면 스크롤바를 만듭니다.
            manualPagePanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true
            };
            // 생성한 패널을 '단동 테스트' 탭 페이지에 추가합니다.
            tabPageManual.Controls.Add(manualPagePanel);

            // 2. 이제부터 모든 UI 컨트롤들은 tabPageManual이 아닌 manualPagePanel에 추가합니다.
            // 이렇게 해야 패널의 스크롤 기능이 적용됩니다.

            // RS-485 연결 그룹
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
            manualPagePanel.Controls.Add(groupRS485Manual); // 스크롤 패널에 추가

            // 명령어 목록 그리드
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
            manualPagePanel.Controls.Add(dgvCommands); // 스크롤 패널에 추가

            // 자동 모드 그룹
            autoModeGroup = new GroupBox { Text = "자동 모드 (핸드툴 1 & 2)", Location = new Point(20, 350), Width = 880, Height = 55 };
            btnStartAutoMode = new Button { Text = "시작", Location = new Point(15, 20), Width = 150 };
            btnStartAutoMode.Click += BtnStartAutoMode_Click;
            btnStopAutoMode = new Button { Text = "중지", Location = new Point(180, 20), Width = 150, Enabled = false };
            btnStopAutoMode.Click += BtnStopAutoMode_Click;
            autoModeGroup.Controls.Add(btnStartAutoMode);
            autoModeGroup.Controls.Add(btnStopAutoMode);
            manualPagePanel.Controls.Add(autoModeGroup); // 스크롤 패널에 추가

            // 패킷 직접 전송 그룹
            sendGroup = new GroupBox { Text = "패킷 직접 전송", Location = new Point(20, 410), Width = 880, Height = 75 };
            sendGroup.Controls.Add(new Label { Text = "16-byte hex:", Location = new Point(15, 35), AutoSize = true });
            txtCustomPacket = new TextBox { Location = new Point(100, 32), Width = 650, Font = new Font("Consolas", 9) };
            btnSendCustomPacket = new Button { Text = "전송", Location = new Point(760, 30), Width = 100 };
            btnSendCustomPacket.Click += BtnSendCustomPacket_Click;
            sendGroup.Controls.Add(txtCustomPacket);
            sendGroup.Controls.Add(btnSendCustomPacket);
            manualPagePanel.Controls.Add(sendGroup); // 스크롤 패널에 추가

            // 수신 시뮬레이션 그룹
            simulationGroup = new GroupBox { Text = "수신 시뮬레이션", Location = new Point(20, 490), Width = 880, Height = 75 };
            simulationGroup.Controls.Add(new Label { Text = "16-byte hex:", Location = new Point(15, 35), AutoSize = true });
            txtSimulatedReceive = new TextBox { Location = new Point(100, 32), Width = 650, Font = new Font("Consolas", 9) };
            btnProcessSimulatedReceive = new Button { Text = "수신 처리", Location = new Point(760, 30), Width = 100 };
            btnProcessSimulatedReceive.Click += BtnProcessSimulatedReceive_Click;
            simulationGroup.Controls.Add(txtSimulatedReceive);
            simulationGroup.Controls.Add(btnProcessSimulatedReceive);
            manualPagePanel.Controls.Add(simulationGroup); // 스크롤 패널에 추가

            // 마지막으로 필요한 메소드들을 호출합니다.
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




        // 기존의 private async void BtnStartOperation1_Click(object sender, EventArgs e) 메소드는 삭제합니다.
        // 그 대신 아래의 새로운 메소드를 추가합니다.
        private async void StartOperation1WithPosition(int position)
        {
            string logTitle = position > 0 ? $"▶️ 동작 1 ({position} 위치) 시작..." : "▶️ 동작 1 시작 (위치 지정 안함)...";
            Log(logTitle, Color.Blue);
            btnStartPosition1.Enabled = false;
            btnStartPosition2.Enabled = false;
            btnStartOperation1.Enabled = false; // 버튼 눌리면 바로 Disable
            btnStartOperation2.Enabled = false; // 다른 버튼도 동시에 Disable 가능 (선택 사항)

            // --- 1. 연결 상태 확인 (기존과 동일) ---
            if (rs485Service == null || !rs485Service.IsOpen) { Log("  - ⚠ RS-485 포트가 연결되지 않았습니다.", Color.Orange); return; }
            if (tcpService == null || !tcpService.IsConnected) { Log("  - ⚠ TCP/IP (좌표)가 연결되지 않았습니다.", Color.Orange); return; }
            if (modbusService == null || !modbusService.IsConnected) { Log("  - ⚠ Modbus TCP/IP가 연결되지 않았습니다.", Color.Orange); return; }

            try
            {
                // 1. RS-485 실린더 전진
                //rs485Service.SendPacket(new CommandData
                //{
                //    Description = "케이스 실린더 전진",
                //    Data = new byte[] { 0, 0, 0, 0x02, 0, 0, 0, 0, 0 }
                //});

                rs485Service.SendPacket(3, 0x02, "케이스 실린더 후진");

                Log("  - RS-485: 실린더 전진 명령 전송", Color.DarkBlue);
                await Task.Delay(500);
                //rs485Service.SendPacket(new CommandData
                //{
                //    Description = "케이스 실린더 전진",
                //    Data = new byte[] { 0, 0, 0, 0x08, 0, 0, 0, 0, 0 }
                //});
                rs485Service.SendPacket(3, 0x08, "케이스 실린더 전진");
                Log("  - RS-485: 실린더 전진 명령 전송", Color.DarkBlue);
                await Task.Delay(500);


                // --- 2. TCP 좌표 수신 (기존과 동일) ---
                string coordinate = "";
                while (true)
                {
                    coordinate = await tcpService.RequestCoordinate(position);
                    if (!string.IsNullOrEmpty(coordinate) && coordinate.StartsWith("PC_Align_align_end"))
                    {
                        Log($"  - TCP 좌표 응답: {coordinate}", Color.DarkCyan);
                        break;
                    }
                    await Task.Delay(200);
                }
                Log("  - TCP: 좌표값 수신 완료", Color.CornflowerBlue);
                // --- 3. Position 값에 따라 좌표 파싱 및 Modbus 쓰기 (핵심 로직) ---
                string payload = coordinate.Replace("PC_Align_align_end_", "");
                string[] blocks = payload.Split(',');
                Log($"  - position {position}", Color.DarkCyan);

                // 3-1. '위치 미지정' 버튼 (블록 2개 처리)
                if (position == 0)
                {
                    if (blocks.Length != 2)
                    {
                        Log($"  - ❌ 좌표 블록 수가 2가 아님 (기대값: 2, 실제: {blocks.Length}). 동작을 중단합니다.", Color.Red);
                        return;
                    }
                    await WriteCoordinateBlock(blocks[0], 132); // 첫 번째 블록 -> 132, 133, 134
                    await WriteCoordinateBlock(blocks[1], 135); // 두 번째 블록 -> 135, 136, 137
                    Log("  - Modbus: 132~137번에 좌표 전송 완료", Color.DarkBlue);
                }
                // 3-2. '1 위치' 또는 '2 위치' 버튼 (블록 1개 처리)
                else if (position == 1 || position == 2)
                {
                    if (blocks.Length != 1)
                    {
                        Log($"  - ❌ 좌표 블록 수가 1이 아님 (기대값: 1, 실제: {blocks.Length}). 동작을 중단합니다.", Color.Red);
                        return;
                    }
                    // 위치에 따라 시작 주소 결정
                    int modbusAddr = (position == 1) ? 132 : 135;
                    await WriteCoordinateBlock(blocks[0], modbusAddr);
                    Log($"  - Modbus: {modbusAddr}~{modbusAddr + 2}번에 좌표 전송 완료", Color.DarkBlue);
                }

                // --- 4. 위치 준비 신호 보내고 응답 확인 (position 1, 2일 때만) ---
                if (position == 1 || position == 2)
                {
                    Log($"  - Modbus: 131번에 위치 값 {position} 쓰기", Color.DarkCyan);
                    modbusService.WriteRegister(131, position);


                    if (!modbusService.IsConnected)
                    {
                        Log("  - ℹ️ Modbus 연결이 끊어져 재연결을 시도합니다.", Color.Orange);
                        if (!modbusService.Connect())
                        {
                            Log("  - ❌ Modbus 재연결에 실패했습니다. 동작을 중단합니다.", Color.Red);
                            return; // 재연결 실패 시 작업 중단
                        }
                    }

                    // 쓰기 시도
                    if (!modbusService.WriteRegister(131, position))
                    {
                        // 쓰기 실패 시 한 번 더 재연결 및 재시도
                        Log("  - ℹ️ Modbus 쓰기 실패. 재연결 후 재시도합니다.", Color.Orange);
                        if (modbusService.Connect() && modbusService.WriteRegister(131, position))
                        {
                            Log("  - ✅ Modbus 쓰기 재시도 성공.", Color.Green);
                        }
                        else
                        {
                            Log($"  - ❌ Modbus 쓰기 최종 실패 (주소: {130}). 동작을 중단합니다.", Color.Red);
                            return; // 최종 실패 시 작업 중단
                        }
                    }


                    bool isPositionReady = false;
                    for (int i = 0; i < 5000; i++) // 최대 1000초 대기
                    {
                        int[] readVal = modbusService.ReadHoldingRegisters(151, 1);
                        if (readVal != null && readVal.Length > 0 && readVal[0] == position)
                        {
                            Log($"  - Modbus: 151번={position} 응답 확인.", Color.DarkBlue);

                            isPositionReady = true;
                            break;
                        }
                        await Task.Delay(200);
                    }

                    if (!isPositionReady)
                    {
                        Log($"  - ❌ Modbus: 151번 주소에서 응답 대기 시간 초과. 동작을 중단합니다.", Color.Red);
                        return;
                    }
                }

                // --- 5. 로봇 동작 시작 신호 보내고 응답 확인 (기존과 동일) ---
                modbusService.WriteRegister(130, 1);
                Log("  - Modbus: 130번에 1 써서 로봇 동작 시작", Color.DarkBlue);
                Thread.Sleep(1000);
                modbusService.WriteRegister(131, 0);
                Log("  - Modbus: 131번 주소 초기화 완료.", Color.DarkBlue);

                // 5. Modbus 150번 모니터링 후 130 초기화
                int a = 0;
                while (true)
                {
                    //a++;
                    //if (a > 2)
                    //{
                    //    Log("  - ❌ Modbus 150번 응답 대기 시간 초과", Color.Red);
                    //    break;
                    //}
                    int[] readVal = modbusService.ReadHoldingRegisters(150, 1);
                    if (readVal != null && readVal.Length > 0 && readVal[0] == 1)
                    {
                        modbusService.WriteRegister(130, 0); // 초기화
                        Log("  - Modbus: 150번=1 확인, 130번 초기화", Color.DarkBlue);
                        break;
                    }
                    await Task.Delay(2000);
                }

                while (true)
                {
                    //a++;
                    //if (a > 2)
                    //{
                    //    Log("  - ❌ Modbus 150번 응답 대기 시간 초과", Color.Red);
                    //    break;
                    //}
                    int[] readVal = modbusService.ReadHoldingRegisters(150, 1);
                    if (readVal != null && readVal.Length > 0 && readVal[0] == 0)
                    {
                        Log("  - Modbus: 150번=0 확인", Color.DarkBlue);
                        // 5. RS-485 실린더 후진
                        //rs485Service.SendPacket(new CommandData
                        //{
                        //    Description = "케이스 실린더 후진",
                        //    Data = new byte[] { 0, 0, 0, 0x01, 0, 0, 0, 0, 0 }
                        //});
                        rs485Service.SendPacket(3, 0x01, "케이스 실린더 후진");
                        Log("  - RS-485: 실린더 후진 명령 전송", Color.DarkBlue);
                        await Task.Delay(500);

                        //rs485Service.SendPacket(new CommandData
                        //{
                        //    Description = "케이스 실린더 후진",
                        //    Data = new byte[] { 0, 0, 0, 0x04, 0, 0, 0, 0, 0 }
                        //});

                        rs485Service.SendPacket(3, 0x04, "케이스 실린더 후진");
                        Log("  - RS-485: 실린더 후진 명령 전송", Color.DarkBlue);
                        await Task.Delay(500);
                        btnStartPosition1.Enabled = true;
                        btnStartPosition2.Enabled = true;
                        btnStartOperation1.Enabled = true; // 버튼 눌리면 바로 Disable
                        btnStartOperation2.Enabled = true; // 다른 버튼도 동시에 Disable 가능 (선택 사항)
                        break;
                    }
                    await Task.Delay(2000);
                }

                Log("✅ 동작 1 완료!", Color.Green);
            }
            catch (Exception ex)
            {
                Log($"❌ 동작 1 중 오류 발생: {ex.Message}", Color.Red);
            }
        }

        // 좌표 블록 하나를 파싱해서 Modbus에 쓰는 보조 메소드
        private async Task WriteCoordinateBlock(string block, int startAddress)
        {
            var vals = block.Split('_');
            if (vals.Length != 3)
            {
                Log($"  - ❌ 좌표 포맷 오류: {block}", Color.Red);
                return; // 오류가 있어도 일단 진행은 되도록 throw 대신 return 사용
            }

            try
            {
                int x = int.Parse(vals[0]);
                int y = int.Parse(vals[1]);
                int rz = int.Parse(vals[2]);

                if (x < 0) x = x * -1 + 10000;
                if (y < 0) y = 10000 + y * -1;
                if (rz < 0) rz = 1000 + rz * -1;

                int[] toWrite = new int[] { x, y, rz };

                for (int i = 0; i < toWrite.Length; i++)
                {
                    // --- 안정성 강화: Modbus 쓰기 전 연결 확인 및 재연결 ---
                    if (!modbusService.IsConnected)
                    {
                        Log("  - ℹ️ Modbus 연결이 끊어져 재연결을 시도합니다.", Color.Orange);
                        if (!modbusService.Connect())
                        {
                            Log("  - ❌ Modbus 재연결에 실패했습니다. 동작을 중단합니다.", Color.Red);
                            return; // 재연결 실패 시 작업 중단
                        }
                    }

                    // 쓰기 시도
                    if (!modbusService.WriteRegister(startAddress + i, toWrite[i]))
                    {
                        // 쓰기 실패 시 한 번 더 재연결 및 재시도
                        Log("  - ℹ️ Modbus 쓰기 실패. 재연결 후 재시도합니다.", Color.Orange);
                        if (modbusService.Connect() && modbusService.WriteRegister(startAddress + i, toWrite[i]))
                        {
                            Log("  - ✅ Modbus 쓰기 재시도 성공.", Color.Green);
                        }
                        else
                        {
                            Log($"  - ❌ Modbus 쓰기 최종 실패 (주소: {startAddress + i}). 동작을 중단합니다.", Color.Red);
                            return; // 최종 실패 시 작업 중단
                        }
                    }
                    await Task.Delay(50); // 안정성을 위해 딜레이
                }



            }
            catch (Exception ex)
            {
                Log($"  - ❌ 좌표 처리 중 오류 ({block}): {ex.Message}", Color.Red);
            }
        }


        private async void BtnStartOperation2_Click(object sender, EventArgs e)
        {
            if (rs485Service == null || !rs485Service.IsOpen)
            {
                Log("  - ⚠ RS-485 포트가 연결되지 않았습니다.", Color.Orange);
                return;
            }

            if (tcpService == null || !tcpService.IsConnected)
            {
                Log("  - ⚠ TCP/IP (좌표)가 연결되지 않았습니다.", Color.Orange);
                return;
            }

            if (modbusService == null || !modbusService.IsConnected)
            {
                Log("  - ⚠ Modbus TCP/IP가 연결되지 않았습니다.", Color.Orange);
                return;
            }

            btnStartOperation1.Enabled = false; // 버튼 눌리면 바로 Disable
            btnStartOperation2.Enabled = false; // 다른 버튼도 동시에 Disable 가능 (선택 사항)

            Log("▶️ 동작 2 시작...", Color.DarkCyan);

            try
            {
                // 1. RS-485 실린더 전진 명령 전송
                //rs485Service.SendPacket(new CommandData
                //{
                //    Description = "동작2 실린더 up",
                //    Data = new byte[] { 0, 0, 0, 0x10, 0, 0, 0, 0, 0 }
                //});
                rs485Service.SendPacket(3, 0x10, "동작2 실린더 up");
                Log("  - RS-485: 실린더 UP 명령 전송", Color.DarkBlue);

                await Task.Delay(50);
                Thread.Sleep(1000);

                //rs485Service.SendPacket(new CommandData
                //{
                //    Description = "동작2 실린더 전진",
                //    Data = new byte[] { 0, 0, 0x20, 0, 0, 0, 0, 0, 0 }
                //});
                rs485Service.SendPacket(2, 0x20, "동작2 실린더 전진");
                Log("  - RS-485: 실린더 전진 명령 전송", Color.DarkBlue);
                await Task.Delay(50);

                Thread.Sleep(1000);

                //rs485Service.SendPacket(new CommandData
                //{
                //    Description = "동작2 실린더 down",
                //    Data = new byte[] { 0, 0, 0, 0x20, 0, 0, 0, 0, 0 }
                //});
                rs485Service.SendPacket(3, 0x20, "동작2 실린더 down");
                Log("  - RS-485: 실린더 UP 명령 전송", Color.DarkBlue);


                await Task.Delay(100); // 명령 후 안정화 시간

                // 1-1. RS-485 응답 확인 (실린더 상태값 Read)
                // → 여기서 실제 장비 응답 패킷 파싱하는 로직 넣으시면 됨
                // ex) byte[] resp = rs485Service.ReadPacket();
                // Log($"실린더 응답: {BitConverter.ToString(resp)}", Color.Gray);

                //byte[] resp = rs485Service.ReadPacket();


                // 2. Modbus 130번에 2 써서 로봇 동작 시작
                modbusService.WriteRegister(130, 2);
                Log("  - Modbus: 130번에 2 쓰기 (로봇 동작 시작)", Color.DarkBlue);

                if (!modbusService.IsConnected)
                {
                    Log("  - ℹ️ Modbus 연결이 끊어져 재연결을 시도합니다.", Color.Orange);
                    if (!modbusService.Connect())
                    {
                        Log("  - ❌ Modbus 재연결에 실패했습니다. 동작을 중단합니다.", Color.Red);
                        return; // 재연결 실패 시 작업 중단
                    }
                }

                // 쓰기 시도
                if (!modbusService.WriteRegister(130, 2))
                {
                    // 쓰기 실패 시 한 번 더 재연결 및 재시도
                    Log("  - ℹ️ Modbus 쓰기 실패. 재연결 후 재시도합니다.", Color.Orange);
                    if (modbusService.Connect() && modbusService.WriteRegister(130, 2))
                    {
                        Log("  - ✅ Modbus 쓰기 재시도 성공.", Color.Green);
                    }
                    else
                    {
                        Log($"  - ❌ Modbus 쓰기 최종 실패 (주소: {130}). 동작을 중단합니다.", Color.Red);
                        return; // 최종 실패 시 작업 중단
                    }
                }
                await Task.Delay(50); // 안정성을 위해 딜레이



                // 3. Modbus 150번 모니터링 후 130 초기화
                int retry = 0;
                while (true)
                {
                    //retry++;
                    //if (retry > 20) // 20회 × 200ms = 4초 타임아웃
                    //{
                    //    Log("  - ❌ Modbus 150번 응답 대기 시간 초과", Color.Red);
                    //    break;
                    //}

                    int[] readVal = modbusService.ReadHoldingRegisters(150, 1);

                    if (readVal != null && readVal.Length > 0 && readVal[0] == 2)
                    {
                        modbusService.WriteRegister(130, 0); // 초기화
                        Log("  - Modbus: 150번=2 확인 → 130번 초기화 완료", Color.DarkBlue);
                        break;
                    }

                    await Task.Delay(200); // 폴링 대기
                }

                while (true)
                {
                    //retry++;
                    //if (retry > 20) // 20회 × 200ms = 4초 타임아웃
                    //{
                    //    Log("  - ❌ Modbus 150번 응답 대기 시간 초과", Color.Red);
                    //    break;
                    //}
                    int[] readVal = modbusService.ReadHoldingRegisters(150, 1);
                    if (readVal != null && readVal.Length > 0 && readVal[0] == 0)
                    {
                        Log("  - Modbus: 150번=0 확인", Color.DarkBlue);
                        // 1. RS-485 실린더 후진/ down 명령 전송
                        //rs485Service.SendPacket(new CommandData
                        //{
                        //    Description = "동작2 실린더 up",
                        //    Data = new byte[] { 0, 0, 0, 0x10, 0, 0, 0, 0, 0 }
                        //});
                        rs485Service.SendPacket(3, 0x10, "동작2 실린더 up");
                        Log("  - RS-485: 실린더 UP 명령 전송", Color.DarkBlue);

                        await Task.Delay(50);
                        Thread.Sleep(1000);
                        //rs485Service.SendPacket(new CommandData
                        //{
                        //    Description = "동작2 실린더 후진",
                        //    Data = new byte[] { 0, 0, 0x10, 0, 0, 0, 0, 0, 0 }
                        //});
                        rs485Service.SendPacket(2, 0x10, "동작2 실린더 후진");
                        Log("  - RS-485: 실린더 전진 명령 전송", Color.DarkBlue);

                        await Task.Delay(100); // 명령 후 안정화 시간
                        btnStartOperation1.Enabled = true; // 버튼 눌리면 바로 Disable
                        btnStartOperation2.Enabled = true; // 다른 버튼도 동시에 Disable 가능 (선택 사항)
                        break;
                    }
                    await Task.Delay(200); // 폴링 대기
                }


                Log("✅ 동작 2 완료!", Color.Green);
            }
            catch (Exception ex)
            {
                Log($"❌ 동작 2 중 오류 발생: {ex.Message}", Color.Red);
            }
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
                btnDisconnectManual.Enabled = true;

                cmbPortStart.SelectedIndex = currentCmb.SelectedIndex;
                cmbPortManual.SelectedIndex = currentCmb.SelectedIndex;

                cmbPortManual.Enabled = false;
                cmbPortStart.Enabled = false;
                btnConnectManual.Enabled = false;
            }
        }

        private void BtnRs485Disconnect_Click(object sender, EventArgs e)
        {
            rs485Service?.Disconnect();
            btnRs485ConnectStart.Enabled = true;
            btnRs485DisconnectStart.Enabled = false;
            cmbPortStart.Enabled = true;
            btnDisconnectManual.Enabled = false;

            cmbPortManual.Enabled = true;
            btnConnectManual.Enabled = true;
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

        private void MainForm_Load(object sender, EventArgs e)
        {
            // "COM3"이 있는 경우 자동으로 연결 시도
            if (cmbPortStart.Items.Contains("COM9"))
            {
                cmbPortStart.SelectedItem = "COM9";
                btnRs485ConnectStart.PerformClick();
            }
            // 모드버스 자동연결 시도
            btnModbusConnect.PerformClick();
            // TCP 자동연결 시도
            btnTcpConnect.PerformClick();
        }

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