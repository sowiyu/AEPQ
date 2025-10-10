using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AEPQ.Services
{
    public class CommandData
    {
        public string Description { get; set; }
        public byte[] Data { get; set; }
    }

    /// <summary>
    /// CRC-16/BUYPASS 계산을 위한 정적 클래스 (사용자 제공 코드로 수정)
    /// </summary>
    public static class Crc16Buypass
    {
        private const ushort Polynomial = 0x8005; // CRC-16/BUYPASS 폴리노미얼
        private static readonly ushort[] Table = new ushort[256];

        static Crc16Buypass()
        {
            for (ushort i = 0; i < 256; ++i)
            {
                ushort crc = i;
                for (int j = 0; j < 8; j++)
                {
                    if ((crc & 0x0001) != 0)
                        crc = (ushort)((crc >> 1) ^ Polynomial);
                    else
                        crc >>= 1;
                }
                Table[i] = crc;
            }
        }

        public static ushort ComputeChecksum(byte[] bytes)
        {
            ushort crc = 0x0000; // 초기값
            foreach (byte b in bytes)
            {
                crc ^= (ushort)(b << 8);
                for (int i = 0; i < 8; i++)
                {
                    if ((crc & 0x8000) != 0)
                        crc = (ushort)((crc << 1) ^ 0x8005);
                    else
                        crc <<= 1;
                }
            }
            return crc;
        }

    }

    public class Rs485Service
    {
        private enum AutoModeState { Idle, VacuumsOn, ReadyForBreak, Breaking }
        private AutoModeState currentAutoModeState1 = AutoModeState.Idle;
        private AutoModeState currentAutoModeState2 = AutoModeState.Idle;
        private AutoModeState currentAutoModeStateTest = AutoModeState.Idle;
        public bool IsAutoModeRunning { get; private set; } = false;

        private readonly SerialPort serialPort;
        private readonly Action<string, Color> logger;
        private readonly List<byte> receiveBuffer = new List<byte>();
        private CancellationTokenSource pollingCts;

        private const byte SEND_STX = 0x22, SEND_ETX = 0x33, SEND_ADDR = 0x03, SEND_CMD = 0x85;
        private const byte RECV_STX = 0x44;
        private const int PACKET_LENGTH = 16;
        // --- 자동 모드 트리거 신호 정의 ---
        //private readonly byte[] triggerSignal1 = { 0x44, 0x10, 0x03, 0x85, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF5, 0x1A, 0x00, 0x00, 0x57, 0x3A, 0x55 };
        //private readonly byte[] triggerSignal2 = { 0x44, 0x10, 0x03, 0x85, 0x00, 0x00, 0x00, 0x00, 0x00, 0xF5, 0x0A, 0x04, 0x00, 0x4E, 0x7A, 0x55 };
        private byte[] outputState1 = new byte[9];
        private byte[] outputState2 = new byte[9];
        public bool IsOpen => serialPort.IsOpen;

        public Rs485Service(Action<string, Color> logAction)
        {
            serialPort = new SerialPort();
            logger = logAction;
        }

        public bool Connect(string portName, int baudRate)
        {
            try
            {
                if (serialPort.IsOpen) serialPort.Close();
                serialPort.PortName = portName;
                serialPort.BaudRate = baudRate;
                serialPort.DataReceived -= OnDataReceived; // 중복 등록 방지
                serialPort.DataReceived += OnDataReceived;
                serialPort.Open();
                logger($"✅ {serialPort.PortName} 포트가 연결되었습니다.", Color.Green);
                return true;
            }
            catch (Exception ex)
            {
                logger($"❌ RS-485 연결 실패: {ex.Message}", Color.Red);
                return false;
            }
        }

        public void Disconnect()
        {
            if (IsAutoModeRunning) StopAutoMode();
            if (serialPort.IsOpen)
            {
                serialPort.Close();
                logger($"🔌 {serialPort.PortName} 포트가 해제되었습니다.", Color.Black);
            }
        }

        public void StartAutoMode()
        {
            if (IsAutoModeRunning) return;
            IsAutoModeRunning = true;
            currentAutoModeState1 = AutoModeState.Idle;
            currentAutoModeState2 = AutoModeState.Idle;
            pollingCts = new CancellationTokenSource();
            Task.Run(() => PollingLoop(pollingCts.Token));
        }

        public void StopAutoMode()
        {
            if (!IsAutoModeRunning) return;
            IsAutoModeRunning = false;
            pollingCts?.Cancel();
        }

        private async Task PollingLoop(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                try
                {
       
                    // 이 메소드는 outputState1과 outputState2를 보고 알아서
                    // '둘 다 끔', '1번만 켬', '2번만 켬', '둘 다 켬' 상태를 조합하여 보내줍니다.
                    UpdateAndSendCombinedOutput("상태 유지 Polling");

                    await Task.Delay(100, token); // 3초는 너무 길 수 있으므로 1초로 줄이는 것을 권장합니다.
                }
                catch (TaskCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    logger($"- Polling Loop Error: {ex.Message}", Color.Red);
                }
            }
        }

        private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                if (!serialPort.IsOpen) return;
                int bytesToRead = serialPort.BytesToRead;
                if (bytesToRead > 0)
                {
                    byte[] buffer = new byte[bytesToRead];
                    serialPort.Read(buffer, 0, bytesToRead);

                    lock (receiveBuffer)
                    {
                        receiveBuffer.AddRange(buffer);
                    }
                    ProcessBuffer();
                }
            }
            catch (Exception) { /* 포트가 닫힐 때 예외 발생 가능, 무시 */ }
        }

        private void ProcessBuffer()
        {
            List<byte[]> packetsToProcess = new List<byte[]>();
            lock (receiveBuffer)
            {
                while (receiveBuffer.Count >= PACKET_LENGTH)
                {
                    int startIndex = receiveBuffer.FindIndex(b => b == RECV_STX);
                    if (startIndex == -1) { receiveBuffer.Clear(); return; }
                    if (startIndex > 0) { receiveBuffer.RemoveRange(0, startIndex); }
                    if (receiveBuffer.Count < PACKET_LENGTH) { break; }

                    byte[] packet = receiveBuffer.GetRange(0, PACKET_LENGTH).ToArray();
                    packetsToProcess.Add(packet);
                    receiveBuffer.RemoveRange(0, PACKET_LENGTH);
                }
            }

            foreach (var packet in packetsToProcess)
            {
                ProcessPacket(packet);
            }
        }

        public void ProcessSimulatedPacket(byte[] packet)
        {
            logger($"📥 수신 (시뮬레이션): {BitConverter.ToString(packet).Replace("-", " ")}", Color.DarkGreen);
            ProcessPacket(packet);
        }

        private void ProcessPacket(byte[] packet)
        {
            logger($"📥 수신: {BitConverter.ToString(packet).Replace("-", " ")}", Color.DarkGreen);
            if (packet.Length < 12) return; // 최소 길이를 확인하여 안정성 확보

            if (IsAutoModeRunning)
            {
                // 각 핸드툴의 트리거 신호가 있는지 미리 확인
                bool trigger1 = (packet[10] & 0xF0) == 0x10;
                bool trigger2 = (packet[11] & 0x0F) == 0x04;

                // --- ★★★ 여기가 핵심 수정 사항 ★★★ ---
                // 1. 동시 시작 조건: 두 트리거가 모두 감지되고, 두 모드가 모두 대기 상태일 때
                if (trigger1 && trigger2 && currentAutoModeState1 == AutoModeState.Idle && currentAutoModeState2 == AutoModeState.Idle)
                {
                    logger("✨ 핸드툴 1+2 동시 눌림 감지!", Color.Cyan);

                    // 두 상태를 한 번에 업데이트
                    outputState1 = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 };
                    outputState2 = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 };
                    UpdateAndSendCombinedOutput("핸드툴 1+2 동시 진공 ON");

                    // 두 모드의 상태를 모두 변경
                    currentAutoModeState1 = AutoModeState.VacuumsOn;
                    currentAutoModeState2 = AutoModeState.VacuumsOn;
                }
                // 2. 동시 시작 조건이 아닐 경우에만 개별 핸들러 호출
                else
                {
                    HandleAutoMode1(packet);
                    HandleAutoMode2(packet);
                }
            }
        }


        private void HandleAutoMode1(byte[] packet)
        {
            // 이 핸들러는 자신의 상태가 Idle일 때만 첫 동작을 시작해야 함
            if (currentAutoModeState1 != AutoModeState.Idle && currentAutoModeState1 != AutoModeState.VacuumsOn && currentAutoModeState1 != AutoModeState.ReadyForBreak)
            {
                return;
            }


            byte triggerByte1 = packet[10];

            switch (currentAutoModeState1)
            {
                case AutoModeState.Idle:
                    // 오직 핸드툴 1의 신호만 확인
                    if ((triggerByte1 & 0xF0) == 0x10)
                    {
                        logger("✨ 핸드툴 1차 눌림 감지! (모드1)", Color.Magenta);
                        outputState1 = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 };
                        UpdateAndSendCombinedOutput("핸드툴 1 진공 ON");
                        currentAutoModeState1 = AutoModeState.VacuumsOn;
                    }
                    break;

                case AutoModeState.VacuumsOn:
                    if ((triggerByte1 & 0xF0) != 0x10)
                    {
                        logger("...핸드툴 릴리즈 감지. (모드1)", Color.CornflowerBlue);
                        currentAutoModeState1 = AutoModeState.ReadyForBreak;
                    }
                    break;

                case AutoModeState.ReadyForBreak:
                    if ((triggerByte1 & 0xF0) == 0x10)
                    {
                        logger("✨ 핸드툴 2차 눌림 감지! (모드1)", Color.Magenta);
                        currentAutoModeState1 = AutoModeState.Breaking;
                        Task.Run(async () =>
                        {
                            outputState1 = new byte[] { 0, 0xA0, 0, 0, 0, 0, 0, 0, 0 };
                            UpdateAndSendCombinedOutput("핸드툴 1 파기");
                            await Task.Delay(1000);

                            outputState1 = new byte[9];
                            UpdateAndSendCombinedOutput("핸드툴 1 리셋");
                            currentAutoModeState1 = AutoModeState.Idle;
                            logger("✅ 자동 모드 1 사이클 완료.", Color.Green);
                        });
                    }
                    break;

                case AutoModeState.Breaking: break;
            }
        }

        private void HandleAutoMode2(byte[] packet)
        {
            if (currentAutoModeState2 != AutoModeState.Idle && currentAutoModeState2 != AutoModeState.VacuumsOn && currentAutoModeState2 != AutoModeState.ReadyForBreak)
            {
                return;
            }

            byte triggerByte2 = packet[11];

            switch (currentAutoModeState2)
            {
                case AutoModeState.Idle:
                    // 오직 핸드툴 2의 신호만 확인
                    if ((triggerByte2 & 0x0F) == 0x04)
                    {
                        logger("✨ 핸드툴 1차 눌림 감지! (모드2)", Color.Tomato);
                        outputState2 = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 };
                        UpdateAndSendCombinedOutput("핸드툴 2 진공 ON");
                        currentAutoModeState2 = AutoModeState.VacuumsOn;
                    }
                    break;

                case AutoModeState.VacuumsOn:
                    if ((triggerByte2 & 0x0F) != 0x04)
                    {
                        logger("...핸드툴 릴리즈 감지. (모드2)", Color.LightSalmon);
                        currentAutoModeState2 = AutoModeState.ReadyForBreak;
                    }
                    break;

                case AutoModeState.ReadyForBreak:
                    if ((triggerByte2 & 0x0F) == 0x04)
                    {
                        logger("✨ 핸드툴 2차 눌림 감지! (모드2)", Color.Tomato);
                        currentAutoModeState2 = AutoModeState.Breaking;
                        Task.Run(async () =>
                        {
                            outputState2 = new byte[] { 0, 0, 0x0A, 0, 0, 0, 0, 0, 0 };
                            UpdateAndSendCombinedOutput("핸드툴 2 파기");
                            await Task.Delay(1000);

                            outputState2 = new byte[9];
                            UpdateAndSendCombinedOutput("핸드툴 2 리셋");
                            currentAutoModeState2 = AutoModeState.Idle;
                            logger("✅ 자동 모드 2 사이클 완료.", Color.Green);
                        });
                    }
                    break;

                case AutoModeState.Breaking: break;
            }
        }


        public void SendPacket(CommandData command)
        {
            if (serialPort == null || !serialPort.IsOpen) { return; }
            try
            {
                byte[] packet = new byte[16];
                byte[] data = command.Data;
                byte[] crcData = new byte[12];
                crcData[0] = 0x10; crcData[1] = SEND_ADDR; crcData[2] = SEND_CMD;
                Buffer.BlockCopy(data, 0, crcData, 3, data.Length);
                ushort crcValue = Crc16Buypass.ComputeChecksum(crcData);
                packet[0] = SEND_STX; packet[1] = 0x10; packet[2] = SEND_ADDR; packet[3] = SEND_CMD;
                Buffer.BlockCopy(data, 0, packet, 4, data.Length);
                packet[13] = (byte)((crcValue >> 8) & 0xFF); packet[14] = (byte)(crcValue & 0xFF);
                packet[15] = SEND_ETX;
                serialPort.Write(packet, 0, packet.Length);

                // 로그 필터링
                // if (command.Description.Contains("Polling") == false && command.Description.Contains("유지") == false)
                // {
                logger($"📤 [{command.Description}] 전송: {BitConverter.ToString(packet).Replace("-", " ")}", Color.Blue);
                // }
            }
            catch (Exception ex) { logger($"❌ 전송 실패: {ex.Message}", Color.Red); }
        }



        public void SendRawPacket(byte[] packet, string description)
        {
            if (serialPort != null && serialPort.IsOpen)
            {
                try
                {
                    serialPort.Write(packet, 0, packet.Length);
                    if (!description.Contains("Polling") && !description.Contains("유지"))
                    {
                        //logger($"📤 [{description}] 전송: {BitConverter.ToString(packet).Replace("-", " ")}", Color.Blue);
                    }
                }
                catch (Exception ex) { logger($"❌ RS-485 Raw 전송 실패: {ex.Message}", Color.Red); }
            }
        }

        // outputState1과 outputState2를 합쳐서 최종 패킷을 만들어 전송하는 메소드
        private void UpdateAndSendCombinedOutput(string description)
        {
            var combinedData = new byte[9];
            // 비트 OR 연산을 사용하여 두 상태를 합칩니다.
            for (int i = 0; i < 9; i++)
            {
                combinedData[i] = (byte)(outputState1[i] | outputState2[i]);
            }
            // 합쳐진 최종 명령어를 한 번만 보냅니다.
            SendPacket(new CommandData { Description = description, Data = combinedData });
        }
    }
}

