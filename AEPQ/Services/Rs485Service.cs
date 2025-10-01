using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO.Ports;
using System.Linq;
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
        private const ushort Polynomial = 0x8005;
        private static readonly ushort[] Table = new ushort[256];

        static Crc16Buypass()
        {
            for (ushort i = 0; i < 256; ++i)
            {
                ushort value = 0;
                ushort temp = (ushort)(i << 8);
                for (byte j = 0; j < 8; ++j)
                {
                    if (((value ^ temp) & 0x8000) != 0)
                    {
                        value = (ushort)((value << 1) ^ Polynomial);
                    }
                    else
                    {
                        value <<= 1;
                    }
                    temp <<= 1;
                }
                Table[i] = value;
            }
        }

        public static ushort ComputeChecksum(byte[] bytes)
        {
            ushort crc = 0;
            foreach (byte b in bytes)
            {
                crc = (ushort)((crc << 8) ^ Table[((crc >> 8) ^ b) & 0xFF]);
            }
            return crc;
        }
    }

    public class Rs485Service
    {
        private enum AutoModeState { Idle, VacuumsOn, ReadyForBreak, Breaking }
        private AutoModeState currentAutoModeState1 = AutoModeState.Idle;
        private AutoModeState currentAutoModeState2 = AutoModeState.Idle;
        public bool IsAutoModeRunning { get; private set; } = false;

        private readonly SerialPort serialPort;
        private readonly Action<string, Color> logger;
        private readonly List<byte> receiveBuffer = new List<byte>();
        private CancellationTokenSource pollingCts;

        private const byte SEND_STX = 0x22, SEND_ETX = 0x33, SEND_ADDR = 0x03, SEND_CMD = 0x85;
        private const byte RECV_STX = 0x44;
        private const int PACKET_LENGTH = 16;
        private readonly byte[] triggerSignal1 = { 0x44, 0x10, 0x03, 0x85, 0x00, 0x00, 0x00, 0x00, 0x00, 0xFA, 0x12, 0x00, 0x00, 0x1B, 0x99, 0x55 };

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
                    CommandData commandToSend = null;

                    if (currentAutoModeState1 != AutoModeState.Idle)
                    {
                        if (currentAutoModeState1 == AutoModeState.VacuumsOn || currentAutoModeState1 == AutoModeState.ReadyForBreak)
                            commandToSend = new CommandData { Description = "상태 유지 1", Data = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 } };
                    }
                    else if (currentAutoModeState2 != AutoModeState.Idle)
                    {
                        if (currentAutoModeState2 == AutoModeState.VacuumsOn || currentAutoModeState2 == AutoModeState.ReadyForBreak)
                            commandToSend = new CommandData { Description = "상태 유지 2", Data = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 } };
                    }

                    commandToSend ??= new CommandData { Description = "상태 요청", Data = new byte[] { 0, 0, 0, 0, 0, 0, 0, 0, 0 } };

                    SendPacket(commandToSend);
                    await Task.Delay(100, token);
                }
                catch (TaskCanceledException) { break; }
                catch (Exception ex) { logger($"- Polling Loop Error: {ex.Message}", Color.Red); }
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
            if (IsAutoModeRunning)
            {
                HandleAutoMode1(packet);
                HandleAutoMode2(packet);
            }
        }

        private void HandleAutoMode1(byte[] packet)
        {
            if (currentAutoModeState2 != AutoModeState.Idle) return;
            switch (currentAutoModeState1)
            {
                case AutoModeState.Idle:
                    if (packet.SequenceEqual(triggerSignal1))
                    {
                        logger("✨ 핸드툴 1차 눌림 감지! (모드1)", Color.Magenta);
                        SendPacket(new CommandData { Description = "이젝터 3+4 동시 진공", Data = new byte[] { 0, 0x50, 0, 0, 0, 0, 0, 0, 0 } });
                        currentAutoModeState1 = AutoModeState.VacuumsOn;
                    }
                    break;
                case AutoModeState.VacuumsOn:
                    if (!packet.SequenceEqual(triggerSignal1))
                    {
                        logger("...핸드툴 릴리즈 감지. (모드1)", Color.CornflowerBlue);
                        currentAutoModeState1 = AutoModeState.ReadyForBreak;
                    }
                    break;
                case AutoModeState.ReadyForBreak:
                    if (packet.SequenceEqual(triggerSignal1))
                    {
                        logger("✨ 핸드툴 2차 눌림 감지! (모드1)", Color.Magenta);
                        currentAutoModeState1 = AutoModeState.Breaking;
                        Task.Run(async () =>
                        {
                            SendPacket(new CommandData { Description = "이젝터 3+4 동시 파기", Data = new byte[] { 0, 0xA0, 0, 0, 0, 0, 0, 0, 0 } });
                            await Task.Delay(1000);
                            SendPacket(new CommandData { Description = "명령 리셋", Data = new byte[9] { 0, 0, 0, 0, 0, 0, 0, 0, 0 } });
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
            if (currentAutoModeState1 != AutoModeState.Idle) return;
            byte triggerByte = packet[10]; // Data[6]
            switch (currentAutoModeState2)
            {
                case AutoModeState.Idle:
                    if (triggerByte == 0x04)
                    {
                        logger("✨ 핸드툴 1차 눌림 감지! (모드2)", Color.Tomato);
                        SendPacket(new CommandData { Description = "이젝터 5+6 동시 진공", Data = new byte[] { 0, 0, 0x05, 0, 0, 0, 0, 0, 0 } });
                        currentAutoModeState2 = AutoModeState.VacuumsOn;
                    }
                    break;
                case AutoModeState.VacuumsOn:
                    if (triggerByte != 0x04)
                    {
                        logger("...핸드툴 릴리즈 감지. (모드2)", Color.LightSalmon);
                        currentAutoModeState2 = AutoModeState.ReadyForBreak;
                    }
                    break;
                case AutoModeState.ReadyForBreak:
                    if (triggerByte == 0x04)
                    {
                        logger("✨ 핸드툴 2차 눌림 감지! (모드2)", Color.Tomato);
                        currentAutoModeState2 = AutoModeState.Breaking;
                        Task.Run(async () =>
                        {
                            SendPacket(new CommandData { Description = "이젝터 5+6 동시 파기", Data = new byte[] { 0, 0, 0x0A, 0, 0, 0, 0, 0, 0 } });
                            await Task.Delay(1000);
                            SendPacket(new CommandData { Description = "명령 리셋", Data = new byte[9] { 0, 0, 0, 0, 0, 0, 0, 0, 0 } });
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
            try
            {
                byte[] packet = new byte[16];
                byte[] data = command.Data;
                byte[] crcData = new byte[12];
                crcData[0] = 0x10; crcData[1] = 0x03; crcData[2] = 0x85;
                Buffer.BlockCopy(data, 0, crcData, 3, data.Length);
                ushort crcValue = Crc16Buypass.ComputeChecksum(crcData);
                packet[0] = SEND_STX; packet[1] = 0x10; packet[2] = SEND_ADDR; packet[3] = SEND_CMD;
                Buffer.BlockCopy(data, 0, packet, 4, data.Length);
                packet[13] = (byte)(crcValue >> 8);
                packet[14] = (byte)(crcValue & 0xFF);
                packet[15] = SEND_ETX;

                SendRawPacket(packet, command.Description);
            }
            catch (Exception ex) { logger($"❌ RS-485 전송 실패: {ex.Message}", Color.Red); }
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
                        logger($"📤 [{description}] 전송: {BitConverter.ToString(packet).Replace("-", " ")}", Color.Blue);
                    }
                }
                catch (Exception ex) { logger($"❌ RS-485 Raw 전송 실패: {ex.Message}", Color.Red); }
            }
        }
    }
}

