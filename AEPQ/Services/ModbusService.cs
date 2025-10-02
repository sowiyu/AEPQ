using Modbus.Device;
using System;
using System.Drawing;
using System.Net.Sockets;
using System.Linq;

namespace AEPQ.Services
{
    /// <summary>
    /// NModbus4 라이브러리를 사용하여 Modbus TCP/IP 통신을 관리하는 서비스 클래스입니다.
    /// </summary>
    public class ModbusService : IDisposable
    {
        private readonly string ipAddress;
        private readonly int port;
        private readonly byte slaveId = 1; // 대부분의 장비는 ID 1을 사용합니다.
        private TcpClient tcpClient;
        private IModbusMaster modbusMaster;
        private readonly Action<string, Color> logger;

        /// <summary>
        /// Modbus Master가 생성되어 있고, 내부 TcpClient가 연결된 상태인지 확인합니다.
        /// </summary>
        public bool IsConnected => modbusMaster != null && (tcpClient?.Connected ?? false);

        public ModbusService(string ipAddress, int port, Action<string, Color> logAction)
        {
            this.ipAddress = ipAddress;
            this.port = port;
            logger = logAction;
        }

        /// <summary>
        /// Modbus 서버에 연결합니다.
        /// </summary>
        /// <returns>연결 성공 여부</returns>
        public bool Connect()
        {
            try
            {
                if (IsConnected)
                {
                    logger("ℹ️ Modbus TCP/IP는 이미 연결되어 있습니다.", Color.Orange);
                    return true;
                }

                // 이전 연결이 남아있을 경우 정리
                Disconnect();

                tcpClient = new TcpClient();
                // 연결 타임아웃을 설정하여 무한 대기를 방지합니다.
                var connectTask = tcpClient.ConnectAsync(ipAddress, port);
                if (!connectTask.Wait(2000)) // 2초 타임아웃
                {
                    throw new TimeoutException("Modbus 서버 연결 시간이 초과되었습니다.");
                }

                modbusMaster = ModbusIpMaster.CreateIp(tcpClient);

                // --- 안정성 개선 코드 추가 ---
                // 응답 대기 시간을 2초로 넉넉하게 설정합니다.
                modbusMaster.Transport.ReadTimeout = 2000;
                modbusMaster.Transport.WriteTimeout = 2000;

                logger($"✅ Modbus TCP/IP 연결 성공 ({ipAddress}:{port})", Color.Green);
                return true;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus TCP/IP 연결 실패: {ex.Message}", Color.Red);
                // 연결 실패 시 리소스 완전 정리
                Disconnect();
                return false;
            }
        }

        /// <summary>
        /// Modbus 서버와의 연결을 해제합니다.
        /// </summary>
        public void Disconnect()
        {
            // IModbusMaster와 TcpClient는 함께 관리되어야 합니다.
            modbusMaster?.Dispose();
            modbusMaster = null;

            tcpClient?.Dispose();
            tcpClient = null;
        }

        /// <summary>
        /// 단일 레지스터에 값을 씁니다. (Write Single Register)
        /// </summary>
        /// <param name="address">쓸 레지스터 주소 (0-based)</param>
        /// <param name="value">쓸 값</param>
        /// <returns>쓰기 성공 여부</returns>
        public bool WriteRegister(int address, int value)
        {
            if (!IsConnected)
            {
                logger("❌ Modbus 쓰기 실패: 연결되어 있지 않습니다.", Color.Red);
                return false;
            }

            try
            {
                // NModbus4는 주소와 값을 ushort 타입으로 받습니다.
                ushort registerAddress = (ushort)address;
                ushort registerValue = (ushort)value;

                logger($"  - Modbus: 주소 {address}에 값 {value} 쓰기 시도...", Color.DarkGray);
                modbusMaster.WriteSingleRegister(slaveId, registerAddress, registerValue);
                logger($"  - Modbus: 쓰기 성공.", Color.Blue);
                return true;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus 쓰기 오류 (주소: {address}): {ex.GetType().Name} - {ex.Message}", Color.Red);
                // 통신 오류 발생 시 연결 상태를 다시 확인하고 정리할 수 있습니다.
                if (!tcpClient.Connected)
                {
                    logger("🔌 Modbus TCP/IP 연결이 끊어진 것을 확인했습니다.", Color.Black);
                    Disconnect();
                }
                return false;
            }
        }

        /// <summary>
        /// 여러 개의 Holding Registers를 읽습니다.
        /// </summary>
        /// <param name="startingAddress">읽기 시작할 주소 (0-based)</param>
        /// <param name="quantity">읽을 레지스터의 수</param>
        /// <returns>읽은 값의 배열, 실패 시 null</returns>
        public int[] ReadHoldingRegisters(int startingAddress, int quantity)
        {
            if (!IsConnected)
            {
                logger("❌ Modbus 읽기 실패: 연결되어 있지 않습니다.", Color.Red);
                return null;
            }
            try
            {
                // NModbus4는 주소와 개수를 ushort 타입으로 받습니다.
                ushort startAddr = (ushort)startingAddress;
                ushort numRegisters = (ushort)quantity;

                logger($"  - Modbus: 주소 {startingAddress}부터 {quantity}개 읽기 시도...", Color.DarkGray);
                ushort[] values = modbusMaster.ReadHoldingRegisters(slaveId, startAddr, numRegisters);

                // 결과를 int[] 배열로 변환합니다.
                int[] intValues = values.Select(v => (int)v).ToArray();
                logger($"  - Modbus: 읽기 성공. 수신 값: {string.Join(", ", intValues)}", Color.Blue);
                return intValues;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus 읽기 오류 (주소: {startingAddress}): {ex.GetType().Name} - {ex.Message}", Color.Red);
                if (!tcpClient.Connected)
                {
                    logger("🔌 Modbus TCP/IP 연결이 끊어진 것을 확인했습니다.", Color.Black);
                    Disconnect();
                }
                return null;
            }
        }

        /// <summary>
        /// 서비스 종료 시 리소스를 정리합니다.
        /// </summary>
        public void Dispose()
        {
            Disconnect();
            logger("🔌 Modbus TCP/IP 연결이 해제되었습니다.", Color.Black);
        }
    }
}

