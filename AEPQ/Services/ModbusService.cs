using EasyModbus;
using System;
using System.Drawing;

namespace AEPQ.Services
{
    /// <summary>
    /// EasyModbus 라이브러리를 사용하여 Modbus TCP/IP 통신을 관리하는 서비스 클래스입니다.
    /// </summary>
    public class ModbusService
    {
        private readonly ModbusClient modbusClient;
        private readonly Action<string, Color> logger;

        public bool IsConnected => modbusClient?.Connected ?? false;

        public ModbusService(string ipAddress, int port, Action<string, Color> logAction)
        {
            modbusClient = new ModbusClient(ipAddress, port);

            // --- 안정성 개선 코드 추가 ---
            // 1. Unit ID(Slave ID)를 설정합니다. 대부분의 장비는 ID 1을 사용합니다.
            //    만약 장비 설정이 다르다면 이 값을 변경해야 합니다.
            modbusClient.UnitIdentifier = 1;

            // 2. 응답 대기 시간을 2초로 넉넉하게 설정합니다.
            modbusClient.ConnectionTimeout = 2000;

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
                modbusClient.Connect();
                logger($"✅ Modbus TCP/IP 연결 성공 ({modbusClient.IPAddress}:{modbusClient.Port})", Color.Green);
                return true;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus TCP/IP 연결 실패: {ex.Message}", Color.Red);
                return false;
            }
        }

        /// <summary>
        /// Modbus 서버와의 연결을 해제합니다.
        /// </summary>
        public void Disconnect()
        {
            try
            {
                if (IsConnected)
                {
                    modbusClient.Disconnect();
                    logger("🔌 Modbus TCP/IP 연결이 해제되었습니다.", Color.Black);
                }
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus TCP/IP 해제 실패: {ex.Message}", Color.Red);
            }
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
                // EasyModbus는 0부터 시작하는 주소를 사용합니다.
                // 예: address 129는 PLC의 40130번 주소를 의미합니다.
                logger($"  - Modbus: 주소 {address}에 값 {value} 쓰기 시도...", Color.DarkGray);
                modbusClient.WriteSingleRegister(address, value);
                logger($"  - Modbus: 쓰기 성공.", Color.Blue);
                return true;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus 쓰기 오류 (주소: {address}): {ex.GetType().Name} - {ex.Message}", Color.Red);
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
                // EasyModbus는 0부터 시작하는 주소를 사용합니다.
                // 예: startingAddress 149는 PLC의 40150번 주소부터 읽기 시작합니다.
                logger($"  - Modbus: 주소 {startingAddress}부터 {quantity}개 읽기 시도...", Color.DarkGray);
                int[] values = modbusClient.ReadHoldingRegisters(startingAddress, quantity);
                logger($"  - Modbus: 읽기 성공. 수신 값: {string.Join(", ", values)}", Color.Blue);
                return values;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus 읽기 오류 (주소: {startingAddress}): {ex.GetType().Name} - {ex.Message}", Color.Red);
                return null;
            }
        }
    }
}

