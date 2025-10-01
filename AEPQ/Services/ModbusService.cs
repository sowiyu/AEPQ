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
                    logger("이미 Modbus 서버에 연결되어 있습니다.", Color.Orange);
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
        /// <param name="address">쓸 레지스터 주소</param>
        /// <param name="value">쓸 값</param>
        /// <returns>쓰기 성공 여부</returns>
        public bool WriteRegister(int address, int value)
        {
            try
            {
                if (!IsConnected)
                {
                    logger("⚠ Modbus가 연결되지 않아 쓰기 작업을 수행할 수 없습니다.", Color.Orange);
                    return false;
                }
                modbusClient.WriteSingleRegister(address, value);
                logger($"  - Modbus Write: 주소 {address}에 값 {value} 쓰기 완료", Color.Blue);
                return true;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus 쓰기 오류 (주소: {address}): {ex.Message}", Color.Red);
                return false;
            }
        }

        /// <summary>
        /// 여러 개의 Holding Registers를 읽습니다.
        /// </summary>
        /// <param name="startingAddress">읽기 시작할 주소</param>
        /// <param name="quantity">읽을 레지스터의 수</param>
        /// <returns>읽은 값의 배열, 실패 시 null</returns>
        public int[] ReadHoldingRegisters(int startingAddress, int quantity)
        {
            try
            {
                if (!IsConnected)
                {
                    logger("⚠ Modbus가 연결되지 않아 읽기 작업을 수행할 수 없습니다.", Color.Orange);
                    return null;
                }
                int[] values = modbusClient.ReadHoldingRegisters(startingAddress, quantity);
                logger($"  - Modbus Read: 주소 {startingAddress}부터 {quantity}개 읽기 완료", Color.Blue);
                return values;
            }
            catch (Exception ex)
            {
                logger($"❌ Modbus 읽기 오류 (주소: {startingAddress}): {ex.Message}", Color.Red);
                return null;
            }
        }
    }
}
