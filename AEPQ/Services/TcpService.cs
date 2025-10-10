using System;
using System.Drawing;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace AEPQ.Services
{
    /// <summary>
    /// 일반적인 TCP/IP 클라이언트 통신을 관리하는 서비스 클래스입니다.
    /// </summary>
    public class TcpService
    {
        private readonly Action<string, Color> logger;
        private TcpClient tcpClient;
        private NetworkStream stream;

        public bool IsConnected => tcpClient?.Connected ?? false;

        public TcpService(Action<string, Color> logAction)
        {
            logger = logAction;
        }

        public async Task<bool> ConnectAsync(string ipAddress, int port)
        {
            try
            {
                if (IsConnected) return true;
                tcpClient = new TcpClient();
                await tcpClient.ConnectAsync(ipAddress, port);
                stream = tcpClient.GetStream();
                logger($"✅ TCP/IP 연결 성공 ({ipAddress}:{port})", Color.Green);
                return true;
            }
            catch (Exception ex)
            {
                logger($"❌ TCP/IP 연결 실패: {ex.Message}", Color.Red);
                return false;
            }
        }

        public void Disconnect()
        {
            if (tcpClient != null)
            {
                stream?.Close();
                tcpClient.Close();
                logger("🔌 TCP/IP 연결이 해제되었습니다.", Color.Black);
            }
        }

        // TODO: 실제 프로토콜에 맞게 이 메서드를 구현해야 합니다.
        // int position 인자를 받도록 메소드 시그니처를 변경합니다.
        public async Task<string> RequestCoordinate(int position)
        {
            if (!IsConnected)
            {
                logger("⚠ TCP/IP가 연결되지 않았습니다.", Color.Orange);
                return "ERROR";
            }

            try
            {
                // --- ★★★ 여기가 핵심 변경 사항입니다 ★★★ ---
                // position 값에 따라 보낼 메시지를 결정합니다.
                string requestMessage = "align"; // 기본값 (position == 0)
                if (position == 1)
                {
                    requestMessage = "align1";
                }
                else if (position == 2)
                {
                    requestMessage = "align2";
                }

                // 1. 요청 메시지 보내기
                byte[] requestData = Encoding.ASCII.GetBytes(requestMessage + "\r");
                await stream.WriteAsync(requestData, 0, requestData.Length);
                logger($"  - TCP/IP: 좌표값 요청 전송 ('{requestMessage}')", Color.CornflowerBlue);

                // 2. 응답 메시지 받기 (기존과 동일)
                byte[] buffer = new byte[1024];
                int bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length);
                string response = Encoding.ASCII.GetString(buffer, 0, bytesRead);
                logger($"  - TCP/IP: 좌표값 '{response.Trim()}' 수신", Color.CornflowerBlue);

                return response.Trim();
            }
            catch (Exception ex)
            {
                logger($"❌ TCP/IP 좌표 요청 실패: {ex.Message}", Color.Red);
                return "ERROR";
            }
        }
    }
}
