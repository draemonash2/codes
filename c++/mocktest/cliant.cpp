#include <iostream>
#include <fstream>
#include <iostream> //標準入出力
#include <sys/socket.h> //アドレスドメイン
#include <sys/types.h> //ソケットタイプ
#include <arpa/inet.h> //バイトオーダの変換に利用
#include <unistd.h> //close()に利用
#include <string> //string型
#include <cstring> //cstring型

const std::string infiletpathbase = "testdata/input_test_vec";
const std::string outfiletpathbase = "testdata/recv_data";

const char* ConnectTcp(const char* s_str, unsigned int size, char* r_str)
{
	//ソケットの生成
	int sockfd = socket(AF_INET, SOCK_STREAM, 0); //アドレスドメイン, ソケットタイプ, プロトコル
	if(sockfd < 0){ //エラー処理
	
		std::cout << "Error socket:" << std::strerror(errno); //標準出力
		exit(1); //異常終了
	}
	
	//アドレスの生成
	struct sockaddr_in addr; //接続先の情報用の構造体(ipv4)
	memset(&addr, 0, sizeof(struct sockaddr_in)); //memsetで初期化
	addr.sin_family = AF_INET; //アドレスファミリ(ipv4)
	addr.sin_port = htons(1234); //ポート番号,htons()関数は16bitホストバイトオーダーをネットワークバイトオーダーに変換
	addr.sin_addr.s_addr = inet_addr("127.0.0.1"); //IPアドレス,inet_addr()関数はアドレスの翻訳
	
	//ソケット接続要求
	connect(sockfd, (struct sockaddr *)&addr, sizeof(struct sockaddr_in)); //ソケット, アドレスポインタ, アドレスサイズ
	
	//データ送信
	send(sockfd, s_str, size, 0); //送信
	std::cout << "send : " << s_str << std::endl;
	
	recv(sockfd, r_str, size, 0);
	std::cout << "recv : "  << r_str << std::endl; //標準出力
	
	close(sockfd); //ソケットクローズ
}

int main(int argc, char *argv[])
{
	unsigned int fileidx;
	if (argc <= 1)
	{
		fileidx = 0;
	}
	else
	{
		fileidx = atoi(argv[1]);
	}
	
	while (1)
	{
		std::string infilepath;
		infilepath = infiletpathbase + std::to_string(fileidx);
		
		std::ifstream ifs;
		ifs.open( infilepath );
		
		if (ifs.is_open())
		{
			std::string s_str;
			while (!ifs.eof())
			{
				std::string line;
				std::getline(ifs, line);
				s_str += line + ',';
			//	std::cout << line << std::endl;
			}
			s_str.erase(s_str.length()-2, 2);
		//	std::cout << s_str.length() << " : " << s_str << std::endl;
			
			char r_str[s_str.length() + 1] = {0};
			memset(r_str, '\0', sizeof(r_str));
			// send and receive
			ConnectTcp(s_str.c_str(), s_str.length(), r_str);
			std::cout << r_str << std::endl;
			std::cout << std::endl;
			
			//std::cout << s_str.c_str() << std::endl;
			
			std::ofstream ofs;
			ofs.open(outfiletpathbase + std::to_string(fileidx));
			ofs << r_str << std::endl;
			ofs.close();
			
			fileidx++;
		}
		else
		{
			break;
		}
		ifs.close();
	}
}
