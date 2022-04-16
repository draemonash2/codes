#include <iostream> //標準入出力
#include <sys/socket.h> //アドレスドメイン
#include <sys/types.h> //ソケットタイプ
#include <arpa/inet.h> //バイトオーダの変換に利用
#include <unistd.h> //close()に利用
#include <string> //string型
#include <cstring> //cstring型

const unsigned int BUF_SIZE = 1000U;

int main(){

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
	
	//ソケット登録
	if(bind(sockfd, (struct sockaddr *)&addr, sizeof(addr)) < 0){ //ソケット, アドレスポインタ, アドレスサイズ //エラー処理
		std::cout << "Error bind:" << std::strerror(errno); //標準出力
		exit(1); //異常終了
	}
	
	char buffer[BUF_SIZE]; //受信用データ格納用
	
	while(1)
	{
		std::cout << "=== waiting receive ===" << std::endl;
		//受信待ち
		if(listen(sockfd,SOMAXCONN) < 0){ //ソケット, キューの最大長 //エラー処理
			std::cout << "Error listen:" << std::strerror(errno); //標準出力
			close(sockfd); //ソケットクローズ
			exit(1); //異常終了
		}
		
		//接続待ち
		struct sockaddr_in get_addr; //接続相手のソケットアドレス
		socklen_t len = sizeof(struct sockaddr_in); //接続相手のアドレスサイズ
		int connect = accept(sockfd, (struct sockaddr *)&get_addr, &len); //接続待ちソケット, 接続相手のソケットアドレスポインタ, 接続相手のアドレスサイズ
		
		if(connect < 0){ //エラー処理
			std::cout << "Error accept:" << std::strerror(errno); //標準出力
			exit(1); //異常終了
		}
		
		memset(buffer, '\0', sizeof(buffer));
		//受信
		recv(connect, buffer, BUF_SIZE, 0);
		std::cout << "recv : " << buffer << std::endl;
		
		//送信
		send(connect, buffer, BUF_SIZE, 0);
		std::cout << "send : " << buffer << std::endl;
		
		std::cout << std::endl;
		
		close(connect);
	}
	
	close(sockfd);
	
	return 0;
}
