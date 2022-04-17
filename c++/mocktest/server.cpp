#include <iostream>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>
#include <string>
#include <cstring>

//const unsigned int RECV_BUF_SIZE = 100;
const unsigned int RECV_BUF_SIZE = 3000;
const unsigned int SEND_BUF_SIZE = 3000;
const unsigned int PORTNO = 1234;
const char* IPADDR = "127.0.0.1";

int main()
{
	int sockfd = socket(AF_INET, SOCK_STREAM, 0);
	if(sockfd < 0){
		std::cout << "Error socket:" << std::strerror(errno);
		exit(1);
	}
	
	struct sockaddr_in addr;
	memset(&addr, 0, sizeof(struct sockaddr_in));
	addr.sin_family = AF_INET;
	addr.sin_port = htons(PORTNO);
	addr.sin_addr.s_addr = inet_addr(IPADDR);
	
	if(bind(sockfd, (struct sockaddr *)&addr, sizeof(addr)) < 0)
	{
		std::cout << "Error bind:" << std::strerror(errno);
		exit(1);
	}
	
	char recv_cstr[RECV_BUF_SIZE];
	char send_cstr[SEND_BUF_SIZE];
	
	while(1)
	{
		std::cout << "waiting for message..." << std::endl;
		
		if(listen(sockfd,SOMAXCONN) < 0){
			std::cout << "Error listen:" << std::strerror(errno);
			close(sockfd);
			exit(1);
		}
		
		struct sockaddr_in get_addr;
		socklen_t len = sizeof(struct sockaddr_in);
		int connect = accept(sockfd, (struct sockaddr *)&get_addr, &len);
		if(connect < 0){
			std::cout << "Error accept:" << std::strerror(errno);
			exit(1);
		}
		
		/* receive */
		memset(recv_cstr, '\0', RECV_BUF_SIZE);
		recv(connect, recv_cstr, RECV_BUF_SIZE, 0);
		std::cout << "recv : " << recv_cstr << std::endl;
		
		// TODO:パース処理実装
		// TODO:演算処理実装
		// TODO:文字列まとめ処理実装
		memset(send_cstr, '\0', SEND_BUF_SIZE);
		memcpy(send_cstr, recv_cstr, RECV_BUF_SIZE);
		//std::cout << recv_cstr << std::endl;
		//std::cout << send_cstr << std::endl;
		
		/* send */
		send(connect, send_cstr, SEND_BUF_SIZE, 0);
		std::cout << "send : " << send_cstr << std::endl;
		
		std::cout << std::endl;
		
		close(connect);
	}
	
	close(sockfd);
	
	return 0;
}
