#include <iostream>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>
#include <string>
#include <cstring>

const unsigned int BUF_SIZE = 3000U;
const char* ipaddr = "127.0.0.1";
const unsigned int portno = 1234;

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
	addr.sin_port = htons(portno);
	addr.sin_addr.s_addr = inet_addr(ipaddr);
	
	if(bind(sockfd, (struct sockaddr *)&addr, sizeof(addr)) < 0)
	{
		std::cout << "Error bind:" << std::strerror(errno);
		exit(1);
	}
	
	char buffer[BUF_SIZE];
	
	while(1)
	{
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
		
		memset(buffer, '\0', sizeof(buffer));
		
		std::cout << "=== waiting for receive ===" << std::endl;
		
		//receive
		recv(connect, buffer, BUF_SIZE, 0);
		std::cout << "recv : " << buffer << std::endl;
		
		//send
		send(connect, buffer, BUF_SIZE, 0);
		std::cout << "send : " << buffer << std::endl;
		
		std::cout << std::endl;
		
		close(connect);
	}
	
	close(sockfd);
	
	return 0;
}
