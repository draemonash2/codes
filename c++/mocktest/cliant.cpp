#include <iostream>
#include <fstream>
#include <iostream>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>
#include <string>
#include <cstring>

const std::string infiletpathbase = "testdata/input_test_vec";
const std::string outfiletpathbase = "testdata/recv_data";
const char* ipaddr = "127.0.0.1";
const unsigned int portno = 1234;

const char* communicateTcp(const char* s_str, unsigned int size, char* r_str)
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
	
	connect(sockfd, (struct sockaddr *)&addr, sizeof(struct sockaddr_in));
	
	//send
	send(sockfd, s_str, size, 0);
	std::cout << "send : " << s_str << std::endl;
	
	//receive
	recv(sockfd, r_str, size, 0);
	std::cout << "recv : "  << r_str << std::endl;
	
	close(sockfd);
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
			communicateTcp(s_str.c_str(), s_str.length(), r_str);
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
