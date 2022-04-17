#include <iostream>
#include <fstream>
#include <iostream>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>
#include <string>
#include <cstring>
#include <vector>

const std::string IN_VEC_PATH_BASE = "testdata/input_test_vec";
const std::string RECV_FILE_PATH_BASE = "testdata/recv_data";
const unsigned int PORTNO = 1234;
const char* IPADDR = "127.0.0.1";
const unsigned int RECV_BUF_SIZE = 3000;

void communicateTcp(
	const char* send_cstr,
	unsigned int send_size,
	char* recv_cstr
)
{
	int sockfd = socket(AF_INET, SOCK_STREAM, 0);
	if(sockfd < 0) {
		std::cout << "Error socket:" << std::strerror(errno);
		exit(1);
	}
	
	struct sockaddr_in addr;
	memset(&addr, 0, sizeof(struct sockaddr_in));
	addr.sin_family = AF_INET;
	addr.sin_port = htons(PORTNO);
	addr.sin_addr.s_addr = inet_addr(IPADDR);
	
	connect(sockfd, (struct sockaddr *)&addr, sizeof(struct sockaddr_in));
	
	/* send */
	send(sockfd, send_cstr, send_size, 0);
	std::cout << "send : " << send_cstr << std::endl;
	
	/* receive */
	recv(sockfd, recv_cstr, RECV_BUF_SIZE, 0);
	std::cout << "recv : "  << recv_cstr << std::endl;
	
	std::cout << std::endl;
	
	close(sockfd);
}

bool readInputVecFile(
	std::string& send_sstr,
	unsigned int fileidx
)
{
	/* open inputvecfile */
	std::string invecpath;
	invecpath = IN_VEC_PATH_BASE + std::to_string(fileidx);
	std::ifstream ifs;
	ifs.open( invecpath );
	if (!ifs.is_open()) {
		return false;
	}
	
	/* combine multiple lines into one line */
	while (!ifs.eof()) {
		std::string line;
		std::getline(ifs, line);
		send_sstr += line + ',';
	//	std::cout << line << std::endl;
	}
	send_sstr.erase(send_sstr.length() - 2, 2);
//	std::cout << send_sstr.length() << " : " << send_sstr << std::endl;
	
	/* close inputvecfile */
	ifs.close();
	
	return true;
}

bool writeRecvDataFile(
	std::string& recv_sstr,
	unsigned int fileidx
)
{
	/* open recvdatafile */
	std::ofstream ofs;
	ofs.open(RECV_FILE_PATH_BASE + std::to_string(fileidx));
	if (!ofs.is_open()) {
		return false;
	}
	
	/* split receive messages with delimiter */
	std::size_t pos;
	std::string delimiter = ",";
	std::vector<std::string> recv_swords;
	while ((pos = recv_sstr.find(delimiter)) != std::string::npos) {
		recv_swords.push_back(recv_sstr.substr(0, pos));
		recv_sstr.erase(0, pos + delimiter.length());
	}
	recv_swords.push_back(recv_sstr);
	
	/* output receive messages to recvdatafile */
	ofs << recv_swords[0] << ',' << recv_swords[1] << std::endl;
	ofs << recv_swords[2] << ',' << recv_swords[3] << ',' << recv_swords[4] << std::endl;
	
	/* close recvdatafile */
	ofs.close();
	
	return true;
}

int main(int argc, char *argv[])
{
	unsigned int fileidx = 0;
	if (argc > 1) {
		fileidx = atoi(argv[1]);
	}
	
	while (1) {
		bool result;
		
		/* read invecfile */
		std::string send_sstr;
		result = readInputVecFile(send_sstr, fileidx);
		if (result == false) {
			break;
		}
		
		/* communicate tcp */
		char recv_cstr[RECV_BUF_SIZE];
		memset(recv_cstr, '\0', RECV_BUF_SIZE);
		communicateTcp(send_sstr.c_str(), send_sstr.length(), recv_cstr);
		
		/* write recvdatafile */
		std::string recv_sstr = recv_cstr;
		result = writeRecvDataFile(recv_sstr, fileidx);
		if (result == false) {
			break;
		}
		
		fileidx++;
	}
}
