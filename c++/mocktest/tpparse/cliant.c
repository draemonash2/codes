#include <stdio.h>
//#include <iostream>
//#include <fstream>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>
//#include <string>
//#include <cstring>
//#include <vector>
#include <stdlib.h>
#include <string.h>

const char* IN_VEC_PATH_BASE = "testdata/input_test_vec";
const char* RECV_FILE_PATH_BASE = "testdata/recv_data";
const unsigned int PORTNO = 1234;
const char* IPADDR = "127.0.0.1";
const unsigned int RECV_BUF_SIZE = 3000;
const unsigned int SEND_BUF_SIZE = 3000;

#define SETMSG_DBL(inptr) \
do { \
	memcpy(p, inptr, sizeof(double)); \
	p += sizeof(double); \
	send_msg_size += sizeof(double); \
} while (0);
#define SETMSG_INT(inptr) \
do { \
	memcpy(p, inptr, sizeof(int)); \
	p += sizeof(int); \
	send_msg_size += sizeof(int); \
} while (0);
#define SETMSG_STR(inptr) \
do { \
	strcpy(p, inptr); \
	p += strlen(inptr) + 1; \
	send_msg_size += strlen(inptr) + 1; \
} while (0);

int main(int argc, char *argv[])
{
	unsigned int fileidx = 0;
	if (argc > 1) {
		fileidx = atoi(argv[1]);
	}
	
	/* communicate tcp */
	int sockfd = socket(AF_INET, SOCK_STREAM, 0);
	if(sockfd < 0) {
		printf("Error socket\n");
		exit(1);
	}
	
	struct sockaddr_in addr;
	memset(&addr, 0, sizeof(struct sockaddr_in));
	addr.sin_family = AF_INET;
	addr.sin_port = htons(PORTNO);
	addr.sin_addr.s_addr = inet_addr(IPADDR);
	
	connect(sockfd, (struct sockaddr *)&addr, sizeof(struct sockaddr_in));
	
	char send_sstr[SEND_BUF_SIZE];
	long send_msg_size = 0;
	char* p = send_sstr;
	
	double a = 31.0;
	int b = 10;
	char* c = "stringdesu";
	char* d = "aaa";
	
	SETMSG_DBL(&a);
	SETMSG_INT(&b);
	SETMSG_STR(c);
	SETMSG_STR(d);
	
	/* send */
	send(sockfd, send_sstr, send_msg_size, 0);
	printf("send : %s\n", send_sstr);
	
	char recv_cstr[RECV_BUF_SIZE];
	
	/* receive */
	recv(sockfd, recv_cstr, RECV_BUF_SIZE, 0);
	printf("recv : %s\n", recv_cstr);
	
	printf("\n");
	
	close(sockfd);
}
