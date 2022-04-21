#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>

//const unsigned int RECV_BUF_SIZE = 100;
const unsigned int RECV_BUF_SIZE = 3000;
const unsigned int SEND_BUF_SIZE = 3000;
const unsigned int PORTNO = 1234;
const char* IPADDR = "127.0.0.1";
const char DELIMITER = ',';

int dim_x_;
double x_[10];
char flame_id_[100];
char flame_id2_[100];
char flame_id3_[100];


#define GETMSG_DBL(outptr) \
do { \
	memcpy(outptr, p, sizeof(double)); \
	p += sizeof(double); \
} while (0);
#define GETMSG_INT(outptr) \
do { \
	memcpy(outptr, p, sizeof(int)); \
	p += sizeof(int); \
} while (0);
#define GETMSG_STR(outptr) \
do { \
	strcpy(outptr, p); \
	p += strlen(outptr) + 1; \
} while (0);

int main()
{
#if 1
	int sockfd = socket(AF_INET, SOCK_STREAM, 0);
	if(sockfd < 0){
		fprintf(stderr, "Error socket");
		return 1;
	}
	
	struct sockaddr_in addr;
	memset((char*)&addr, 0, (long)sizeof(struct sockaddr_in));
	addr.sin_family = AF_INET;
	addr.sin_port = htons(PORTNO);
	addr.sin_addr.s_addr = inet_addr(IPADDR);
	
	if( bind(sockfd, (struct sockaddr *)&addr, sizeof(addr)) < 0 )
	{
		fprintf(stderr, "Error bind");
		return 1;
	}
	
	char recv_buf[RECV_BUF_SIZE];
	char send_buf[SEND_BUF_SIZE];
	
	while(1)
	{
		printf("waiting for message...\n");
		
		if(listen(sockfd,SOMAXCONN) < 0){
			fprintf(stderr, "Error listen");
			close(sockfd);
			return 1;
		}
		
		struct sockaddr_in get_addr;
		socklen_t len = sizeof(struct sockaddr_in);
		int connect = accept(sockfd, (struct sockaddr *)&get_addr, &len);
		if(connect < 0){
			fprintf(stderr, "Error accept");
			return 1;
		}
		
		/* receive */
		memset(recv_buf, '\0', RECV_BUF_SIZE);
		recv(connect, recv_buf, RECV_BUF_SIZE, 0);
		
		char* p = recv_buf;
		double a;
		int b;
		char c[100];
		char d[100];
		
		GETMSG_DBL(&a);
		GETMSG_INT(&b);
		GETMSG_STR(c);
		GETMSG_STR(d);
		
		printf("%lf,%d,%s,%s\n", a, b, c, d);
		
		close(connect);
	}
	
	close(sockfd);
	
	return 0;
#else
	printf("%ld\n",sizeof(int));
	printf("%ld\n",sizeof(long));
	printf("%ld\n",sizeof(double));
#endif
}
