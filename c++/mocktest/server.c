#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>
#include "mystring.h"

//const unsigned int RECV_BUF_SIZE = 100;
const unsigned int RECV_BUF_SIZE = 3000;
const unsigned int SEND_BUF_SIZE = 3000;
const unsigned int PORTNO = 1234;
const char* IPADDR = "127.0.0.1";
const char DELIMITER = ',';

int dim_x_;
char flame_id_[100];
char flame_id2_[100];
char flame_id3_[100];
double x_[10];

#define atoi myatoi
//#define atof myatof
#define strcpy mystrcpy
#define memset mymemset

void parse_receive_msg( const char* instr )
{
	long i = 0;
	char buf[RECV_BUF_SIZE];
	char ret;
	ret = parse_word(instr, buf, DELIMITER, &i);	dim_x_ = atoi(buf);
	ret = parse_word(instr, buf, DELIMITER, &i);	strcpy(flame_id_, buf);
	ret = parse_word(instr, buf, DELIMITER, &i);	strcpy(flame_id2_, buf);
	ret = parse_word(instr, buf, DELIMITER, &i);	strcpy(flame_id3_, buf);
	ret = parse_word(instr, buf, DELIMITER, &i);	x_[0] = atof(buf);
	
	//printf("dim_x_ = %d\n", dim_x_);
	//printf("flame_id_ = %s\n", flame_id_);
	//printf("flame_id2_ = %s\n", flame_id2_);
	//printf("flame_id3_ = %s\n", flame_id3_);
	//printf("x_[0] = %lf\n", x_[0]);
}

void move_ptr_to_null(char** p)
{
	while ( **p != '\0' )
	{
		(*p)++;
	};
}

void create_send_msg( char* outstr )
{
	char* p = outstr;
	sprintf(p, "%d,%s,%s", dim_x_, flame_id_, flame_id2_);
	move_ptr_to_null(&p);
	sprintf(p, ",%s,%lf", flame_id3_, x_[0]);
	//printf("%s\n", outstr);
}
void timer_callback( void )
{
}

int main()
{
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
		printf("recv : %s\n", recv_buf);
		
		/* parse receive message */
		parse_receive_msg(recv_buf);
		
		timer_callback();
		
		/* create send message */
		create_send_msg(send_buf);
		
		/* send */
		send(connect, send_buf, SEND_BUF_SIZE, 0);
		printf("send : %s\n", send_buf);
		
		printf("\n");
		
		close(connect);
	}
	
	close(sockfd);
	
	return 0;
}
