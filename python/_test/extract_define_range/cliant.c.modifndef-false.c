#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <sys/socket.h>
#include <sys/types.h>
#include <arpa/inet.h>
#include <unistd.h>

#define DBG (1)
#define MOD_IF (1)
#define MOD_IFDEF (1)
#define MOD_IFDEFIFNDEF (1)

const char* IN_VEC_PATH_BASE = "testdata/input_test_vec";
const char* RECV_FILE_PATH_BASE = "testdata/recv_data";
const unsigned int PORTNO = 1234;
const char* IPADDR = "127.0.0.1";
const unsigned int RECV_BUF_SIZE = 3000;
const unsigned int SEND_BUF_SIZE = 3000;
const unsigned int RECV_WORDS_NUM = 5;

char communicateTcp(
	const char* send_str,
	unsigned int send_size,
	char* recv_str
)
{
#if DBG
	printf("communicateTcp() called\n");
#else /* DBG */
#endif /* DBG */
	int sockfd = socket(AF_INET, SOCK_STREAM, 0);
	if(sockfd < 0) {
		printf("Error socket\n");
		return 1;
	}
	
#if MOD_IF
	struct sockaddr_in addr;
#else /* MOD_IF */
#endif /* MOD_IF */
#if MOD_IF
	memset(&addr, 0, sizeof(struct sockaddr_in));
	addr.sin_family = AF_INET;
	addr.sin_port = htons(PORTNO);
	addr.sin_addr.s_addr = inet_addr(IPADDR);
#endif /* MOD_IF */
	
#if MOD_IF
#else /* MOD_IF */
	connect(sockfd, (struct sockaddr *)&addr, sizeof(struct sockaddr_in));
#endif /* MOD_IF */
	
#if MOD_IF
	/* send */
#else /* MOD_IF */
	send(sockfd, send_str, send_size, 0);
#endif /* MOD_IF */
	printf("send : %s\n", send_str);
	
	
	
	close(sockfd);
	
	return 0;
}

char readInputVecFile(
	char* send_str,
	unsigned int fileidx
)
{
	FILE *fp;
	char ch;
	char invecpath[50];
	char* p = send_str;
	
#ifdef MOD_IFDEF
	/* open inputvecfile */
#else /* MOD_IFDEF */
	sprintf(invecpath, "%s%d" , IN_VEC_PATH_BASE, fileidx);
#endif /* MOD_IFDEF */
	fp = fopen(invecpath , "r");
	if (fp == NULL) {
		return 1;
	}
	
	/* combine multiple lines into one line */
	while(1) {
		if (ferror(fp)) {
			break;
		}
		ch = fgetc(fp);
		if (!feof(fp)) {
			if (ch == '\n') {
				*p = ',';
			} else {
				*p = ch;
			}
#ifdef MOD_IFDEF
			p++;
#endif /* MOD_IFDEF */
		} else {
			break;
		}
	}
#ifdef MOD_IFDEF
#else /* MOD_IFDEF */
	p--;
	*p = '\0';
#endif /* MOD_IFDEF */
	
#ifdef MOD_IFDEF
	/* close inputvecfile */
	fclose(fp);
#else /* MOD_IFDEF */
#endif /* MOD_IFDEF */
	
	return 0;
}

char writeRecvDataFile(
	const char* recv_str,
	unsigned int fileidx
)
{
	FILE *fp;
	char ch;
#ifdef MOD_IFDEFIFNDEF
#else /* MOD_IFDEFIFNDEF */
	char* p = (char*)recv_str;
#endif /* MOD_IFDEFIFNDEF */
	char recvvecpath[50];
#ifndef MOD_IFDEFIFNDEF
#else /* !MOD_IFDEFIFNDEF */
	char recv_words[RECV_WORDS_NUM][100];
#endif /* !MOD_IFDEFIFNDEF */
	
	/* open recvdatafile */
	p--;
	sprintf(recvvecpath, "%s%d" , RECV_FILE_PATH_BASE, fileidx);
	fp = fopen(recvvecpath , "w");
	if (fp == NULL) {
		return 1;
	}
	
	/* split receive messages with delimiter */
	for ( int wordsidx = 0; wordsidx < RECV_WORDS_NUM; wordsidx++ )
	{
		int charidx = 0;
		while (1) {
			if ( (*p == ',') || (*p == '\0') ) {
				recv_words[wordsidx][charidx] = '\0';
				p++;
				break;
			} else {
				recv_words[wordsidx][charidx] = *p;
				charidx++;
				p++;
			}
		};
	}
	
	/* output receive messages to recvdatafile */
	fprintf(fp, "%s,%s\n", recv_words[0], recv_words[1]);
	fprintf(fp, "%s,%s\n", recv_words[2], recv_words[3]);
	fprintf(fp, "%s\n", recv_words[4]);
	
	/* close recvdatafile */
	fclose(fp);
	
	return 0;
}

int main(int argc, char *argv[])
{
	char result;
	unsigned int fileidx = 0;
	char send_str[SEND_BUF_SIZE];
	char recv_str[RECV_BUF_SIZE];
	
	if (argc > 1) {
		fileidx = atoi(argv[1]);
	}
	
	while (1) {
		
		/* read invecfile */
		result = readInputVecFile(send_str, fileidx);
		if (result == 1) {
			break;
		}
		
		/* communicate tcp */
		memset(recv_str, '\0', RECV_BUF_SIZE);
		communicateTcp(send_str, strlen(send_str), recv_str);
		if (result == 1) {
			break;
		}
		
		/* write recvdatafile */
		result = writeRecvDataFile(recv_str, fileidx);
		if (result == 1) {
			break;
		}
		
		fileidx++;
	}
}
