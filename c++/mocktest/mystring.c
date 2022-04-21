void mystrcpy(char* outstr, const char* instr)
{
	long i = 0;
	while ( instr[i] != '\0' )
	{
		outstr[i] = instr[i];
		i++;
	}
	outstr[i] = '\0';
}
void mymemset(char* outstr, char ch, long size)
{
	for ( long i = 0; i < size; i++ )
	{
		outstr[i] = ch;
	}
}
int myatoi(const char* str) {
	int num = 0;
	int type = 0;
	
	// 先頭に+付いてたら無視する
	if ( *str == '+' ) {
		str++;
	}
	// 先頭に-付いてたらtypeフラグを立てておく
	else if ( *str == '-' ) {
		type = 1;
		str++;
	}
	
	while(*str != '\0'){
		// 0〜9以外の文字列ならそこで終了
		if ( *str < 48 || *str > 57 ) {
			break;
		}
		num += *str - 48;
		num *= 10;
		str++;
	}
	
	num /= 10;
	
	// -符号が付いていたら0から引くことで負の値に変換する
	if ( type ) {
		num = 0 - num;
	}
	
	return num;
}
double myatof(const char* str)
{
	//TODO:実装
}

char parse_word( const char* instr, char* buf, char delimiter, long* idx )
{
	long bufi = 0;
	char finished = 0;
	char reachtail = 0;
	while ( finished == 0 )
	{
		if ( instr[*idx] == '\0' )
		{
			buf[bufi] = '\0';
			finished = 1;
			reachtail = 1;
		}
		else if ( instr[*idx] == delimiter )
		{
			buf[bufi] = '\0';
			bufi = 0;
			finished = 1;
		}
		else
		{
			buf[bufi] = instr[*idx];
			bufi++;
		}
		(*idx)++;
	};
	return reachtail;
}


