#include <iostream>
#include <string>
#include <vector>

typedef struct DATA_STRUCT
{
	int a;
	std::string b;
	std::string c;
	std::string d;
	double e;
} T_DATA_STRUCT;

void parseStr(
	std::string& text,
	std::string& delimiter,
	std::vector<std::string>& words,
	DATA_STRUCT& data
)
{
	std::cout << "=== parse str ===" << std::endl;
	std::size_t pos;
	while ((pos = text.find(delimiter)) != std::string::npos) {
		words.push_back(text.substr(0, pos));
		text.erase(0, pos + delimiter.length());
	}
	words.push_back(text);
	
	unsigned int idx = 0;
	data.a = std::stoi(words[idx]); idx++;
	data.b = words[idx]; idx++;
	data.c = words[idx]; idx++;
	data.d = words[idx]; idx++;
	data.e = std::stod(words[idx]); idx++;
	
	std::cout << data.a << std::endl;
	std::cout << data.b << std::endl;
	std::cout << data.c << std::endl;
	std::cout << data.d << std::endl;
	std::cout << data.e << std::endl;
	
	std::cout << std::endl;
}

void createStr(
	DATA_STRUCT& data
)
{
	std::cout << "=== create str ===" << std::endl;

	std::string outstr;
	outstr += std::to_string(data.a) + ',';
	outstr += data.b + ',';
	outstr += data.c + ',';
	outstr += data.d + ',';
	outstr += std::to_string(data.e);
	std::cout << outstr << std::endl;
	std::cout << std::endl;
}

int main()
{
	std::string text = "1,base,sadfas,saf,1e-5";
	std::string delimiter = ",";
	std::vector<std::string> words{};
	DATA_STRUCT data;
	
	std::cout << text << std::endl;
	std::cout << std::endl;
	
	parseStr(text, delimiter, words, data);
	createStr(data);
}
