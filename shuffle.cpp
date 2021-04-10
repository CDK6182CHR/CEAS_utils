/*
 * This program shuffles list in text file specified by `FILENAME` and output with number.
 * To compile this program, C++11 or newer standard is required. 
 * An example compile command:
 * $ g++ shuffle.cpp -o shuffle.exe -std=c++11
 */
#include<iostream>
#include<fstream>
#include<algorithm>
#include<vector>
#include<string>
#include<ctime>
using namespace std;

#define FILENAME "data.txt"
#define OUTPUT "out.txt"

int main()
{
	ifstream file(FILENAME, ios::in);
	if (!file.is_open()) {
		cerr << "FETAL: data file not found, existing..." << endl;
		cerr << "HINT: Put items to be chosen in file `" << FILENAME << "` and run this program, " <<
			"or see source code." << endl;
		system("pause");
		exit(1);
	}
	vector<string> items;
	char buffer[1024];
	while (!file.eof()) {
		file.getline(buffer, 1024);
		string&& s(buffer);
		if (!s.empty()) {
			items.push_back(s);
		}
	}
	file.close();
	ofstream fout(OUTPUT, ios::out);
	for (int i = 0; i < 20; i++) {
		shuffle(items.begin(), items.end(), default_random_engine(time(nullptr)));
	}
	for (int i = 0; i < items.size(); i++) {
		cout << i + 1 << '\t' << items[i] << endl;
		fout << i + 1 << '\t' << items[i] << endl;
	}
	fout.close();
	cout << "===================== end of list =====================" << endl;
	cout << "Hint: this program shuffles input list in file " << FILENAME << " and print "
		<< "to stdout. The result is also saved in file " << OUTPUT << endl;
	system("pause");
}