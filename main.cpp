
#include <iostream>
#include <sstream>
#include <string>
#include <vector>
#include <map>
#include <set>
#include <algorithm>
#include <locale>
#include <filesystem>


#include "pugixml-1.14/src/pugixml.hpp"
#include "rapidcsv/src/rapidcsv.h"
#include "OpenXLSX/OpenXLSX/OpenXLSX.hpp"

using namespace pugi;
using namespace OpenXLSX;
using std::string, std::vector, std::map, std::pair, std::endl, std::cout, std::cin, std::stringstream, std::move, std::to_string, std::cerr, std::sort, std::unordered_map, std::set, std::stoi;

rapidcsv::Document doc("bratsk-file.csv");

const size_t MAX_KVITANC_PER_FILE = 5500;

string getVidcity(const string& kvitanc_content);

string getNaspunkt(const string& kvitanc_content);

string getVidstreet(const string& kvitanc_content);

string getUlica(const string& kvitanc_content);

string getDom(const string& kvitanc_content);

string getKorpus(const string& kvitanc_content);

string getFlat(const string& kvitanc_content);

void importCSV(const string& fileName, unordered_map<string, vector<string>>& output);

bool addDocumentXML(map<string, xml_document>& Documents, const string& nameFile);

bool removeDocumentXML(map<string, xml_document>& Documents, const string& nameFile);

void printAllDocuments(const map<string, xml_document>& Documents);

void addKvitances(const map<string, pugi::xml_document>& Documents, map<string, vector<string>>& Kvitances, const unordered_map<string, vector<string>>& dataFromCSV, set<string>& checkUniq);

string getCourier(const unordered_map<string, vector<string>>& data, const vector<string>& values);

vector<string> formVector(const string& kvitanc_content);

void sortKvitances(map<string, vector<string>>& Kvitances);

void createXMLFiles(const map<string, vector<string>>& kvitances);

string extractBeforeDelimiter(const string& input);

bool safeStoi(const string& str, int& result);

string extractBeforeDelimiterForFlat(const string& input);

void addAllXMLFilesFromDirectory(map<string, pugi::xml_document>& Documents, const string& directory);

void writeResultToXlsx(const std::map<std::string, std::vector<std::string>>& Kvitances);

void writeKvitancesToXlsx(const map<string, vector<string>>& Kvitances);


int main()
{
    int private_sector {0};
    int variant {0};
    bool flag {true};
    cout << "Привет, выбери распределять по частному сектору да - 1, нет - 0 \n";
    cin >> private_sector;

    map<string, pugi::xml_document> Documents;
    map<string, vector<string>> Kvitances;
    unordered_map<string, vector<string>> dataFromCSV;
    dataFromCSV.reserve(60000);
    set<string> checkUniq;
    

    importCSV("bratsk-file.csv", dataFromCSV);

    while(flag)
    {
        cout <<"\n1 - add xml \
                \n2 - remove xml \
                \n3 - print all xml docs \
                \n4 - add kvitances from docs \
                \n5 - watch amount of kvitances \
                \n6 - sorting Kvitances\
                \n7 - create xml\
                \n8 - create xlsx files\
                ";
        cin >> variant;

        switch(variant)
        {
            case 1: {
                string nameDir {"example.xml"};
                cout << "enter name of dir with .xml \n";
                cin >> nameDir;

                addAllXMLFilesFromDirectory(Documents, nameDir);

                break;
            }
            case 2: {
                string nameFile;
                cout<<"enter name of file \n";
                cin>>nameFile;
                if(removeDocumentXML(Documents, nameFile)) {
                    cout<< "file was removed \n";
                }
                else {
                    cout<< "file was not removed \n";
                }

                break;
            }
            case 3: {
                cout << Documents.size() << " - количество документов\n";
                printAllDocuments(Documents);
                
                break;
            }

            case 4: {
                addKvitances(move(Documents), Kvitances, move(dataFromCSV), checkUniq);
                
                break;
            }
            case 5: {
                cout<<"amount of couriers - "<< Kvitances.size()<<endl;
                for(const auto& cour : Kvitances) {
                    cout<< cour.first<< " - courier \n";
                    cout<< cour.second.size()<< " - size \n"<<endl<<endl;
                }

                break;
            }
            case 6: {
                sortKvitances(Kvitances);
                cout<<"sorting is over \n";

                break;
            }
            case 7: {
                createXMLFiles(Kvitances);

                break;
            }
            case 8: {
                writeResultToXlsx(Kvitances);
                writeKvitancesToXlsx(Kvitances);
                
                break;
            }

            case 9: {
                
            }

            default:
                flag = false;
                break;
        }
    }


    return 0;

}





string getVidcity(const string& kvitanc_content) {

    size_t startIndex = kvitanc_content.find("VIDCITY=");
    if (startIndex != string::npos) {
        startIndex += 9; // 9 символов в VIDCITY="
        size_t endIndex = kvitanc_content.find('"', startIndex);
        if (endIndex != string::npos)
            return kvitanc_content.substr(startIndex, endIndex - startIndex);
    }
    return "";

}

string getNaspunkt(const string& kvitanc_content) {
    
    size_t startIndex = kvitanc_content.find("NASPUNKT=");
    if (startIndex != string::npos) {
        startIndex += 10; // 10 символов в NASPUNKT="
        size_t endIndex = kvitanc_content.find('"', startIndex);
        if (endIndex != string::npos)
            return kvitanc_content.substr(startIndex, endIndex - startIndex);
    }
    return "";

}

string getVidstreet(const string& kvitanc_content){

    size_t startIndex = kvitanc_content.find("VIDSTREET=");
    if (startIndex != string::npos) {
        startIndex += 11; // 11 символов в VIDSTREET="
        size_t endIndex = kvitanc_content.find('"', startIndex);
        if (endIndex != string::npos)
            return kvitanc_content.substr(startIndex, endIndex - startIndex);
    }
    return "";

}

string getUlica(const string& kvitanc_content) {

    size_t startIndex = kvitanc_content.find("ULICA=");
    if (startIndex != string::npos) {
        startIndex += 7; // 7 символов в ULICA="
        size_t endIndex = kvitanc_content.find('"', startIndex);
        if (endIndex != string::npos)
            return kvitanc_content.substr(startIndex, endIndex - startIndex);
    }
    return "";

}

string getDom(const string& kvitanc_content) {

    size_t startIndex = kvitanc_content.find("DOM=");
    if (startIndex != string::npos) {
        startIndex += 5; // 5 символов в DOM="
        size_t endIndex = kvitanc_content.find('"', startIndex);
        if (endIndex != string::npos)
            return kvitanc_content.substr(startIndex, endIndex - startIndex);
    }
    return "";

}

string getKorpus(const string& kvitanc_content) {

    size_t startIndex = kvitanc_content.find("KORPUS=");
    if (startIndex != string::npos) {
        startIndex += 8; // 8 символов в KORPUS="
        size_t endIndex = kvitanc_content.find('"', startIndex);
        if (endIndex != string::npos)
            return kvitanc_content.substr(startIndex, endIndex - startIndex);   
    }
    return "";

}

string getFlat(const string& kvitanc_content) {

    size_t startIndex = kvitanc_content.find("FLAT=");
    if (startIndex != string::npos) {
        startIndex += 6; // 6 символов в FLAT="
        size_t endIndex = kvitanc_content.find('"', startIndex);
        if (endIndex != string::npos)
            return kvitanc_content.substr(startIndex, endIndex - startIndex);   
    }
    return "";

}

void importCSV(const string& fileName, unordered_map<string, vector<string>>& output) {

    vector<string> column;
    string columnName;

    for(size_t col = 0; col < doc.GetColumnCount(); ++col) {
        column = doc.GetColumn<string>(col);
        columnName = doc.GetColumnName(col);
        output.insert({columnName, column});
    }
}


bool addDocumentXML(map<string, xml_document>& Documents, const string& nameFile) {
    xml_document doc;
    xml_parse_result result = doc.load_file(nameFile.c_str());
    if (result) {
        Documents.emplace(nameFile, move(doc));
        cout<<nameFile << " is good \n";
        return true;
    } else {
        cerr << "Failed to load file: " << nameFile << endl;
        cerr << "Error description: " << result.description() << endl;
        cerr << "Error offset: " << result.offset << endl; 
        return false;
    }
}

void addAllXMLFilesFromDirectory(map<string, xml_document>& Documents, const string& directory) {
    for (const auto& entry : std::filesystem::directory_iterator(directory)) {
        if (entry.is_regular_file() && entry.path().extension() == ".xml") {
            addDocumentXML(Documents, entry.path().string());
        }
    }
}

bool removeDocumentXML(map<string, xml_document>& Documents, const string& nameFile) {
    auto it = Documents.find(nameFile);
    if (it != Documents.end()) {
        Documents.erase(it);
        return true;
    }
    else {
        return false;
    }
}

void printAllDocuments(const map<string, pugi::xml_document>& Documents) {
    for (const auto& pair : Documents) {
        cout << "Document name: " << pair.first << endl;
    }
}

void addKvitances(const map<string, xml_document>& Documents, map<string, vector<string>>& Kvitances, const unordered_map<string, vector<string>>& dataFromCSV, set<string>& checkUniq)
{
    for(const auto& elem : Documents)
    { 
        stringstream ss;

        xml_node root = elem.second.child("asrn"); 
        for (xml_node kvitanc : root.children("Kvitanc")) {
            
            
            kvitanc.print(ss, "", pugi::format_raw);
            
            checkUniq.emplace(ss.str());

            ss.str("");
            ss.clear();   
        }
    }

    Kvitances.clear();

    string courierName;
    int i =0;
    for(const auto& kvit : checkUniq) {
        courierName = getCourier(move(dataFromCSV), formVector(move(kvit)));
        Kvitances[courierName].push_back(move(kvit)); 
        ++i;
            if(i%1000 == 0){
                cout<<"i - "<<i<<endl;
            }
    }
}

size_t findValueInBucket(const unordered_map<string, vector<string>>& data, const string& key, const string& value) {
    // Проверяем, существует ли ключ
    auto it = data.find(key);
    if (it != data.end()) {
        
        const vector<string>& values = it->second;
        
        
        auto valIt = find(values.begin(), values.end(), value);
        
        if (valIt != values.end()) {
            return distance(values.begin(), valIt);
        }
    }
    return -1; 
}


string getCourier(const unordered_map<string, vector<string>>& data,const vector<string>& values) {
    size_t answerRow = doc.GetRowCount(); 

    
    size_t temp = findValueInBucket(data, "ТипНП", values[0]);
    if(temp >= 0 and temp < doc.GetRowCount()) {
        answerRow = temp; 

        temp = findValueInBucket(data, "НаселенныйПункт", values[1]);
        if(temp >= 0 and temp < doc.GetRowCount()) {
            answerRow = temp;

            temp = findValueInBucket(data, "Улица", values[2]);
            if(temp >= 0 and temp < doc.GetRowCount()) {
                answerRow = temp; 

                for (size_t startRow = answerRow; startRow < doc.GetRowCount(); ++startRow) {
        
                    if (extractBeforeDelimiter(data.at("Дом")[startRow])== values[3]){
                        answerRow = startRow;
                        break; 
                        
                    }
                }
            } 
        }

    }

     

    if (answerRow < doc.GetRowCount()) {
        return doc.GetCell<string>(9, answerRow);
    }

    return "нет_курьера"; 
}

vector<string> formVector(const string& kvitanc_content){

    vector<string> values{
        getVidcity(kvitanc_content), 
        getNaspunkt(kvitanc_content),
        getUlica(kvitanc_content), 
        getDom(kvitanc_content) };


    return values;
}



bool compare(const string& first, const string& second) {
    if (getVidcity(first) != getVidcity(second))
        return getVidcity(first) < getVidcity(second);
    if (getNaspunkt(first) != getNaspunkt(second))
        return getNaspunkt(first) < getNaspunkt(second);
    if (getVidstreet(first) != getVidstreet(second))
        return getVidstreet(first) < getVidstreet(second);
    if (getUlica(first) != getUlica(second))
        return getUlica(first) < getUlica(second);
    
    int dom1, dom2;
    if (safeStoi(getDom(first), dom1) && safeStoi(getDom(second), dom2)) {
        if (dom1 != dom2)
            return dom1 < dom2;
    }

    if (getKorpus(first) != getKorpus(second))
        return getKorpus(first) < getKorpus(second);
    
    int flat1, flat2;
    if (safeStoi(extractBeforeDelimiterForFlat(getFlat(first)), flat1) && safeStoi(extractBeforeDelimiterForFlat(getFlat(second)), flat2)) {
        return flat1 < flat2;
    }

    return false; 
}

void sortKvitances(map<string, vector<string>>& Kvitances) {
    for(auto& cour : Kvitances) {
        sort(cour.second.begin(), cour.second.end(), compare);
    }
}


void createXMLFiles(const map<string, vector<string>>& kvitances) {

    std::filesystem::create_directory("papka");

    for (const auto& kv : kvitances) {
        const string& courier_name = kv.first;
        const vector<string>& kvitanc_list = kv.second;
        size_t file_count = 0;
        size_t kvitanc_count = 0;

        while (kvitanc_count < kvitanc_list.size()) {
            xml_document doc;
            xml_node root = doc.append_child("Root");

            for (size_t i = 0; i < MAX_KVITANC_PER_FILE && kvitanc_count < kvitanc_list.size(); ++i, ++kvitanc_count) {
                root.append_buffer(kvitanc_list[kvitanc_count].c_str(), kvitanc_list[kvitanc_count].size(), parse_default | parse_fragment);
            }

            string filename = "papka/" + courier_name + (file_count > 0 ? " (" + to_string(file_count) + ")" : "") + ".xml";
            if (!doc.save_file(filename.c_str())) {
                cerr << "Failed to save file: " << filename << endl;
            } else {
                cout << "File saved: " << filename << endl;
            }
            ++file_count;
        }
    }
}

bool isAlpha(char ch) {
    string alphabet = "АБВГДЕЖЗИК";
    
    for(const auto& elem : alphabet) {
        if(elem == ch)  
            return true;
    }
    return false;
}

string extractBeforeDelimiter(const string& input) {
    size_t pos = 0;
    while (pos < input.size()) {
        if (input[pos] == '/' || input[pos] == '-' || isAlpha(input[pos])) {
            break;
        }
        ++pos;
    }
    return input.substr(0, pos);
}

string extractBeforeDelimiterForFlat(const string& input) {
    size_t pos = 0;
    while (pos < input.size()) {
        if (input[pos] == '/') {
            break;
        }
        ++pos;
    }
    return input.substr(0, pos);
}


bool isValidInteger(const string& str) {
    if (str.empty()) return false;
    for (char ch : str) {
        if (!isdigit(ch)) return false;
    }
    return true;
}

bool safeStoi(const string& str, int& result) {
    if (isValidInteger(str)) {
        try {
            result = stoi(str);
            return true;
        } catch (const std::invalid_argument&) {
            return false;
        } catch (const std::out_of_range&) {
            return false;
        }
    }
    return false;
}

void writeResultToXlsx(const map<string, vector<string>>& Kvitances) {

    std::filesystem::create_directory("XLSX_output");
    
    try {

        XLDocument doc;
        string filePath = "XLSX_output/result.xlsx";
        doc.create(filePath);
        auto sheet = doc.workbook().worksheet("Sheet1");

        sheet.cell("A1").value() = "Курьеры";
        sheet.cell("B1").value() = "Кол-во квт";
        doc.save();


        size_t row = 2; // Начинаем с 2-й строки, так как 1-я строка занята заголовками
        int sum = 0;
        for (const auto& kv : Kvitances) {
            sheet.cell("A" + to_string(row)).value() = kv.first;
            sheet.cell("B" + to_string(row)).value() = kv.second.size();
            sum += kv.second.size();
            ++row;
        }
        sheet.cell("A" + to_string(row)).value() = "Итого";
        sheet.cell("B" + to_string(row)).value() = sum;

        doc.save();
        doc.close();

        cout << "XLSX file created successfully!" << std::endl;
    } catch (const std::exception& e) {
        cerr << "Error: " << e.what() << std::endl;
    }
}

void writeKvitancesToXlsx(const map<string, vector<string>>& Kvitances) {

    std::filesystem::create_directory("XLSX_output");

    for (const auto& cour : Kvitances) {
        try {
            XLDocument doc;
            // Изменяем путь для создания файла в папке XLSX_output
            string filePath = "XLSX_output/" + cour.first + ".xlsx";
            doc.create(filePath);
            auto sheet = doc.workbook().worksheet("Sheet1");

            sheet.cell("A1").value() = "ТипНП";
            sheet.cell("B1").value() = "НаселенныйПункт";
            sheet.cell("C1").value() = "ТипУлицы";
            sheet.cell("D1").value() = "Улица";
            sheet.cell("E1").value() = "Дом";
            sheet.cell("F1").value() = "Квартира";

            size_t row = 2; // Начинаем с 2-й строки, так как 1-я строка занята заголовками

            for (const auto& kv : cour.second) {
                sheet.cell("A" + to_string(row)).value() = getVidcity(kv);
                sheet.cell("B" + to_string(row)).value() = getNaspunkt(kv);
                sheet.cell("C" + to_string(row)).value() = getVidstreet(kv);
                sheet.cell("D" + to_string(row)).value() = getUlica(kv);
                sheet.cell("E" + to_string(row)).value() = getDom(kv);
                sheet.cell("F" + to_string(row)).value() = getFlat(kv);

                ++row;
            }

            doc.save();
            doc.close();

            cout << "XLSX file created successfully: " << filePath << endl;
        } catch (const std::exception& e) {
            cerr << "Error: " << e.what() << endl;
        }
    }
}

