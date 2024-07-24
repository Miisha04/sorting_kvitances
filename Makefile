# Имя исполняемого файла
TARGET = main

# Исходные файлы
SRCS = main.cpp pugixml-1.14/src/pugixml.cpp

# Директории для заголовочных файлов
INCLUDES = -I OpenXLSX/build/OpenXLSX/

# Директории для библиотек
LIBDIRS = -L OpenXLSX/build/output

# Библиотеки
LIBS = -lOpenXLSX

# Компилятор и флаги
CXX = g++
CXXFLAGS = $(INCLUDES)
LDFLAGS = $(LIBDIRS) $(LIBS)

# Правило для сборки всего проекта
$(TARGET): $(SRCS)
	$(CXX) -o $(TARGET) $(SRCS) $(CXXFLAGS) $(LDFLAGS)

# Правило для очистки
clean:
	rm -f $(TARGET)
