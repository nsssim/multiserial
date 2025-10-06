# Makefile for Multi Serial Monitor

CC = g++
CFLAGS = -mwindows
LIBS = -lcomctl32 -lriched20 -lgdi32 -luser32 -lkernel32 -ladvapi32
TARGET = SerialMonitor.exe
RESOURCE = SerialMonitor_res.o

all: $(TARGET)

$(RESOURCE): SerialMonitor.rc
	windres SerialMonitor.rc -o $(RESOURCE)

$(TARGET): main.cpp $(RESOURCE)
	$(CC) -o $(TARGET) main.cpp $(RESOURCE) $(CFLAGS) $(LIBS)

clean:
	del /Q $(RESOURCE) $(TARGET) 2>nul || true

.PHONY: all clean
