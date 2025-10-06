windres SerialMonitor.rc -o SerialMonitor_res.o

g++ -o SerialMonitor.exe main.cpp SerialMonitor_res.o -mwindows -lcomctl32 -lriched20 -lgdi32 -luser32 -lkernel32 -ladvapi32
