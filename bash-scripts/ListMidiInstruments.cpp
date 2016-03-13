/**
 * find midi-instrument (numbers) 
 * 
 * output: 
 *   channels
 *   midi-instrument numbers
 */
#include <iostream>
#include <fstream>
#include <sstream>
using namespace std;

ifstream::pos_type size;
char *memblock;

int main(int argc, char **argv) {
  if (argc != 2) {
    cout << "Usage: " << argv[0] << " MidiFile" << endl;
    return 1;
  }

  ifstream file(argv[1], ios::in | ios::binary | ios::ate);

  if (file.is_open()) {
    size = file.tellg();
    memblock = new char[size];
    file.seekg(0, ios::beg);
    file.read(memblock, size);
    file.close();

    int currChannel = 0;
    unsigned int pattern = 0xC0;

    for (unsigned char c = 0; c < 16; ++c) {
      for (int i = 0; i < size; ++i) {
        unsigned int haufen = static_cast<unsigned char>(memblock[i]);

        if (haufen == pattern) {
          cout << "Channel: " << currChannel << endl;

          unsigned int midiInstrument;

          midiInstrument = static_cast<unsigned char>(memblock[i + 1]);

          cout << "Instrument: " << midiInstrument << endl;
        }
      }

      currChannel++;
      pattern++;
    }

  } else {
    cout << "file open error" << endl;
    return 1;
  }

  return 0;
}
