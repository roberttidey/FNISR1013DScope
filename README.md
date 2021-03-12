# FNISR1013D scope waveform decode

Work in progress.

This is a set of files helping an investigation decoding the waveform (.wav) files from a FNISR1013D tablet oscilloscope.
These are 15000 byte binary files containing, bufffer data, display data and settings when a capture is made.

The aim is to produce a conversion tool that can make this available in a readable and computable form.

The vbscript file scopedump.vbs is an evolving script that is picking apart the wav file as the investigation proceeds. To use run and enter full filename like 1.wav

The python program FNISR1013D-JSON.py is the target conversion program. To use run it and enter the root wav filename like 1. It produces a JSON file with settings plus buffer and screen data. It seems to produce valid results except for the measures section as the mapping and encoding is still under investigation.

wavoigt has used this data to produce a nice viewer running in Excel

https://github.com/wavoigt/FNIRSI-1013D-WAV-Viewer-in-Excel-VBA
