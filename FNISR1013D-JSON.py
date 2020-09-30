#!/usr/bin/python
# FNISR1013D-JSON.py
# transforms wav  waveform captures from FNISR1013D scope to Json
#
# Author : Bob Tidey
# Date   : 30/09/2020
import time
import array
import json

# -----------------------
# Main Script
# -----------------------
# set json_indent to None for smallest file
json_indent = 1

voltList = [[5.0,"V",1],[2.5,"V",1],[1.0,"V",1],[500,"mV",0.001],[200,"mV",0.001],[100,"mV",0.001],[50,"mV",0.001]]
timeList = [[50,"S",1],[20,"S",1],[10,"S",1],[5,"S",1],[2,"S",1],[1,"S",1],[500,"mS",.001],[200,"mS",.001],[100,"mS",.001],[50,"mS",.001],[20,"mS",.001],[10,"mS",.001],[5,"mS",.001],[2,"mS",.001],[1,"mS",.001],[500,"uS",1E-6],[200,"uS",1E-6],[100,"uS",1E-6],[50,"uS",1E-6],[20,"uS",1E-6],[10,"uS",1E-6],[5,"uS",1E-6],[2,"uS",1E-6],[1,"uS",1E-6],[500,"nS",1E-9],[200,"nS",1E-9],[100,"nS",1E-9],[50,"nS",1E-9],[20,"nS",1E-9],[10,"nS",1E-9]]
measureList = ["Vpp","Vrms","Freq","Time+","Time-","Cycle","Vavg","Vmax","Vmin","Vp","Duty+","Duty-"]
header = [bytes(208)]
measures = [bytes(48), bytes(48)]
dataBuff = [bytes(3000), bytes(3000)]
dataScreen = [bytes(1500), bytes(1500)]
voltScale = [[5.0,"V",1],[2.5,"V",1]]
voltProbe = [1,1]
voltCoupling = ["DC","DC"]
timeScale = [50,"S",1.0]

jsObj = {
  "voltage": {"volts":[50,50], "units":["mV/div","mv/div"], "multiplier":[0.001,0.001], "probe":[1,1], "coupling":["DC","DC"]},
  "timebase": {"time":50, "units":"mS/div", "multiplier": 0.001},
  "trigger": {"channel":0,"edge":0,"type":0},
  "settings": {"screenBright":0,"gridBright":0,"scroll":0,"trig50":0},
  "measures": {"Vpp":[0,0], "Vrms":[0,0], "Freq":[0,0], "Time+":[0,0], "Time-":[0,0], "Cycle":[0,0], "Vavg":[0,0], "Vmax":[0,0], "Vmin":[0,0], "Vp":[0,0], "Duty+":[0,0], "Duty-":[0,0]},
  "dataBuffer": [
    {"channel" :"CH1", "units":"mV", "values" : [0 for i in range(1500)]},
    {"channel" :"CH2", "units":"mV", "values" : [0 for i in range(1500)]},
  ],
  "dataScreen": [
    {"channel" :"CH1", "units":"mV", "values" : [0 for i in range(750)]},
    {"channel" :"CH2", "units":"mV", "values" : [0 for i in range(750)]},
  ],
}


def getBinaryData(filename):
	f = open(filename, "rb")
	f.seek(0)
	header[0] = f.read(208)
	f.seek(208)
	measures[0] = f.read(48)
	f.seek(256)
	measures[1] = f.read(48)
	f.seek(1000)
	dataBuff[0] = f.read(3000) 
	f.seek(4000)
	dataBuff[1] = f.read(3000) 
	f.seek(7000)
	dataScreen[0] = f.read(1500) 
	f.seek(8500)
	dataScreen[1] = f.read(1500) 
	f.close()

def saveAsJson(filename, ctl):
	f =open(filename, "w")
	f.write(json.dumps(jsObj, indent=ctl))
	f.close()

def parseHeader():
	for x in range(2):
		voltScale[x] = voltList[header[0][4 + x*10]]
		voltProbe[x] = [1,10,100][header[0][10 + x*10]]
		voltCoupling[x] = ["DC", "AC"][header[0][8 + x*10]]
		jsObj["voltage"]["volts"][x] = voltScale[x][0]
		jsObj["voltage"]["units"][x] = voltScale[x][1]
		jsObj["voltage"]["multiplier"][x] = voltScale[x][2]
		jsObj["voltage"]["probe"][x] = voltProbe[x]
		jsObj["voltage"]["coupling"][x] = voltScale[x]
	timeScale = timeList[header[0][22]]
	jsObj["timebase"]["time"] = timeScale[0]
	jsObj["timebase"]["units"] = timeScale[1]
	jsObj["timebase"]["multiplier"] = timeScale[2]
	jsObj["trigger"]["channel"] = ["CH1","CH2"][header[0][30]]
	jsObj["trigger"]["edge"] = ["rising","falling"][header[0][28]]
	jsObj["trigger"]["type"] = ["auto","single","normal"][header[0][26]]
	jsObj["settings"]["screenBright"] = [header[0][120]]
	jsObj["settings"]["gridBright"] = [header[0][122]]
	jsObj["settings"]["scroll"] = ["fast","slow"][header[0][24]]
	jsObj["settings"]["trig50"] = ["Off","On"][header[0][124]]

def getMeasure(ch,mIndex):
	ad = mIndex * 4
	mt = measures[ch][ad]
	mv = (measures[ch][ad+1] * 256 + measures[ch][ad+2]) * 256 + measures[ch][ad+2]
	return str(mt) + "%" + str(mv)

def parseMeasures():
	for x in range(2):
		for y in range(12):
			jsObj["measures"][measureList[y]][x] = getMeasure(x,y)

def parseData():
	for x in range(2):
		jsObj["dataBuffer"][x]["units"] = voltScale[x][1]
		for y in range(0,1500):
			jsObj["dataBuffer"][x]["values"][y] = (dataBuff[x][y*2] + 256 * dataBuff[x][y*2+1] - 200) * voltScale[x][0] / 50
		jsObj["dataScreen"][x]["units"] = voltScale[x][1]
		for y in range(0,750):
			jsObj["dataScreen"][x]["values"][y] = dataScreen[x][y*2] + 256 * dataBuff[x][y*2+1]

# Main routine
filename = input('filename:')
getBinaryData(filename + '.wav')
parseHeader()
parseMeasures()
parseData()
saveAsJson(filename + '.json', json_indent)
print("finished")
