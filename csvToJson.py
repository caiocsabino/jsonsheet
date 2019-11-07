import re
import sys
import os
import subprocess
import zipfile
import io
import base64
import json
import re

jsonBasicTypes = ["string", "bool", "uint", "int", "uint64", "float", "object"]
jsonTypesAsStrings = [];
jsonArrayTypesAsStrings = []
arraySeparator = ","
emptyLinesAllowed = True
tempSheetIdentifierPreamble = "__.TEMP_JSON_SHEET.__"
tempSheetIdentifierPostamble = "__.LOCATION.__"
compressedObjPreamble = ":__JSON_OBJ__:"
isRootObjectArray = True

class Sheet:
	def __init__(self, csvString):
		self.csvString = csvString
		self.setupFromCSV()

	def setupFromCSV(self):
		# print(self.csvString)
		rows = self.csvString.splitlines()

		self.totalRows = len(rows)
		self.totalCols = 1
		self.values = []

		for i in range(0,self.totalRows):
			self.values.append([])

			cols = rows[i].split("\t")

			if len(cols) > self.totalCols:
				self.totalCols = len(cols)

			for k in range(0,self.totalCols):
				if k < len(cols):
					self.values[i].append(cols[k])
				else:
					self.values[i].append("")

	def getSheetValues(self,startRow, startCol, totalRows, totalCols):

		values = []

		for i in range(0,totalRows):
			values.append([])

			for j in range(0,totalCols):
				values[i].append(self.values[startRow-1+i][startCol-1+j])

		return values

	def getRange(self, rows, cols):
		return self.getSheetValues(1,1, rows, cols)

	def getLastRow(self):
		return self.totalRows

	def getLastColumn(self):
		return self.totalCols

	def setValue(self, row, col, value):
		self.values[row-1][col-1] = value

	def setFormula(self, row, col, formula):
		print("NOT IMPLEMENTED formula")


for i in range(0,len(jsonBasicTypes)):
	jsonTypesAsStrings.append("(" + jsonBasicTypes[i] + ")")
	jsonTypesAsStrings.append("(" + jsonBasicTypes[i] + "s)")

	jsonArrayTypesAsStrings.append("(" + jsonBasicTypes[i] + "s)");

def isObjectType(type):
	return type == "(object)"

def isArrayType(type):
	return type in jsonArrayTypesAsStrings

def getArrayBasicType(arrayType):
	lastLetters = arrayType[-2:]

	if (lastLetters == "s)"):
		return arrayType[1:-2]

	return None

def isEmptyRowOrCol(index, useRow, sheet):
	totalRows = sheet.getLastRow()
	totalCols = sheet.getLastColumn()

	values = []
	end = -1

	if useRow:
		values = sheet.getSheetValues(index + 1, 1, 1, totalCols)
		end = totalCols
	else:
		values = sheet.getSheetValues(1, index + 1, totalRows, 1)
		end = totalRows

	for i in range(0,end):
		value = ""

		if useRow:
			value = values[0][i]
		else:
			value = values[i][0]

		if (value != ""):
			return False;

	return True;

def getCelTypeAndName(content):
	result = {};

	# TODO: Use regex to extract
	for i in range(0,len(jsonTypesAsStrings)):
		target = jsonTypesAsStrings[i];

		if target in content:
			result["type"] = target;
			result["name"] = content.replace(target, "");
			return result;


def isJSONType(givenType):
	typeLower = givenType.lower();

	for k in range(0,len(jsonTypesAsStrings)):
		targetLower = jsonTypesAsStrings[k].lower();

		if targetLower in typeLower:
			return True;

	return False;


def detectDirection(sheet):
	validRowsIndices = [];
	invalidRowsIndices = [];

	validColsIndices = [];
	invalidColsIndices = [];

	totalRows = sheet.getLastRow();
	totalCols = sheet.getLastColumn();
	values = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
	lineIndexWithFirstJsonType = -1;
	colIndexWithFirstJsonType = -1;

	for i in range(0,totalRows):
		for j in range(0,totalCols):
			entryValue = "" + values[i][j];
			entryValue = entryValue.lower();

			if (isJSONType(entryValue)):
				if (lineIndexWithFirstJsonType == -1):
					lineIndexWithFirstJsonType = i;
					colIndexWithFirstJsonType = j;

				typedCelsInSameRow = 1;
				typedCelsInSameCol = 1;

				for kk in range(i+1,totalRows):
					celValue = "" + values[kk][j];
					celValue = celValue.lower();
					if (isJSONType(celValue)):
						typedCelsInSameCol = typedCelsInSameCol + 1;
					else:
						break;

				for kk in range(j+1,totalCols):
					celValue = "" + values[i][kk];
					celValue = celValue.lower();
					if (isJSONType(celValue)):
						typedCelsInSameRow = typedCelsInSameRow + 1;
					else:
						break;

				return typedCelsInSameRow > typedCelsInSameCol;

def printObj(obj):
	if (obj != None):
		s = json.dumps(obj)

		print(s);

def getLastValidRowAndNonEmptyRow(_emptyLinesAllowed, sheet, directionIsHorizontal):
	returnObj = {};

	result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	valueToUse = "validRowsIndices"

	if not directionIsHorizontal:
		valueToUse = "validColsIndices"

	if (len(result[valueToUse]) > 0):
		firstValidRow = result["validRowsIndices"][0];
		firstValidCol = result["validColsIndices"][0];
		lastValidCol = result["validColsIndices"][len(result["validColsIndices"]) - 1];
		lastValidRow = result["validRowsIndices"][len(result["validRowsIndices"]) - 1];
		lastValidRowOrCol = -1;
		lastNonEmptyRowOrCol = -1;

		totalRows = sheet.getLastRow();
		totalCols = sheet.getLastColumn();

		values = sheet.getSheetValues(1, 1, totalRows, totalCols);

		startI = firstValidRow
		endI = totalRows

		if not directionIsHorizontal:
			startI = firstValidCol
			endI = totalCols

		for i in range(startI,endI):
			isEmptyLineOrCol = True;

			startJ = firstValidCol
			endJ = lastValidCol

			if not directionIsHorizontal:
				startJ = firstValidRow
				endJ = lastValidRow

			for j in range(startJ,endJ):	
				entryValue = "" + (values[i][j]);

				if not directionIsHorizontal:
					entryValue =  values[j][i]

				if (entryValue != ""):
					isEmptyLineOrCol = False;
					break;

			testRows = directionIsHorizontal and (len(result["invalidRowsIndices"]) > 0 and result["invalidRowsIndices"][0] == i and result["invalidColsIndices"][0] == j);
			testCols = not directionIsHorizontal and (len(result["invalidColsIndices"]) > 0 and result["invalidColsIndices"][0] == i and result["invalidRowsIndices"][0] == j);

			shouldEnd = (isEmptyLineOrCol and not _emptyLinesAllowed) or testRows or testCols;


			if (not shouldEnd):
				lastValidRowOrCol = i;
				if (not isEmptyLineOrCol):
					lastNonEmptyRowOrCol = i;
			else:
				break;

		if (directionIsHorizontal):
			returnObj["lastValidRow"] = lastValidRowOrCol;
			returnObj["lastNonEmptyRow"] = lastNonEmptyRowOrCol;
		else:
			returnObj["lastValidCol"] = lastValidRowOrCol;
			returnObj["lastNonEmptyCol"] = lastNonEmptyRowOrCol;

		return returnObj;

def isJsonString(str):
	try:
		json_object = json.loads(str)
	except ValueError as e:
		return False
	return True

def parseValueIntoObject(object, entryName, entryBasicType, value, sheet, row, col):
	if (entryBasicType == "(int)" or entryBasicType == "(uint)" or entryBasicType == "(int64)" or entryBasicType == "(uint64)"):
		object[entryName] = int(value);
	elif (entryBasicType == "(float)"):
		object[entryName] = parseFloat(value);
	elif (entryBasicType == "(bool)"):
		valueLower = value.toLowerCase();

		if valueLower == "1" or valueLower == "true":
			object[entryName] = True;
		else:
			object[entryName] =  False;
	elif (entryBasicType == "(string)"):
		object[entryName] = value;
	elif (entryBasicType == "(object)"):

		if (isJsonString(value)):
			object[entryName] = value;
		else:
			targetName = tempSheetIdentifierPreamble + sheet.getSheetId() + tempSheetIdentifierPostamble + row + "," + col;

			newSheet = deserializeSheet(targetName, value);

			directionIsHorizontal = detectDirection(newSheet);

			newObject = createObject(newSheet, entryName, directionIsHorizontal);

			if (newObject):
				object[entryName] = newObject;

				activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
				activeSpreadsheet.deleteSheet(newSheet);

			spreadsheet = SpreadsheetApp.getActive();
			spreadsheet.setActiveSheet(sheet);

			# by default all deserialized sheets are treated as arrays, so since here we know it is a single object we must create it from the first entry

			object[entryName] = newObject[entryName][0];

			#object[entryName] = value;
	else:
		print("WILL BREAK, unkown type " + entryBasicType);

def pushValueIntoArray(array, entryName, basicType, value, sheet, row, col):
	if (value == None):
		return;

	value = "" + value;

	hasSeparator = arraySeparator in value

	values = [];

	if (hasSeparator and basicType != "object"):
		values = value.split(arraySeparator);
	else:
		values.append(value);

	for i in range(0,len(values)):
		value = values[i];

	if (basicType == "int" or basicType == "uint" or basicType == "int64" or basicType == "uint64"):
		for i in range(0,len(values)):
			value = values[i];
			array.append(int(value));
	elif (basicType == "float"):
		for i in range(0,len(values)):
			value = values[i];
			array.append(parseFloat(value));
	elif (basicType == "bool"):
		for i in range(0,len(values)):
			value = values[i];
			valueLower = value.lower();

			if valueLower == "1" or valueLower == "true":
				array.append(True);
			else:
				array.append(False)
	elif (basicType == "string"):
		for i in range(0,len(values)):
			value = values[i];
			array.append(value);
	elif (basicType == "object"):
		if (isJsonString(value)):
			jsonObject = json.loads(value)

			for i in range(0,len(jsonObject)):
				array.append(jsonObject[i]);
		else:
			# targetName = tempSheetIdentifierPreamble + sheet.getSheetId() + tempSheetIdentifierPostamble + row + "," + col;
			targetName = "TEMPSHEET"

			newSheet = deserializeSheet(targetName, value);

			directionIsHorizontal = detectDirection(newSheet);

			newObject = createObject(newSheet, entryName, directionIsHorizontal, False);

			# if (newObject):
			# 	object[entryName] = newObject;

			# 	activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
			# 	activeSpreadsheet.deleteSheet(newSheet);

			# spreadsheet = SpreadsheetApp.getActive();
			# spreadsheet.setActiveSheet(sheet);

			for i in range(0,len(newObject[entryName])):
				array.append(newObject[entryName][i]);

def isAllArrays(sheet, directionIsHorizontal):
	result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	if (len(result["validRowsIndices"]) > 0):
		object = {}

		totalRows = sheet.getLastRow();
		totalCols = sheet.getLastColumn();

		values = sheet.getSheetValues(1, 1, totalRows, totalCols);

		firstValidRow = result["validRowsIndices"][0];
		firstValidCol = result["validColsIndices"][0];

		totalValidRows = len(result["validRowsIndices"]);
		totalValidCols = len(result["validColsIndices"]);

		marker = totalValidCols

		if not directionIsHorizontal:
			marker = totalValidRows

		for i in range(0,marker):
			index = result["validColsIndices"][i];

			if not directionIsHorizontal:
				index = result["validRowsIndices"][i]

			celTypeAndName = getCelTypeAndName(values[firstValidRow][index]);

			if not directionIsHorizontal:
				celTypeAndName = getCelTypeAndName(values[index][firstValidCol]);

			celType = celTypeAndName["type"];
			celName = celTypeAndName["name"];

			isArray = isArrayType(celType);

			if (not isArray):
				return False;
		return True;
	return False;


def deserializeSheet(sheetName, data):

	# insert data

	if (data != ""):
		if (compressedObjPreamble in data):
			data = data.replace(compressedObjPreamble, "");


		z = zipfile.ZipFile(io.BytesIO(base64.decodestring(data)))

		unzipped = z.read(z.infolist()[0])

		z.close()

		jsonObject = json.loads(unzipped)


		if (jsonObject != None and "cels" in jsonObject):
			totalRows = jsonObject["cels"]["totalRows"];
			totalCols = jsonObject["cels"]["totalCols"];

			values = jsonObject["cels"]["values"];
			formulas = jsonObject["cels"]["formulas"];
			bgs = jsonObject["cels"]["bgs"];

			tsvString = ""

			for i in range(0,totalRows):
				for j in range(0,totalCols):
					tsvString = tsvString + str(values[i*totalCols+j]) 
					if j < totalCols-1:
						tsvString = tsvString + "\t"

				if i < totalRows - 1:
					tsvString = tsvString + "\n"

			newSheet = Sheet(tsvString)

			return newSheet

	return None;

def createObject(sheet, name, directionIsHorizontal, isRoot):
	result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	emptyLineEndsArray = False;

	object = None;

	if (len(result["validRowsIndices"]) > 0):
		object = {}

		firstValidRow = result["validRowsIndices"][0];
		firstValidCol = result["validColsIndices"][0];
		lastValidCol = result["validColsIndices"][len(result["validColsIndices"]) - 1];
		lastValidRow = result["validRowsIndices"][len(result["validRowsIndices"]) - 1];

		lastValidRowAndNonEmptyRow = getLastValidRowAndNonEmptyRow(emptyLinesAllowed, sheet, directionIsHorizontal);

		if "lastNonEmptyRow" in lastValidRowAndNonEmptyRow:
			lastNonEmptyRow = lastValidRowAndNonEmptyRow["lastNonEmptyRow"];

		if "lastNonEmptyCol" in lastValidRowAndNonEmptyRow:
			lastNonEmptyCol = lastValidRowAndNonEmptyRow["lastNonEmptyCol"];

		object[name] = None
		if isRoot and not isRootObjectArray:
			object[name] = {};
		else:
			object[name] = [];

		totalRows = sheet.getLastRow();
		totalCols = sheet.getLastColumn();

		values = sheet.getSheetValues(1, 1, totalRows, totalCols);

		currentObject = None;

		allArrays = isAllArrays(sheet, directionIsHorizontal);

		currentObjectEmptyArrayEntriesFound = []

		start = firstValidRow + 1
		end = lastNonEmptyRow

		if not directionIsHorizontal:
			start = firstValidCol + 1
			end = lastNonEmptyCol

		for i in range(start,end+1):
			if (isEmptyRowOrCol(i, directionIsHorizontal, sheet) and not allArrays):
				continue;

			newObjectStarting = True;

			# detects if this line is starting a new object
			if (currentObject != None):
				#Browser.msgBox('Result', "has object " + " " + i + " "  + values[i][firstValidCol], Browser.Buttons.OK); 
				newObjectStarting = False;

				start2 = firstValidCol;
				end2 = lastValidCol;

				if not directionIsHorizontal:
					start2 = firstValidRow
					end2 = lastValidRow

				for j in range(start2,end2):
					content = values[firstValidRow][j];

					if not directionIsHorizontal:
						content = values[j][firstValidCol];

					if (content == ""):
						continue;

					celTypeAndName = getCelTypeAndName(content);
					celType = celTypeAndName["type"];
					celName = celTypeAndName["name"];
					isArray = isArrayType(celType);
					basicType = celType;
					value = values[i][j];

					if not directionIsHorizontal:
						value = values[j][i]

					isEmpty = value == "";
					objStr = json.dumps(currentObject);
					simpleEntryAlreadyInput = (not isArray and celName in currentObject and not isEmpty);
					arrayHadAlreadyEmptyEntryInAllArraySetup = (allArrays and j in currentObjectEmptyArrayEntriesFound and not isEmpty)

					if (simpleEntryAlreadyInput or arrayHadAlreadyEmptyEntryInAllArraySetup):
						#Browser.msgBox('Result', "WILL START NEW OBJECT " + " " + i + " "  + values[i][j] + " is array " + isArray + " " + currentObject[colName], Browser.Buttons.OK); 
						newObjectStarting = True;
						break;

			# end of detecting if a new object is been started

			if (newObjectStarting):
				newObjectToReplace = {};
				if (isRoot and not isRootObjectArray):
					if (currentObject == None):
						object[name] = newObjectToReplace;
					else:
						return object;
				else:
					object[name].append(newObjectToReplace);

				currentObject = newObjectToReplace;
				currentObjectEmptyArrayEntriesFound = [];

			start2 = firstValidCol;
			end2 = lastValidCol;

			if not directionIsHorizontal:
				start2 = firstValidRow
				end2 = lastValidRow

			for j in range(start2,end2+1):
				content = values[firstValidRow][j];

				if not directionIsHorizontal:
					content = values[j][firstValidCol]

				if (content == ""):
					continue;

				celTypeAndName = getCelTypeAndName(content);
				celType = celTypeAndName["type"];
				celName = celTypeAndName["name"];
				isArray = isArrayType(celType);
				basicType = celType;
				value = ""

				if directionIsHorizontal:
					value = value + values[i][j]
				else:
					value = value + values[j][i]

				isEmpty = value == "";

				if (isArray):
					basicType = getArrayBasicType(celType);

					if (not celName in currentObject):
						currentObject[celName] = [];

					if (not isEmpty):
						rowToPush = i
						colToPush = j

						if not directionIsHorizontal:
							rowToPush = j
							colToPush = i

						pushValueIntoArray(currentObject[celName], celName, basicType, value, sheet, rowToPush, colToPush);
					elif (allArrays):
						currentObjectEmptyArrayEntriesFound.append(j);
				else:
					if (not isEmpty):
						#Browser.msgBox('Result', "pushin " + " " + colName + " "  + value , Browser.Buttons.OK); 

						rowToPush = i
						colToPush = j

						if not directionIsHorizontal:
							rowToPush = j
							colToPush = i

						parseValueIntoObject(currentObject, celName, basicType, value, sheet, rowToPush, colToPush);

				#Browser.msgBox('Result', colTypeAndName["type"] + " " + values[firstValidRow][j], Browser.Buttons.OK);


	return object;

def getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal):
	validRowsIndices = [];
	invalidRowsIndices = [];

	validColsIndices = [];
	invalidColsIndices = [];

	totalRows = sheet.getLastRow();
	totalCols = sheet.getLastColumn();
	values = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
	lineOrColIndexWithFirstJsonType = -1;

	knownTypes = [];

	end = totalRows

	if not directionIsHorizontal:
		end = totalCols

	for i in range(0,end):

		end2 = totalCols

		if not directionIsHorizontal:
			end2 = totalRows

		for j in range(0,end2):

			entryValue = ""

			if directionIsHorizontal:
				entryValue = entryValue + values[i][j];
			else:
				entryValue = entryValue + values[j][i];

			entryValue = entryValue.lower();

			if (isJSONType(entryValue)):
				if (lineOrColIndexWithFirstJsonType == -1):
					lineOrColIndexWithFirstJsonType = i;

				if (lineOrColIndexWithFirstJsonType == i and not entryValue in knownTypes):
					knownTypes.append(entryValue);

					rowToPush = i
					colToPush = j

					if not directionIsHorizontal:
						rowToPush = j
						colToPush = i

					validRowsIndices.append(rowToPush);
					validColsIndices.append(colToPush);
				else:
					rowToPush = i
					colToPush = j

					if not directionIsHorizontal:
						rowToPush = j
						colToPush = i

					invalidRowsIndices.append(rowToPush);
					invalidColsIndices.append(colToPush);
			elif (entryValue != "" and lineOrColIndexWithFirstJsonType == i):
				rowToPush = i
				colToPush = j

				if not directionIsHorizontal:
					rowToPush = j
					colToPush = i

				invalidRowsIndices.append(rowToPush);
				invalidColsIndices.append(colToPush);

	returnObj = {};
	returnObj["validRowsIndices"] = validRowsIndices;
	returnObj["invalidRowsIndices"] = invalidRowsIndices;
	returnObj["validColsIndices"] = validColsIndices;
	returnObj["invalidColsIndices"] = invalidColsIndices;

	#printObj(returnObj);

	return returnObj;


tsvFileName = sys.argv[1]

if not os.path.isfile(tsvFileName):
    print('File does not exist.')
    exit(1)

f = open(tsvFileName, "r")
content = f.read()

sheet = Sheet(content)

directionIsHorizontal = detectDirection(sheet);

object = createObject(sheet, tsvFileName, directionIsHorizontal, True);

str = json.dumps(object)

print(str)

