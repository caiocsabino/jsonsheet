import re
import sys
import os
import subprocess
import zipfile
import io
import base64
import json
import re

jsonBasicTypes = ["string", "bool", "uint", "int", "uint64", "float", "double", "long", "ulong", "object"]
jsonTypesAsStrings = [];
jsonArrayTypesAsStrings = []
arraySeparator = ","
emptyLinesAllowed = True
tempSheetIdentifierPreamble = "_%_";
tempSheetIdentifierPostamble = "_#_";
compressedObjPreamble = ":__JSON_OBJ__:";
isRootObjectArray = True;
tempSheetHierarchySeparator = ">";
validationHierarchySeparator = ".";
errorList = [];
columnsAsLetters = True;
allowDefaultValuesWhenEmpty = True;

class Sheet:
	def __init__(self, csvString, jsonString):
		self.csvString = csvString
		self.jsonString = jsonString

		if csvString != "":
			self.setupFromCSV()

		if jsonString != "":
			self.setupFromJson()

	def setupFromCSV(self):
		rows = self.csvString.splitlines()

		self.totalRows = len(rows)
		self.totalCols = 1
		self.sheetValues = []

		for i in range(0,self.totalRows):
			self.sheetValues.append([])

			cols = rows[i].split("\t")

			if len(cols) > self.totalCols:
				self.totalCols = len(cols)

			for k in range(0,self.totalCols):
				if k < len(cols):
					self.sheetValues[i].append(cols[k])
				else:
					self.sheetValues[i].append("")

		for i in range(0,self.totalRows):
			diff = self.totalCols - len(self.sheetValues[i])
			for k in range(0,diff):
				self.sheetValues[i].append("")






	def setupFromJson(self):
		jsonObject = json.loads(self.jsonString)

		if (jsonObject != None and "cels" in jsonObject):
			self.totalRows = jsonObject["cels"]["totalRows"];
			self.totalCols = jsonObject["cels"]["totalCols"];
			self.sheetValues = []

			jsonValues = jsonObject["cels"]["values"];
			formulas = jsonObject["cels"]["formulas"];
			bgs = jsonObject["cels"]["bgs"];

			tsvString = ""

			for i in range(0,self.totalRows):
				self.sheetValues.append([])
				for j in range(0,self.totalCols):
					self.sheetValues[i].append(str(jsonValues[i*self.totalCols+j]))

	def getSheetValues(self,startRow, startCol, totalRows, totalCols):

		tempValues = []

		for i in range(0,totalRows):
			tempValues.append([])

			for j in range(0,totalCols):
				tempValues[i].append(self.sheetValues[startRow-1+i][startCol-1+j])

		return tempValues

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

	def setName(self, name):
		self.name = name

	def getName(self):
		return self.name


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
				entryValue = None

				if not directionIsHorizontal:
					entryValue =  values[j][i]
				else:
					entryValue = "" + (values[i][j]);

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

def isIntString(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def isFloatString(s):
    try:
        float(s)
        return True
    except ValueError:
        return False

def getColAsLetter(col):
	if (not columnsAsLetters):
		return "" + col + 1;

	letters = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"];
	multiplier = int(col / len(letters));

	if (multiplier > 0):
		return letters[multiplier] + letters[col % letters.length];
	else:
		return letters[col];

def parseValueIntoObject(object, entryName, entryBasicType, value, sheet, row, col, currentSheetName):
	if (entryBasicType == "(int)" or entryBasicType == "(uint)" or entryBasicType == "(int64)" or entryBasicType == "(uint64)"or entryBasicType == "(long)" or entryBasicType == "(ulong)"):
		if value == "" and allowDefaultValuesWhenEmpty:
			value = "0"

		if isIntString(value):
			object[entryName] = int(value);
		else:
			errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not an integer: " + value);

	elif (entryBasicType == "(float)" or entryBasicType == "(double)"):
		if value == "" and allowDefaultValuesWhenEmpty:
			value = "0"

		if isFloatString(value):
			object[entryName] = float(value);
		else:
			errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not a float: " + value);

	elif (entryBasicType == "(bool)"):
		if value == "" and allowDefaultValuesWhenEmpty:
			value = "0"

		valueLower = value.lower();

		if valueLower == "1" or valueLower == "0" or valueLower == "true" or valueLower == "false":
			if valueLower == "1" or valueLower == "true":
				object[entryName] = True;
			else:
				object[entryName] =  False;
		else:
			errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not a boolean: " + value);

	elif (entryBasicType == "(string)"):
		object[entryName] = value;
	elif (entryBasicType == "(object)"):

		if (isJsonString(value)):
			object[entryName] = value;
		else:
			# targetName = tempSheetIdentifierPreamble + sheet.getSheetId() + tempSheetIdentifierPostamble + row + "," + col;

			if compressedObjPreamble in value:
				targetName = "TEMPSHEET"
				newSheet = deserializeSheet(targetName, value);

				horizontal = detectDirection(newSheet);

				newObject = createObject(newSheet, entryName, horizontal, False, currentSheetName + validationHierarchySeparator + entryName + "(" + str(row + 1) + "," + getColAsLetter(col) + ")");

				# if (newObject):
				# 	object[entryName] = newObject;

				# 	activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
				# 	activeSpreadsheet.deleteSheet(newSheet);

				# spreadsheet = SpreadsheetApp.getActive();
				# spreadsheet.setActiveSheet(sheet);

				# by default all deserialized sheets are treated as arrays, so since here we know it is a single object we must create it from the first entry

				object[entryName] = newObject[entryName][0];

				#object[entryName] = value;
			else:
				errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not a sheet nor json object: " + value);
	else:
		print("WILL BREAK, unkown type " + entryBasicType);

def pushValueIntoArray(array, entryName, basicType, value, sheet, row, col, currentSheetName):
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

			if value == "" and allowDefaultValuesWhenEmpty:
				value = "0"

			if isIntString(value):
				array.append(int(value));
			else:
				errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not an int: " + value);

	elif (basicType == "float" or basicType == "double"):
		for i in range(0,len(values)):
			value = values[i];

			if value == "" and allowDefaultValuesWhenEmpty:
				value = "0"

			if isFloatString(value):
				array.append(float(value));
			else:
				errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not a float: " + value);

	elif (basicType == "bool"):
		for i in range(0,len(values)):
			value = values[i];

			if value == "" and allowDefaultValuesWhenEmpty:
				value = "0"

			valueLower = value.lower();

			if valueLower == "1" or valueLower == "0" or valueLower == "true" or valueLower == "false":
				if valueLower == "1" or valueLower == "true":
					array.append(True);
				else:
					array.append(False)
			else:
				errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not a boolean: " + value);

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

			if compressedObjPreamble in value:
				targetName = "TEMPSHEET"

				newSheet = deserializeSheet(targetName, value);

				horizontal = detectDirection(newSheet);

				newObject = createObject(newSheet, entryName, horizontal, False, currentSheetName + validationHierarchySeparator + entryName + "(" + str(row + 1) + "," + getColAsLetter(col) + ")");

				# if (newObject):
				# 	object[entryName] = newObject;

				# 	activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
				# 	activeSpreadsheet.deleteSheet(newSheet);

				# spreadsheet = SpreadsheetApp.getActive();
				# spreadsheet.setActiveSheet(sheet);

				for i in range(0,len(newObject[entryName])):
					array.append(newObject[entryName][i]);
			else:
				errorList.append("Error:" + currentSheetName + " Row: " + str(row + 1) + " Col: " + getColAsLetter(col) + "; Value is not a sheet nor json object: " + value);

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

			newSheet = Sheet("", unzipped)

			return newSheet

	return None;

def showErrors():
	if (len(errorList) == 0):
		print('No errors found');
	else:
		errorStr = "";
		for i in range(0,len(errorList)):
			print(errorList[i])

def validateSheet(sheet, sheetName):

	directionIsHorizontal = detectDirection(sheet);

	errorList = [];

	object = createObject(sheet, sheetName, directionIsHorizontal, true, sheetName);

	showErrors()

def isInvalidCel(invalidRowsList, invalidColsList, row, col):
	for i in range(0,len(invalidRowsList)):
		if invalidRowsList[i] == row and invalidColsList[i] == col:
			return True
	return False


def createObject(sheet, name, directionIsHorizontal, isRoot, currentSheetName):
	result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	checkForDuplicateEntries(sheet, directionIsHorizontal, currentSheetName)

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

		start = None
		end = None

		if not directionIsHorizontal:
			start = firstValidCol + 1
			end = lastNonEmptyCol
		else:
			start = firstValidRow + 1
			end = lastNonEmptyRow

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
					copy = values
					content = None

					if not directionIsHorizontal:
						content = values[j][firstValidCol];
					else:
						content = values[firstValidRow][j];

					if (content == ""):
						continue;

					if (directionIsHorizontal and isInvalidCel(result["invalidRowsIndices"], result["invalidColsIndices"],firstValidRow, j)) or (not directionIsHorizontal and isInvalidCel(result["invalidRowsIndices"], result["invalidColsIndices"],j, firstValidCol)):
						continue;

					celTypeAndName = getCelTypeAndName(content);
					celType = celTypeAndName["type"];
					celName = celTypeAndName["name"];
					isArray = isArrayType(celType);
					basicType = celType;
					value = None;

					if not directionIsHorizontal:
						value = values[j][i]
					else:
						value = values[i][j];

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
				content = None;

				if not directionIsHorizontal:
					content = values[j][firstValidCol]
				else:
					content = values[firstValidRow][j];

				if (content == ""):
					continue;

				if (directionIsHorizontal and isInvalidCel(result["invalidRowsIndices"], result["invalidColsIndices"],firstValidRow, j)) or (not directionIsHorizontal and isInvalidCel(result["invalidRowsIndices"], result["invalidColsIndices"],j, firstValidCol)):
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

						pushValueIntoArray(currentObject[celName], celName, basicType, value, sheet, rowToPush, colToPush, currentSheetName);
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

						parseValueIntoObject(currentObject, celName, basicType, value, sheet, rowToPush, colToPush, currentSheetName);

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

			entryValueLower = entryValue.lower();

			if (isJSONType(entryValueLower)):
				entryName = getCelTypeAndName(entryValue)["name"];

				if (lineOrColIndexWithFirstJsonType == -1):
					lineOrColIndexWithFirstJsonType = i;

				if (lineOrColIndexWithFirstJsonType == i and not entryName in knownTypes):
					knownTypes.append(entryName);

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

def checkForDuplicateEntries(sheet, directionIsHorizontal, currentSheetName):
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

			entryValueLower = entryValue.lower();

			if (isJSONType(entryValueLower)):
				entryName = getCelTypeAndName(entryValue)["name"];

				if (lineOrColIndexWithFirstJsonType == -1):
					lineOrColIndexWithFirstJsonType = i;

				if (lineOrColIndexWithFirstJsonType == i and not entryName in knownTypes):
					knownTypes.append(entryName);

				else:
					if (entryName in knownTypes):
						warningMessage = "";

						if (directionIsHorizontal):
							warningMessage = ("Warning: Ignoring " + currentSheetName + " Row: " + str(lineOrColIndexWithFirstJsonType + 1) + " Col: " + getColAsLetter(j) + "; Duplicate entry in same sheet: " + entryValue);
						else:
							warningMessage = ("Warning: Ignoring " + currentSheetName + " Row: " + str(j + 1) + " Col: " + getColAsLetter(lineOrColIndexWithFirstJsonType) + "; Duplicate entry in same sheet: " + entryValue);

						if (not warningMessage in errorList):
							errorList.append(warningMessage);
					else:
						warningMessage = "";

						if (directionIsHorizontal):
							warningMessage = ("Warning: Ignoring " + currentSheetName + " Row: " + str(lineOrColIndexWithFirstJsonType + 1) + " Col: " + getColAsLetter(j) + "; " + content + ". Bad location, type declaratios must be at row: " + str(lineOrColIndexWithFirstJsonType + 1));
						else:
							warningMessage = ("Warning: Ignoring " + currentSheetName + " Row: " + str(j + 1) + " Col: " + getColAsLetter(lineOrColIndexWithFirstJsonType) + "; " + content + ". Bad location, type declaratios must be at col: " + getColAsLetter(lineOrColIndexWithFirstJsonType));

						if (not warningMessage in errorList):
							errorList.append(warningMessage);


			elif (entryValue != "" and lineOrColIndexWithFirstJsonType == i):
				warningMessage = "";

				if (directionIsHorizontal):
					warningMessage = ("Warning: Ignoring " + currentSheetName + " Row: " + str(lineOrColIndexWithFirstJsonType + 1) + " Col: " + getColAsLetter(j) + "; Invalid type: " + entryValue);
				else:
					warningMessage = ("Warning: Ignoring " + currentSheetName + " Row: " + str(j + 1) + " Col: " + getColAsLetter(lineOrColIndexWithFirstJsonType) + "; Invalid type: " + entryValue);
				
				if (not warningMessage in errorList):
					errorList.append(warningMessage);



tsvFileName = sys.argv[1]

if (len(sys.argv) > 1):
	outputFile = sys.argv[2]

if not os.path.isfile(tsvFileName):
    print('File does not exist.')
    exit(1)

f = open(tsvFileName, "r")
content = f.read()

sheet = Sheet(content, "")

horizontal = detectDirection(sheet);

strippedTSVFileName = tsvFileName.replace(".tsv", "")

index = strippedTSVFileName.rfind("/")

strippedTSVFileName = strippedTSVFileName[index+1:]

object = createObject(sheet, strippedTSVFileName, horizontal, True, strippedTSVFileName);

showErrors()

if (len(sys.argv) > 1):
	str = json.dumps(object)

	fileOut = open(outputFile, "w")

	fileOut.write(json.dumps(object, sort_keys=False, indent=0, separators=(',', ':')))

	fileOut.close()

print(str)


