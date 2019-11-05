// [START apps_script_bracketmaker]
// This script works with the Brackets Test spreadsheet to create a tournament bracket
// given a list of players or teams.

var RANGE_PLAYER1 = 'FirstPlayer';
var SHEET_PLAYERS = 'Players';
var SHEET_BRACKET = 'Bracket';
var CONNECTOR_WIDTH = 15;

// large spreadsheets


/**
 * This method adds a custom menu item to run the script
 */
function onOpen()
{
	var ss = SpreadsheetApp.getActiveSpreadsheet();

	var ui = SpreadsheetApp.getUi();
	// Or DocumentApp or FormApp.
	ui.createMenu('JSONSheet')
		.addItem('Validate Sheet', 'validateSheet').addItem('Export as JSON', 'exportJSON').addItem('Import JSON', 'importJSON').addItem('Go Inside', 'goInside').addItem('Save Temp Sheets', 'saveTempSheets').addItem('Delete Temp Sheets', 'deleteTempSheets')
		.addToUi();


}

var jsonBasicTypes = ["string", "bool", "uint", "int", "uint64", "float", "object"];
var jsonTypesAsStrings = [];
var jsonArrayTypesAsStrings = [];
var arraySeparator = ",";
var emptyLinesAllowed = true;
var tempSheetIdentifierPreamble = "__.TEMP_JSON_SHEET.__";
var tempSheetIdentifierPostamble = "__.LOCATION.__";
var compressedObjPreamble = ":__JSON_OBJ__:";
var isRootObjectArray = true;

for (var i = 0; i < jsonBasicTypes.length; i++)
{
	jsonTypesAsStrings.push("(" + jsonBasicTypes[i] + ")");
	jsonTypesAsStrings.push("(" + jsonBasicTypes[i] + "s)");

	jsonArrayTypesAsStrings.push("(" + jsonBasicTypes[i] + "s)");
}

function isObjectType(type)
{
	return type == "(object)"
}

function isArrayType(type)
{
	return jsonArrayTypesAsStrings.indexOf(type) > -1;
}

function getArrayBasicType(arrayType)
{
	var lastLetters = arrayType.substr(arrayType.length - 2, 2)

	if (lastLetters == "s)")
	{
		return arrayType.substr(1, arrayType.length - 3)
	}

	return null
}

function deleteTempSheets()
{
	var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var sheets = activeSpreadsheet.getSheets();

	var sheetsToDelete = [];

	for (var i = 0; i < sheets.length; i++)
	{
		var candidate = sheets[i];

		var index = ("" + candidate.getName()).indexOf(tempSheetIdentifierPreamble);

		if (index > -1)
		{
			sheetsToDelete.push(candidate);
		}

	}

	for (var i = 0; i < sheetsToDelete.length; i++)
	{
		activeSpreadsheet.deleteSheet(sheetsToDelete[i]);
	}


}

function serializeSheet(sheet)
{
	var obj = {};
	obj["cels"] = {};



	var totalRows = sheet.getLastRow();
	var totalCols = sheet.getLastColumn();

	var values = sheet.getSheetValues(1, 1, totalRows, totalCols);

	var range = sheet.getDataRange();
	var formulas = range.getFormulas();
	var bgs = range.getBackgrounds();

	obj["cels"]["totalRows"] = totalRows;
	obj["cels"]["totalCols"] = totalCols;
	obj["cels"]["values"] = [];
	obj["cels"]["formulas"] = [];
	obj["cels"]["bgs"] = [];
	obj["cels"]["fcs"] = [];

	for (var i = 0; i < totalRows; i++)
	{
		for (var j = 0; j < totalCols; j++)
		{
			obj["cels"]["values"].push(values[i][j]);
			obj["cels"]["formulas"].push(formulas[i][j]);

			obj["cels"]["bgs"].push(bgs[i][j]);
		}
	}

	obj["json"] = "";

	var str = JSON.stringify(obj);

	var blob = Utilities.newBlob(str, 'application/octet-stream');

	var compressedBlob = Utilities.zip([blob]);

	var encoded = compressedObjPreamble + Utilities.base64Encode(compressedBlob.getBytes());
	return encoded;
}



function deserializeSheet(sheetName, data)
{
	var activeSheet = SpreadsheetApp.getActiveSheet();

	var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var sheets = activeSpreadsheet.getSheets();

	var targetSheet = null;

	for (var i = 0; i < sheets.length; i++)
	{
		var candidate = sheets[i];

		if (candidate.getSheetId() == sheetName)
		{
			targetSheet = candidate;
			break;
		}

	}

	if (targetSheet != null)
	{
		activeSpreadsheet.deleteSheet(targetSheet);
	}

	var newSheet = activeSpreadsheet.insertSheet();
	newSheet.setName(sheetName);

	// insert data

	if (data != "")
	{
		if (data.indexOf(compressedObjPreamble) > -1)
		{
			data = data.replace(compressedObjPreamble, "");
		}


		var decoded = Utilities.base64Decode(data);

		var blob = Utilities.newBlob(decoded, 'application/zip');

		var unzipped = Utilities.unzip(blob);

		var jsonObject = JSON.parse(unzipped[0].getAs('application/octet-stream').getDataAsString());


		if (jsonObject != null && jsonObject["cels"] != null)
		{
			var totalRows = jsonObject["cels"]["totalRows"];
			var totalCols = jsonObject["cels"]["totalCols"];

			var values = jsonObject["cels"]["values"];
			var formulas = jsonObject["cels"]["formulas"];
			var bgs = jsonObject["cels"]["bgs"];

			for (var i = 0; i < values.length; i++)
			{
				var row = i / totalCols;
				var col = i % totalCols;

				//SpreadsheetApp.getActiveSheet().getRange('F2').setValue('Hello');

				var range = newSheet.getRange(row + 1, col + 1);

				if (formulas != null && formulas[i] != "")
				{
					range.setFormula(formulas[i]);

				}
				else
				{
					var valueToInsert = "" + values[i];

					range.setValue(valueToInsert);

					if (valueToInsert.indexOf(compressedObjPreamble) > -1)
					{
						range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
					}
				}

				if (bgs != null)
				{
					range.setBackground(bgs[i]);
				}
			}
		}

	}

	//Browser.msgBox('Data', 'Deserialization with ' + data, Browser.Buttons.OK);

	return newSheet;



}

function getSheetById(sheetId)
{
	var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var sheets = activeSpreadsheet.getSheets();

	var sheetsToDelete = [];

	for (var i = 0; i < sheets.length; i++)
	{
		var sheet = sheets[i];
		if (sheetId == sheet.getSheetId())
		{
			return sheet
		}

	}

	return null;
}

function saveTempSheets()
{
	//insertLargeJson();
	//readLargeJson();
	//return;

	var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	var sheets = activeSpreadsheet.getSheets();

	var sheetsToDelete = [];

	for (var i = 0; i < sheets.length; i++)
	{
		var candidate = sheets[i];

		var candidateName = candidate.getName();

		var index = candidateName.indexOf(tempSheetIdentifierPreamble);

		if (index > -1)
		{
			var postEmbleIndex = candidateName.indexOf(tempSheetIdentifierPostamble);

			if (postEmbleIndex > -1)
			{

				var parentSheetId = candidateName.substr(tempSheetIdentifierPreamble.length, postEmbleIndex - tempSheetIdentifierPreamble.length);
				var locationStr = candidateName.substr(postEmbleIndex + tempSheetIdentifierPostamble.length, candidateName.length - postEmbleIndex + tempSheetIdentifierPostamble.length);
				var cels = locationStr.split(",");
				var parentSheetRow = parseInt(cels[0]);
				var parentSheetCol = parseInt(cels[1]);

				var serialization = serializeSheet(candidate);

				var parentSheet = getSheetById(parentSheetId);

				while (parentSheet != null)
				{
					var range = parentSheet.getRange(parentSheetRow + 1, parentSheetCol + 1);
					range.setValue(serialization);
					range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

					SpreadsheetApp.flush();

					var parentSheetName = parentSheet.getName();


					var newIndex = parentSheetName.indexOf(tempSheetIdentifierPreamble);

					if (newIndex > -1)
					{
						var newPostEmbleIndex = parentSheetName.indexOf(tempSheetIdentifierPostamble);

						if (newPostEmbleIndex > -1)
						{
							var newParentSheetId = parentSheetName.substr(tempSheetIdentifierPreamble.length, newPostEmbleIndex - tempSheetIdentifierPreamble.length);
							var newLocationStr = parentSheetName.substr(newPostEmbleIndex + tempSheetIdentifierPostamble.length, parentSheetName.length - newPostEmbleIndex + tempSheetIdentifierPostamble.length);
							var newCels = newLocationStr.split(",");
							parentSheetRow = parseInt(newCels[0]);
							parentSheetCol = parseInt(newCels[1]);

							serialization = serializeSheet(parentSheet);

							parentSheet = getSheetById(newParentSheetId);
						}
						else
						{
							parentSheet = null;
						}

					}
					else
					{
						parentSheet = null;
					}

				}

				sheetsToDelete.push(candidate);

			}

		}

	}



	for (var i = 0; i < sheetsToDelete.length; i++)
	{
		//activeSpreadsheet.deleteSheet(sheetsToDelete[i]);
	}

}

function canGoInside(row, col, sheet, directionIsHorizontal)
{

	var result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	if (result["validRowsIndices"].length > 0)
	{
		var firstValidRow = result["validRowsIndices"][0];
		var firstValidCol = result["validColsIndices"][0];
		var lastValidCol = result["validColsIndices"][result["validColsIndices"].length - 1];
		var lastValidRow = result["validRowsIndices"][result["validRowsIndices"].length - 1];

		var totalRows = sheet.getLastRow();
		var totalCols = sheet.getLastColumn();

		var values = sheet.getSheetValues(1, 1, totalRows, totalCols);

		if (directionIsHorizontal)
		{
			var index = result["validColsIndices"].indexOf(col);

			if (index > -1 && row > firstValidRow && col >= firstValidCol && col <= lastValidCol)
			{
				var celTypeAndName = getCelTypeAndName(values[firstValidRow][index]);
				var celType = celTypeAndName["type"];
				var isArray = isArrayType(celType);
				var isObject = isObjectType(celType);

				return (isArray || isObject);
			}
		}
		else
		{
			var index = result["validRowsIndices"].indexOf(row);

			if (index > -1 && col > firstValidCol && row >= firstValidRow && row <= lastValidRow)
			{
				var celTypeAndName = getCelTypeAndName(values[row][firstValidCol]);
				var celType = celTypeAndName["type"];
				var isArray = isArrayType(celType);
				var isObject = isObjectType(celType);

				return (isArray || isObject);
			}
		}
	}

	return false;
}

function goInside()
{
	var activeSheet = SpreadsheetApp.getActiveSheet();
	var selection = activeSheet.getSelection();

	var directionIsHorizontal = detectDirection(activeSheet);

	var row = selection.getCurrentCell().getRow() - 1;
	var col = selection.getCurrentCell().getColumn() - 1;

	if (canGoInside(row, col, activeSheet, directionIsHorizontal))
	{
		var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
		var sheets = activeSpreadsheet.getSheets();

		var targetName = tempSheetIdentifierPreamble + activeSheet.getSheetId() + tempSheetIdentifierPostamble + row + "," + col;

		var existingSheet = activeSpreadsheet.getSheetByName(targetName);

		if (existingSheet)
		{
			activeSpreadsheet.setActiveSheet(existingSheet);
		}
		else
		{
			var totalRows = activeSheet.getLastRow();
			var totalCols = activeSheet.getLastColumn();

			var values = activeSheet.getSheetValues(1, 1, totalRows, totalCols);

			var value = values[row][col];

			var newSheet = deserializeSheet(targetName, value);
		}



	}
	else
	{
		Browser.msgBox('Error', 'Not an array or object type', Browser.Buttons.OK);
	}

}

function isEmptyRowOrCol(index, useRow, sheet)
{
	var totalRows = sheet.getLastRow();
	var totalCols = sheet.getLastColumn();

	var values = useRow ? sheet.getSheetValues(index + 1, 1, 1, totalCols) : sheet.getSheetValues(1, index + 1, totalRows, 1);

	var end = useRow ? totalCols : totalRows;

	for (var i = 0; i < end; i++)
	{
		var value = useRow ? values[0][i] : values[i][0];
		if (value != "")
		{
			return false;
		}

	}

	return true;
}

function getCelTypeAndName(content)
{
	var result = {};

	// TODO: Use regex to extract
	for (var i = 0; i < jsonTypesAsStrings.length; i++)
	{
		var target = jsonTypesAsStrings[i];

		var index = content.indexOf(target);

		if (index > -1)
		{
			//result["type"] = target.substr(1, target.length - 2);
			result["type"] = target;
			result["name"] = content.replace(target, "");
			return result;
		}
	}
}

function isJSONType(givenType)
{
	var typeLower = givenType.toLowerCase();

	for (var k = 0; k < jsonTypesAsStrings.length; k++)
	{
		var targetLower = jsonTypesAsStrings[k].toLowerCase();

		//Browser.msgBox('Result', type, Browser.Buttons.OK);

		if (typeLower.indexOf(targetLower) > -1)
		{
			return true;
		}

	}

	return false;
}

function detectDirection(sheet)
{
	validRowsIndices = [];
	invalidRowsIndices = [];

	validColsIndices = [];
	invalidColsIndices = [];

	var totalRows = sheet.getLastRow();
	var totalCols = sheet.getLastColumn();
	var values = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
	var lineIndexWithFirstJsonType = -1;
	var colIndexWithFirstJsonType = -1;

	for (var i = 0; i < totalRows; i++)
	{
		for (var j = 0; j < totalCols; j++)
		{
			var entryValue = "" + values[i][j];
			entryValue = entryValue.toLowerCase();

			if (isJSONType(entryValue))
			{
				if (lineIndexWithFirstJsonType == -1)
				{
					lineIndexWithFirstJsonType = i;
					colIndexWithFirstJsonType = j;
				}

				var typedCelsInSameRow = 1;
				var typedCelsInSameCol = 1;

				for (var kk = i + 1; kk < totalRows; kk++)
				{
					var celValue = "" + values[kk][j];
					celValue = celValue.toLowerCase();
					if (isJSONType(celValue))
					{
						typedCelsInSameCol = typedCelsInSameCol + 1;
					}
					else
					{
						break;
					}

				}

				for (var kk = j + 1; kk < totalCols; kk++)
				{
					var celValue = "" + values[i][kk];
					celValue = celValue.toLowerCase();
					if (isJSONType(celValue))
					{
						typedCelsInSameRow = typedCelsInSameRow + 1;
					}
					else
					{
						break;
					}
				}

				return typedCelsInSameRow > typedCelsInSameCol;
			}


		}
	}


}

function validateSheet()
{
	//Browser.msgBox('Result', basicType, Browser.Buttons.OK);

	var spreadsheet = SpreadsheetApp.getActive();

	var sheet = spreadsheet.getActiveSheet();

	var directionIsHorizontal = detectDirection(sheet);

	paintValidation(sheet, directionIsHorizontal);

	var htmlOutput = HtmlService
		.createHtmlOutput('<p>A change of speed, a change of style...</p>')
		.setWidth(250)
		.setHeight(300);
	SpreadsheetApp.getUi().showModelessDialog(htmlOutput, 'My add-on');

	return true;
}

function exportJSON()
{
	var spreadsheet = SpreadsheetApp.getActive();

	var sheet = spreadsheet.getActiveSheet();

	var directionIsHorizontal = detectDirection(sheet);

	var object = createObject(sheet, sheet.getName(), directionIsHorizontal, true);

	var spreadsheet = SpreadsheetApp.getActive();
	spreadsheet.setActiveSheet(sheet);
	/*
	if (validateSheet())
	{

	}
    */

	var str = JSON.stringify(object)

	Browser.msgBox('Result', str, Browser.Buttons.OK);
}


function paintValidation(sheet, directionIsHorizontal)
{
	var result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	if (result["validRowsIndices"].length > 0)
	{
		var firstValidRow = result["validRowsIndices"][0];
		var firstValidCol = result["validColsIndices"][0];
		var lastValidCol = result["validColsIndices"][result["validColsIndices"].length - 1];
		var lastValidRow = result["validRowsIndices"][result["validRowsIndices"].length - 1];

		var lastValidRowAndNonEmptyRowOrCol = getLastValidRowAndNonEmptyRow(emptyLinesAllowed, directionIsHorizontal);
		var lastValidRowOrCol = lastValidRowAndNonEmptyRowOrCol[directionIsHorizontal ? "lastValidRow" : "lastValidCol"];
		var lastNonEmptyRowOrCol = lastValidRowAndNonEmptyRowOrCol[directionIsHorizontal ? "lastNonEmptyRow" : "lastNonEmptyCol"];

		if (directionIsHorizontal)
		{
			sheet.getRange(firstValidRow + 1, firstValidCol + 1, lastNonEmptyRowOrCol - firstValidRow + 1, lastValidCol - firstValidCol + 1).setBackground("green");
		}
		else
		{
			sheet.getRange(firstValidRow + 1, firstValidCol + 1, lastValidRow - firstValidRow + 1, lastNonEmptyRowOrCol - firstValidCol + 1).setBackground("green");
		}


	}
}

function printObj(obj)
{
	if (obj != null)
	{
		var str = JSON.stringify(obj)

		Browser.msgBox('Debug obj', str, Browser.Buttons.OK);
	}
}


function getLastValidRowAndNonEmptyRow(_emptyLinesAllowed, sheet, directionIsHorizontal)
{
	var returnObj = {};

	var result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	if (result[(directionIsHorizontal ? "validRowsIndices" : "validColsIndices")].length > 0)
	{
		var firstValidRow = result["validRowsIndices"][0];
		var firstValidCol = result["validColsIndices"][0];
		var lastValidCol = result["validColsIndices"][result["validColsIndices"].length - 1];
		var lastValidRow = result["validRowsIndices"][result["validRowsIndices"].length - 1];
		var lastValidRowOrCol = -1;
		var lastNonEmptyRowOrCol = -1;

		var totalRows = sheet.getLastRow();
		var totalCols = sheet.getLastColumn();

		var values = sheet.getSheetValues(1, 1, totalRows, totalCols);

		for (var i = (directionIsHorizontal ? firstValidRow : firstValidCol); i < (directionIsHorizontal ? totalRows : totalCols); i++)
		{

			var isEmptyLineOrCol = true;

			for (var j = (directionIsHorizontal ? firstValidCol : firstValidRow); j <= (directionIsHorizontal ? lastValidCol : lastValidRow); j++)
			{

				var entryValue = "" + (directionIsHorizontal ? values[i][j] : values[j][i]);

				if (entryValue != "")
				{
					isEmptyLineOrCol = false;
					break;
				}

			}

			var testRows = directionIsHorizontal && (result["invalidRowsIndices"].length > 0 && result["invalidRowsIndices"][0] == i && result["invalidColsIndices"][0] == j);
			var testCols = !directionIsHorizontal && (result["invalidColsIndices"].length > 0 && result["invalidColsIndices"][0] == i && result["invalidRowsIndices"][0] == j);

			var shouldEnd = (isEmptyLineOrCol && !_emptyLinesAllowed) || testRows || testCols;


			if (!shouldEnd)
			{
				lastValidRowOrCol = i;
				if (!isEmptyLineOrCol)
				{
					lastNonEmptyRowOrCol = i;
				}
			}
			else
			{
				break;
			}

		}

		if (directionIsHorizontal)
		{
			returnObj["lastValidRow"] = lastValidRowOrCol;
			returnObj["lastNonEmptyRow"] = lastNonEmptyRowOrCol;
		}
		else
		{
			returnObj["lastValidCol"] = lastValidRowOrCol;
			returnObj["lastNonEmptyCol"] = lastNonEmptyRowOrCol;
		}

		return returnObj;
	}
}

function isJsonString(str)
{
	try
	{
		JSON.parse(str);
	}
	catch (e)
	{
		return false;
	}

	return true;
}

function parseValueIntoObject(object, entryName, entryBasicType, value, sheet, row, col)
{
	if (entryBasicType == "(int)" || entryBasicType == "(uint)" || entryBasicType == "(int64)" || entryBasicType == "(uint64)")
	{
		object[entryName] = parseInt(value);
	}
	else if (entryBasicType == "(float)")
	{
		object[entryName] = parseFloat(value);
	}
	else if (entryBasicType == "(bool)")
	{
		var valueLower = value.toLowerCase();
		object[entryName] = valueLower == "1" || valueLower == "true" ? true : false;
	}
	else if (entryBasicType == "(string)")
	{
		object[entryName] = value;
	}
	else if (entryBasicType == "(object)")
	{

		if (isJsonString(value))
		{
			object[entryName] = value;
		}
		else
		{
			var targetName = tempSheetIdentifierPreamble + sheet.getSheetId() + tempSheetIdentifierPostamble + row + "," + col;

			var newSheet = deserializeSheet(targetName, value);

			var directionIsHorizontal = detectDirection(newSheet);

			var newObject = createObject(newSheet, entryName, directionIsHorizontal);

			if (newObject)
			{
				object[entryName] = newObject;

				var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
				activeSpreadsheet.deleteSheet(newSheet);
			}

			var spreadsheet = SpreadsheetApp.getActive();
			spreadsheet.setActiveSheet(sheet);

			// by default all deserialized sheets are treated as arrays, so since here we know it is a single object we must create it from the first entry

			object[entryName] = newObject[entryName][0];


			//object[entryName] = value;
		}



	}
	else
	{
		Browser.msgBox('Result', "WILL BREAK, unkown type " + entryBasicType, Browser.Buttons.OK);
	}

}

function pushValueIntoArray(array, entryName, basicType, value, sheet, row, col)
{
	if (value == null)
	{
		return;
	}
	value = "" + value;

	var hasSeparator = value.indexOf(arraySeparator) > -1;

	var values = [];

	if (hasSeparator && basicType != "object")
	{
		values = value.split(arraySeparator);
	}
	else
	{
		values.push(value);
	}

	for (var i = 0; i < values.length; i++)
	{
		value = values[i];
	}

	if (basicType == "int" || basicType == "uint" || basicType == "int64" || basicType == "uint64")
	{
		for (var i = 0; i < values.length; i++)
		{
			value = values[i];
			array.push(parseInt(value));
		}
	}
	else if (basicType == "float")
	{
		for (var i = 0; i < values.length; i++)
		{
			value = values[i];
			array.push(parseFloat(value));
		}
	}
	else if (basicType == "bool")
	{
		for (var i = 0; i < values.length; i++)
		{
			value = values[i];
			var valueLower = value.toLowerCase();
			array.push(valueLower == "1" || valueLower == "true" ? true : false);
		}

	}
	else if (basicType == "string")
	{
		for (var i = 0; i < values.length; i++)
		{
			value = values[i];
			array.push(value);
		}
	}
	else if (basicType == "object")
	{

		if (isJsonString(value))
		{
			var jsonObject = JSON.parse(value);
			for (var i = 0; i < jsonObject.length; i++)
			{
				array.push(jsonObject[i]);
			}
		}
		else
		{

			var targetName = tempSheetIdentifierPreamble + sheet.getSheetId() + tempSheetIdentifierPostamble + row + "," + col;

			var newSheet = deserializeSheet(targetName, value);

			var directionIsHorizontal = detectDirection(newSheet);

			var newObject = createObject(newSheet, entryName, directionIsHorizontal);

			if (newObject)
			{
				object[entryName] = newObject;

				var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
				activeSpreadsheet.deleteSheet(newSheet);
			}

			var spreadsheet = SpreadsheetApp.getActive();
			spreadsheet.setActiveSheet(sheet);

			for (var i = 0; i < newObject[entryName].length; i++)
			{
				array.push(newObject[entryName][i]);
			}
		}

	}

}

function isAllArrays(sheet, directionIsHorizontal)
{

	var result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	if (result["validRowsIndices"].length > 0)
	{
		object = {}

		var totalRows = sheet.getLastRow();
		var totalCols = sheet.getLastColumn();

		var values = sheet.getSheetValues(1, 1, totalRows, totalCols);

		var firstValidRow = result["validRowsIndices"][0];
		var firstValidCol = result["validColsIndices"][0];

		var totalValidRows = result["validRowsIndices"].length;
		var totalValidCols = result["validColsIndices"].length;

		var marker = directionIsHorizontal ? totalValidCols : totalValidRows;

		for (var i = 0; i < marker; i++)
		{
			var index = directionIsHorizontal ? result["validColsIndices"][i] : result["validRowsIndices"][i];
			var celTypeAndName = getCelTypeAndName(directionIsHorizontal ? values[firstValidRow][index] : values[index][firstValidCol]);
			var celType = celTypeAndName["type"];
			var celName = celTypeAndName["name"];

			var isArray = isArrayType(celType);

			if (!isArray)
			{

				return false;
			}
		}
		return true;
	}
	return false;
}

function createObject(sheet, name, directionIsHorizontal, isRoot)
{
	var result = getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal);

	var emptyLineEndsArray = false;

	var object = null;

	if (result["validRowsIndices"].length > 0)
	{
		object = {}

		var firstValidRow = result["validRowsIndices"][0];
		var firstValidCol = result["validColsIndices"][0];
		var lastValidCol = result["validColsIndices"][result["validColsIndices"].length - 1];
		var lastValidRow = result["validRowsIndices"][result["validRowsIndices"].length - 1];

		var lastValidRowAndNonEmptyRow = getLastValidRowAndNonEmptyRow(emptyLinesAllowed, sheet, directionIsHorizontal);
		var lastNonEmptyRow = lastValidRowAndNonEmptyRow["lastNonEmptyRow"];
		var lastNonEmptyCol = lastValidRowAndNonEmptyRow["lastNonEmptyCol"];

		object[name] = isRoot && !isRootObjectArray ?
		{} : [];

		var totalRows = sheet.getLastRow();
		var totalCols = sheet.getLastColumn();

		var values = sheet.getSheetValues(1, 1, totalRows, totalCols);

		var currentObject = null;

		var allArrays = isAllArrays(sheet, directionIsHorizontal);

		var currentObjectEmptyArrayEntriesFound = []

		var start = directionIsHorizontal ? firstValidRow + 1 : firstValidCol + 1;
		var end = directionIsHorizontal ? lastNonEmptyRow : lastNonEmptyCol;

		for (var i = start; i <= end; i++)
		{
			if (isEmptyRowOrCol(i, directionIsHorizontal, sheet) && !allArrays)
			{

				continue;
			}

			var newObjectStarting = true;

			// detects if this line is starting a new object
			if (currentObject != null)
			{
				//Browser.msgBox('Result', "has object " + " " + i + " "  + values[i][firstValidCol], Browser.Buttons.OK); 
				newObjectStarting = false;

				var start2 = directionIsHorizontal ? firstValidCol : firstValidRow;
				var end2 = directionIsHorizontal ? lastValidCol : lastValidRow;

				for (var j = start2; j < end2; j++)
				{
					var celTypeAndName = getCelTypeAndName(directionIsHorizontal ? values[firstValidRow][j] : values[j][firstValidCol]);
					var celType = celTypeAndName["type"];
					var celName = celTypeAndName["name"];
					var isArray = isArrayType(celType);
					var basicType = celType;
					var value = directionIsHorizontal ? values[i][j] : values[j][i];

					var isEmpty = value == "";
					var objStr = JSON.stringify(currentObject);
					var simpleEntryAlreadyInput = (!isArray && currentObject[celName] != null && !isEmpty);
					var arrayHadAlreadyEmptyEntryInAllArraySetup = (allArrays && currentObjectEmptyArrayEntriesFound.indexOf(j) > -1 && !isEmpty)

					if (simpleEntryAlreadyInput || arrayHadAlreadyEmptyEntryInAllArraySetup)
					{
						//Browser.msgBox('Result', "WILL START NEW OBJECT " + " " + i + " "  + values[i][j] + " is array " + isArray + " " + currentObject[colName], Browser.Buttons.OK); 
						newObjectStarting = true;
						break;
					}

				}
			}

			// end of detecting if a new object is been started

			if (newObjectStarting)
			{
				var newObjectToReplace = {};
				if (currentObject != null && isRoot && !isRootObjectArray)
				{
					object[name] = newObjectToReplace;
				}
				else
				{
					object[name].push(newObjectToReplace);
				}

				currentObject = newObjectToReplace;
				currentObjectEmptyArrayEntriesFound = [];
			}

			var start2 = directionIsHorizontal ? firstValidCol : firstValidRow;
			var end2 = directionIsHorizontal ? lastValidCol : lastValidRow;

			for (var j = start2; j <= end2; j++)
			{
				var celTypeAndName = getCelTypeAndName(directionIsHorizontal ? values[firstValidRow][j] : values[j][firstValidCol]);
				var celType = celTypeAndName["type"];
				var celName = celTypeAndName["name"];
				var isArray = isArrayType(celType);
				var basicType = celType;
				var value = "" + (directionIsHorizontal ? values[i][j] : values[j][i]);

				var isEmpty = value == "";

				if (isArray)
				{
					basicType = getArrayBasicType(celType);

					if (currentObject[celName] == null)
					{
						currentObject[celName] = [];
					}

					if (!isEmpty)
					{
						pushValueIntoArray(currentObject[celName], celName, basicType, value, sheet, directionIsHorizontal ? i : j, directionIsHorizontal ? j : i);
					}
					else if (allArrays)
					{
						currentObjectEmptyArrayEntriesFound.push(j);
					}
				}
				else
				{
					if (!isEmpty)
					{
						//Browser.msgBox('Result', "pushin " + " " + colName + " "  + value , Browser.Buttons.OK); 
						parseValueIntoObject(currentObject, celName, basicType, value, sheet, directionIsHorizontal ? i : j, directionIsHorizontal ? j : i);
					}
				}

				//Browser.msgBox('Result', colTypeAndName["type"] + " " + values[firstValidRow][j], Browser.Buttons.OK);

			}

		}
	}



	return object;

}


function getValidAndInvalidColumnsWithJsonTypes(sheet, directionIsHorizontal)
{
	validRowsIndices = [];
	invalidRowsIndices = [];

	validColsIndices = [];
	invalidColsIndices = [];

	var totalRows = sheet.getLastRow();
	var totalCols = sheet.getLastColumn();
	var values = sheet.getSheetValues(1, 1, sheet.getLastRow(), sheet.getLastColumn());
	var lineOrColIndexWithFirstJsonType = -1;

	var knownTypes = [];

	for (var i = 0; i < (directionIsHorizontal ? totalRows : totalCols); i++)
	{
		for (var j = 0; j < (directionIsHorizontal ? totalCols : totalRows); j++)
		{
			var entryValue = "" + (directionIsHorizontal ? values[i][j] : values[j][i]);
			entryValue = entryValue.toLowerCase();

			if (isJSONType(entryValue))
			{
				if (lineOrColIndexWithFirstJsonType == -1)
				{
					lineOrColIndexWithFirstJsonType = i;
				}

				if (lineOrColIndexWithFirstJsonType == i && knownTypes.indexOf(entryValue) == -1)
				{
					knownTypes.push(entryValue);
					validRowsIndices.push((directionIsHorizontal ? i : j));
					validColsIndices.push((directionIsHorizontal ? j : i));
				}
				else
				{
					invalidRowsIndices.push((directionIsHorizontal ? i : j));
					invalidColsIndices.push((directionIsHorizontal ? j : i));
				}
			}
			else if (entryValue != "" && lineOrColIndexWithFirstJsonType == i)
			{
				invalidRowsIndices.push((directionIsHorizontal ? i : j));
				invalidColsIndices.push((directionIsHorizontal ? j : i));
			}

		}


	}

	var returnObj = {};
	returnObj["validRowsIndices"] = validRowsIndices;
	returnObj["invalidRowsIndices"] = invalidRowsIndices;
	returnObj["validColsIndices"] = validColsIndices;
	returnObj["invalidColsIndices"] = invalidColsIndices;

	//printObj(returnObj);

	return returnObj;
}

function importJSON()
{
	// Options
	// Array separator, default ','
	// Arrays in multiple lines
	// Arrays in multiple lines if below X entries, 50 default
}


function createSheetFromJSON(jsonStr)
{

}

function validateJSON(jsonStr)
{
	return true;
}