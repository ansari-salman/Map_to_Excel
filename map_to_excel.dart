///Created by Ansari Salman
///Add these dependency in pubspec.yaml
///  excel:
///  path_provider:
/// mapToExcel function take Map and produce excel sheet in App data folder
/// excelToMap open excel and produce Map in App data folder

import 'dart:convert';
import 'dart:io';
import 'dart:typed_data';
import 'package:flutter/services.dart';
import 'package:travel_pro/locale/Languages/ar.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';

mapToExcel() async {
  print("init Map to Excel converter");
  var excel = Excel.createExcel();
  var sheet = excel['arabic'];

  print("Loaded Map file");

  for(int i = 0 ; i < arabic().length ; i++){
    var cell = sheet.cell(CellIndex.indexByString("A${i+1}"));
    cell.value = arabic().keys.toList()[i];
    var cell2 = sheet.cell(CellIndex.indexByString("B${i+1}"));
    cell2.value = arabic().values.toList()[i];
  }

  // arabic().forEach((key, value) {
  //
  // });

  excel.setDefaultSheet(sheet.sheetName).then((isSet) {
    // isSet is bool which tells that whether the setting of default sheet is successful or not.
    if (isSet) {
      print("${sheet.sheetName} is set to default sheet.");
    } else {
      print("Unable to set ${sheet.sheetName} to default sheet.");
    }
  });

  print("load path to save the file");

  Directory outputFile = await getApplicationDocumentsDirectory();
  excel.encode().then((onValue) {
    File(outputFile.path + "/arabic.xlsx")
      ..createSync(recursive: true)
      ..writeAsBytesSync(onValue);
  });
  print("file created at ${outputFile.path + "/arabic.xlsx"} successfully");
}

excelTopMap() async {
  print("init Excel to Map converter");
  ByteData data = await rootBundle.load("assets/FileName.xlsx");
  var bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
  var excel = Excel.decodeBytes(bytes);
  print("Loaded Excel file");

  Map<String,String> arabic = new Map();
  for (var table in excel.tables.keys) {
    print(table); //sheet Name
    print(excel.tables[table].maxCols);
    print(excel.tables[table].maxRows);
    for (var row in excel.tables[table].rows) {
      arabic.addAll({row.first: row[1]});
    }
  }
  print("Total values ${excel.tables.length}");
  print("Write to file");
  writeResponseToFile(jsonEncode(arabic));
}

Future<void> writeResponseToFile(String response) async {
final Directory extDir = await getApplicationDocumentsDirectory();
final File file = File(extDir.path + "/api_response.json");
print("File Excel is created in directory "+file.path);
file.writeAsString(response);
}
