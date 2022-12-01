import 'dart:io';
import 'package:travel_pro/locale/Languages/ar.dart';
import 'package:excel/excel.dart';
import 'package:path_provider/path_provider.dart';

mapToExcel()async {
  print("init Map to Excel converter");
  var excel = Excel.createExcel();
  var sheet = excel['arabic'];

  print("Loaded Map file");

  for(int i = 1 ; i < arabic().length ; i++){
    var cell = sheet.cell(CellIndex.indexByString("A$i"));
    cell.value = arabic().keys.toList()[i-0];
    var cell2 = sheet.cell(CellIndex.indexByString("B$i"));
    cell2.value = arabic().values.toList()[i-0];
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
  print("file saved successfully");
}
