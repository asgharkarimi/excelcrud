import 'package:flutter/material.dart';
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:path_provider/path_provider.dart';
import 'dart:io';

Future<Excel> readExcelFile() async {
  FilePickerResult? result = await FilePicker.platform.pickFiles(
    type: FileType.custom,
    allowedExtensions: ['xlsx'],
  );

  if (result != null) {
    File file = File(result.files.single.path!);
    var bytes = file.readAsBytesSync();
    var excel = Excel.decodeBytes(bytes);
    return excel;
  } else {
    throw Exception("File not picked");
  }
}

void addRow(Excel excel) {
  var sheet = excel['Sheet1'];
  sheet.appendRow([
    '4',
    '1234567890',
    '09123456789',
    'New',
    'User',
    '13600101',
    '14000101',
    'ندارد',
    'ندارد',
    '0',
    '0',
    'بلي',
    'تکميل نشده',
    'فاقد',
    '-',
    'زنجان',
    '-',
    'زنجان',
    'زنجان',
    'زنجان',
    'زنجان'
  ]);
}

void updateRow(Excel excel, int rowIndex) {
  var sheet = excel['Sheet1'];
  sheet.updateCell(
      CellIndex.indexByString("A${rowIndex + 1}"), 'Updated Value');
}

class ExcelCRUDPage extends StatefulWidget {
  @override
  _ExcelCRUDPageState createState() => _ExcelCRUDPageState();
}

class _ExcelCRUDPageState extends State<ExcelCRUDPage> {
  Excel? excel;

  @override
  void initState() {
    super.initState();
    loadExcelFile();
  }

  Future<void> loadExcelFile() async {
    excel = await readExcelFile();
    setState(() {});
  }

  void deleteRow(Excel excel, int rowIndex) {
    var sheet = excel['Sheet1'];
    sheet.removeRow(rowIndex);
  }

  Future<void> saveExcelFile(Excel excel) async {
    Directory directory = await getApplicationDocumentsDirectory();
    String path = '${directory.path}/appiiiitest_updated.xlsx';
    File file = File(path);
    file.writeAsBytesSync(excel.encode()!);
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Excel CRUD Operations'),
      ),
      body: excel == null
          ? Center(child: CircularProgressIndicator())
          : ListView.builder(
              itemCount: excel!.tables['Sheet1']!.rows.length,
              itemBuilder: (context, index) {
                var row = excel!.tables['Sheet1']!.rows[index];
                return ListTile(
                  title: Text(row[3].toString()),
                  subtitle: Text(row[4].toString()),
                  trailing: Row(
                    mainAxisSize: MainAxisSize.min,
                    children: [
                      IconButton(
                        icon: Icon(Icons.edit),
                        onPressed: () {
                          updateRow(excel!, index);
                          saveExcelFile(excel!);
                          setState(() {});
                        },
                      ),
                      IconButton(
                        icon: Icon(Icons.delete),
                        onPressed: () {
                          deleteRow(excel!, index);
                          saveExcelFile(excel!);
                          setState(() {});
                        },
                      ),
                    ],
                  ),
                );
              },
            ),
      floatingActionButton: FloatingActionButton(
        onPressed: () {
          addRow(excel!);
          saveExcelFile(excel!);
          setState(() {});
        },
        child: Icon(Icons.add),
      ),
    );
  }
}

class ExcelListViewPage extends StatefulWidget {
  @override
  _ExcelListViewPageState createState() => _ExcelListViewPageState();
}

class _ExcelListViewPageState extends State<ExcelListViewPage> {
  List<List<dynamic>> _data = [];

  @override
  void initState() {
    super.initState();
    loadExcelData();
  }

  Future<List<List<dynamic>>> readExcelFile() async {
    FilePickerResult? result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (result != null) {
      File file = File(result.files.single.path!);
      var bytes = file.readAsBytesSync();
      var excel = Excel.decodeBytes(bytes);
      var sheet = excel.tables['Sheet1']; // Assuming the sheet name is 'Sheet1'

      // Extract data from the sheet and return it as a List<List<dynamic>>
      List<List<dynamic>> rows = [];
      for (var row in sheet!.rows) {
        // Convert each row to a List<dynamic>
        List<dynamic> rowData = [];
        for (var cell in row) {
          rowData.add(cell?.value); // Add the cell value to the row data
        }
        rows.add(rowData); // Add the row to the list of rows
      }

      return rows;
    } else {
      throw Exception("File not picked");
    }
  }

  Future<void> loadExcelData() async {
    try {
      List<List<dynamic>> data = await readExcelFile();
      setState(() {
        _data = data;
      });
    } catch (e) {
      print("Error loading Excel file: $e");
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Excel Data in ListView'),
      ),
      body: _data.isEmpty
          ? Center(child: CircularProgressIndicator())
          : ListView.builder(
              itemCount: _data.length,
              itemBuilder: (context, index) {
                var row = _data[index];
                return Card(
                  margin: EdgeInsets.all(8.0),
                  child: ListTile(
                    title: Text(row[3].toString()),
                    // Displaying the 'نام' column
                    subtitle: Text(row[4].toString()),
                    // Displaying the 'نام خانوادگی' column
                    trailing: Text(
                        row[1].toString()), // Displaying the 'کد ملی' column
                  ),
                );
              },
            ),
      floatingActionButton: FloatingActionButton(
        onPressed: () async {
          await loadExcelData();
        },
        child: Icon(Icons.refresh),
      ),
    );
  }
}

void main() {
  runApp(MaterialApp(
    home: ExcelListViewPage(),
  ));
}
