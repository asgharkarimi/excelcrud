import 'package:flutter/material.dart';
import 'package:excel/excel.dart' as excel; // Use prefix 'excel'
import 'package:flutter/services.dart'; // For rootBundle
import 'package:flutter_localizations/flutter_localizations.dart';

void main() {
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: TextDirection.rtl, // Set app direction to RTL
      child: MaterialApp(
        title: 'Excel Data App',
        localizationsDelegates: [
          GlobalMaterialLocalizations.delegate,
          GlobalWidgetsLocalizations.delegate,
          GlobalCupertinoLocalizations.delegate,
        ],
        supportedLocales: [
          Locale('en'), // English
          Locale('es'), // Spanish
        ],
        theme: ThemeData(
          fontFamily: 'vazir',
          primarySwatch: Colors.blue,
          visualDensity: VisualDensity.adaptivePlatformDensity,
          appBarTheme: AppBarTheme(
            elevation: 0,
            centerTitle: true,
          ),
        ),
        home: ExcelFromAssetsPage(),
        debugShowCheckedModeBanner: false,
      ),
    );
  }
}

class ExcelFromAssetsPage extends StatefulWidget {
  @override
  _ExcelFromAssetsPageState createState() => _ExcelFromAssetsPageState();
}

class _ExcelFromAssetsPageState extends State<ExcelFromAssetsPage> {
  List<List<dynamic>> _data = []; // Original data from Excel
  List<List<dynamic>> _filteredData = []; // Filtered data for display
  TextEditingController _searchController = TextEditingController(); // Controller for search bar

  @override
  void initState() {
    super.initState();
    loadExcelData();
  }

  Future<void> loadExcelData() async {
    try {
      // Load the Excel file from the assets folder
      ByteData data = await rootBundle.load('assets/list_zanjan.xlsx');
      var bytes = data.buffer.asUint8List();

      // Decode the Excel file using the prefix 'excel'
      var excelFile = excel.Excel.decodeBytes(bytes); // Use 'excel.Excel'
      var sheet = excelFile.tables['Sheet1']; // Assuming the sheet name is 'Sheet1'

      if (sheet == null) {
        print("Sheet 'Sheet1' not found in the Excel file.");
        return;
      }

      // Extract data from the sheet
      List<List<dynamic>> rows = [];
      for (var row in sheet.rows) {
        List<dynamic> rowData = [];
        for (var cell in row) {
          rowData.add(cell?.value); // Add the cell value to the row data
        }
        rows.add(rowData); // Add the row to the list of rows
      }

      setState(() {
        _data = rows;
        _filteredData = rows; // Initialize filtered data with all rows
      });
    } catch (e) {
      print("Error loading Excel file: $e");
    }
  }

  void _filterData(String query) {
    setState(() {
      _filteredData = _data.where((row) {
        // Search in the 'نام' (column 3), 'نام خانوادگی' (column 4),
        // 'دهستان' (column 19), and 'نام روستا' (column 20)
        String name = row[3].toString().toLowerCase();
        String lastName = row[4].toString().toLowerCase();
        String dehestan = row[19].toString().toLowerCase(); // دهستان
        String villageName = row[20].toString().toLowerCase(); // نام روستا

        return name.contains(query.toLowerCase()) ||
            lastName.contains(query.toLowerCase()) ||
            dehestan.contains(query.toLowerCase()) ||
            villageName.contains(query.toLowerCase());
      }).toList();
    });
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Excel Data from Assets'),
      ),
      body: Column(
        children: [
          // Search Bar with RTL direction
          Padding(
            padding: const EdgeInsets.all(8.0),
            child: TextField(
              controller: _searchController,
              textAlign: TextAlign.right, // Align text to the right
              decoration: InputDecoration(
                hintText: 'جستجو بر اساس نام، نام خانوادگی، دهستان، یا نام روستا...',
                hintStyle: TextStyle(fontSize: 14),
                prefixIcon: Icon(Icons.search),
                border: OutlineInputBorder(
                  borderRadius: BorderRadius.circular(10.0),
                ),
              ),
              onChanged: _filterData, // Filter data as the user types
            ),
          ),
          // ListView
          Expanded(
            child: _filteredData.isEmpty
                ? Center(child: CircularProgressIndicator())
                : ListView.builder(
              itemCount: _filteredData.length,
              itemBuilder: (context, index) {
                var row = _filteredData[index];
                // Check if the item is the first one (index == 0)
                if (index == 0) {
                  // Render the first item without onTap functionality
                  return Card(
                    margin: EdgeInsets.all(8.0),
                    child: ListTile(
                      title: Directionality(
                        textDirection: TextDirection.rtl, // Set RTL for the row
                        child: Row(
                          children: [
                            Expanded(
                              flex: 2,
                              child: Text(
                                row[3].toString(), // Displaying the 'نام' column
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                            ),
                            Expanded(
                              flex: 2,
                              child: Text(
                                row[4].toString(), // Displaying the 'نام خانوادگی' column
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                            ),
                            Expanded(
                              flex: 2,
                              child: Text(
                                row[1].toString(), // Displaying the 'کد ملی' column
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                            ),
                          ],
                        ),
                      ),
                      // Disable onTap for the first item
                      onTap: null,
                    ),
                  );
                } else {
                  // Render other items with onTap functionality
                  return Card(
                    margin: EdgeInsets.all(8.0),
                    child: ListTile(
                      title: Directionality(
                        textDirection: TextDirection.rtl, // Set RTL for the row
                        child: Row(
                          children: [
                            Expanded(
                              flex: 2,
                              child: Text(
                                row[3].toString(), // Displaying the 'نام' column
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                            ),
                            Expanded(
                              flex: 2,
                              child: Text(
                                row[4].toString(), // Displaying the 'نام خانوادگی' column
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                            ),
                            Expanded(
                              flex: 2,
                              child: Text(
                                row[1].toString(), // Displaying the 'کد ملی' column
                                style: TextStyle(fontWeight: FontWeight.bold),
                              ),
                            ),
                          ],
                        ),
                      ),
                      onTap: () {
                        // Navigate to the detail page when a row is clicked
                        Navigator.push(
                          context,
                          MaterialPageRoute(
                            builder: (context) => DetailPage(userData: row),
                          ),
                        );
                      },
                    ),
                  );
                }
              },
            ),
          ),
        ],
      ),
    );
  }
}

// Detail Page to show all user data
class DetailPage extends StatelessWidget {
  final List<dynamic> userData;

  DetailPage({required this.userData});

  @override
  Widget build(BuildContext context) {
    return Directionality(
      textDirection: TextDirection.rtl, // Set RTL direction for the detail page
      child: Scaffold(
        appBar: AppBar(
          title: Text('جزئیات کاربر'),
        ),
        body: Padding(
          padding: const EdgeInsets.all(16.0),
          child: ListView(
            children: [
              _buildDetailCard('نام', userData[3].toString()),
              _buildDetailCard('نام خانوادگی', userData[4].toString()),
              _buildDetailCard('کد ملی', userData[1].toString()),
              _buildDetailCard('تاریخ تولد', userData[5].toString()),
              _buildDetailCard('شماره همراه', userData[2].toString()),
              _buildDetailCard('آدرس', userData[17].toString()),
              _buildDetailCard('استان', userData[16].toString()),
              _buildDetailCard('شهرستان', userData[18].toString()),
              _buildDetailCard('دهستان', userData[19].toString()),
              _buildDetailCard('روستا', userData[20].toString()),
            ],
          ),
        ),
      ),
    );
  }

  // Helper method to build a Card for each detail
  Widget _buildDetailCard(String title, String value) {
    return Card(
      elevation: 0, // No shadow
      shape: Border.all(
        color: Colors.grey.withOpacity(0.5), // Grey border with 0.5 opacity
        width: 1.0, // Border width
      ),
      margin: EdgeInsets.only(bottom: 8.0), // Space between cards
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              title,
              style: TextStyle(
                fontSize: 14,
                color: Colors.grey[600], // Grey text for the title
              ),
            ),
            SizedBox(height: 4),
            Text(
              value,
              style: TextStyle(
                fontSize: 16,
                fontWeight: FontWeight.bold,
              ),
            ),
          ],
        ),
      ),
    );
  }
}