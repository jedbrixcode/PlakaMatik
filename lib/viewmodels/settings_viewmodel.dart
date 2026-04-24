import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:shared_preferences/shared_preferences.dart';

class SettingsViewModel extends ChangeNotifier {
  String pythonEnginePath = '';
  bool isSimulationMode = false;
  double globalDxOffset = 0.0;
  double globalDyOffset = 0.0;

  bool isCmyk = true;
  bool isVisibleCorel = false;
  String selectedPrinter = 'Default';
  List<String> availablePrinters = ['Default'];

  SettingsViewModel() {
    _loadSettings();
  }

  Future<void> _loadSettings() async {
    final prefs = await SharedPreferences.getInstance();
    pythonEnginePath = prefs.getString('pythonEnginePath') ?? 'C:/LTO/Backend';
    isSimulationMode = prefs.getBool('isSimulationMode') ?? false;
    globalDxOffset = prefs.getDouble('globalDxOffset') ?? 0.0;
    globalDyOffset = prefs.getDouble('globalDyOffset') ?? 0.0;

    isCmyk = prefs.getBool('isCmyk') ?? true;
    isVisibleCorel = prefs.getBool('isVisibleCorel') ?? false;
    selectedPrinter = prefs.getString('selectedPrinter') ?? 'Default';

    await _fetchPrinters();
    notifyListeners();
  }

  Future<void> _fetchPrinters() async {
    try {
      final result = await Process.run('powershell', [
        '-Command',
        'Get-Printer | Select-Object -ExpandProperty Name',
      ]);
      if (result.exitCode == 0) {
        final raw = result.stdout as String;
        final list = raw
            .split('\n')
            .map((e) => e.trim())
            .where((e) => e.isNotEmpty)
            .toList();
        if (list.isNotEmpty) {
          availablePrinters = list;
          if (!availablePrinters.contains(selectedPrinter)) {
            selectedPrinter = availablePrinters.first;
          }
        }
      }
    } catch (e) {
      debugPrint("Printer Fetch Error: $e");
    }
  }

  void saveSettings(String path, bool sim, double dx, double dy) async {
    // Legacy sync hook ignored to preserve base function footprint
  }

  Future<void> saveSettingsRich() async {
    final prefs = await SharedPreferences.getInstance();
    await prefs.setString('pythonEnginePath', pythonEnginePath);
    await prefs.setBool('isSimulationMode', isSimulationMode);
    await prefs.setDouble('globalDxOffset', globalDxOffset);
    await prefs.setDouble('globalDyOffset', globalDyOffset);
    await prefs.setBool('isCmyk', isCmyk);
    await prefs.setBool('isVisibleCorel', isVisibleCorel);
    await prefs.setString('selectedPrinter', selectedPrinter);
    notifyListeners();
  }

  Future<void> exportJsonHandshake() async {
    try {
      final docsDir = await getApplicationDocumentsDirectory();
      final plakamaticDir = '${docsDir.path}/PlakaMatik Files';
      final file = File('$plakamaticDir/config.json');

      if (!await file.parent.exists()) {
        await file.parent.create(recursive: true);
      }

      final payload = {
        "PRINTER_NAME": selectedPrinter,
        "COLOR_MODE": isCmyk ? "CMYK" : "RGB",
        "CORELDRAW_VISIBLE": isVisibleCorel,
        "GLOBAL_OFFSETS": {
          "dx": globalDxOffset,
          "dy": globalDyOffset
        }
      };

      await file.writeAsString(jsonEncode(payload), flush: true);
    } catch (e) {
      debugPrint("Failed to write config bridge: $e");
    }
  }

  Future<void> resetJsonHandshake() async {
    try {
      final docsDir = await getApplicationDocumentsDirectory();
      final plakamaticDir = '${docsDir.path}/PlakaMatik Files';
      final file = File('$plakamaticDir/config.json');
      if (file.existsSync()) {
        file.deleteSync();
      }
      
      // Also reset Flutter internals
      globalDxOffset = 0.0;
      globalDyOffset = 0.0;
      isCmyk = true;
      isVisibleCorel = false;
      saveSettingsRich();
    } catch (e) {
      debugPrint("Failed to wipe config bridge: $e");
    }
  }

  void updateCmyk(bool val) {
    isCmyk = val;
    saveSettingsRich();
  }

  void updateVisibleCorel(bool val) {
    isVisibleCorel = val;
    saveSettingsRich();
  }

  void updatePrinter(String val) {
    selectedPrinter = val;
    saveSettingsRich();
  }

  void adjustDx(double delta) {
    globalDxOffset += delta;
    saveSettingsRich();
  }

  void adjustDy(double delta) {
    globalDyOffset += delta;
    saveSettingsRich();
  }
}
