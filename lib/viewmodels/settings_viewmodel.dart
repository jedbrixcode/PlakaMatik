import 'package:flutter/material.dart';
import 'package:shared_preferences/shared_preferences.dart';

class SettingsViewModel extends ChangeNotifier {
  String pythonEnginePath = '';
  bool isSimulationMode = false;
  double globalDxOffset = 0.0;
  double globalDyOffset = 0.0;

  SettingsViewModel() {
    _loadSettings();
  }

  Future<void> _loadSettings() async {
    final prefs = await SharedPreferences.getInstance();
    pythonEnginePath = prefs.getString('pythonEnginePath') ?? 'C:/LTO/Backend';
    isSimulationMode = prefs.getBool('isSimulationMode') ?? false;
    globalDxOffset = prefs.getDouble('globalDxOffset') ?? 0.0;
    globalDyOffset = prefs.getDouble('globalDyOffset') ?? 0.0;
    notifyListeners();
  }

  void saveSettings(String path, bool sim, double dx, double dy) async {
    pythonEnginePath = path;
    isSimulationMode = sim;
    globalDxOffset = dx;
    globalDyOffset = dy;
    
    final prefs = await SharedPreferences.getInstance();
    await prefs.setString('pythonEnginePath', path);
    await prefs.setBool('isSimulationMode', sim);
    await prefs.setDouble('globalDxOffset', dx);
    await prefs.setDouble('globalDyOffset', dy);
    
    notifyListeners();
  }
}
