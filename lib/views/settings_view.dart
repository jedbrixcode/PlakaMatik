import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'dart:io';
import 'package:path_provider/path_provider.dart';
import '../viewmodels/settings_viewmodel.dart';

class SettingsView extends StatelessWidget {
  const SettingsView({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<SettingsViewModel>(context);

    return SingleChildScrollView(
      child: Padding(
        padding: const EdgeInsets.all(25.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text('LTO DIGITAL UV PRINTING SETTINGS', style: TextStyle(fontSize: 26, fontWeight: FontWeight.w900, color: Color(0xFF1E3A5F))),
            const SizedBox(height: 20),
            Container(
              decoration: BoxDecoration(
                color: Colors.white.withOpacity(0.9),
                borderRadius: BorderRadius.circular(15),
                border: Border.all(color: const Color(0xFF3B6B88), width: 3),
              ),
              padding: const EdgeInsets.all(30),
              child: Row(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  // Left Column
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        _buildSectionHeader('Directory & Path Management'),
                        _buildActionButtonRow('Open Output Directory', 'OPEN', () async {
                          final docs = await getApplicationDocumentsDirectory();
                          Process.run('explorer', ['${docs.path}\\PlakaMatik Files\\Outputs']);
                        }),
                        _buildActionButtonRow('Open Logs Folder', 'OPEN', () async {
                          final docs = await getApplicationDocumentsDirectory();
                          Process.run('explorer', ['${docs.path}\\PlakaMatik Files\\Logs']);
                        }),
                        _buildActionButtonRow('Template Path Configuration', 'OPEN', () async {
                           Process.run('explorer', ['${Directory.current.path}\\python_engine\\Core\\CorelDRAW Templates']);
                        }),
                        const SizedBox(height: 30),
                        _buildSectionHeader('Maintenance & Cleanup'),
                        _buildActionButtonRow('Delete Temporary Files', 'DELETE', () {}, color: Colors.red),
                        _buildActionButtonRow('Clear Log History', 'CLEAR', () {}, color: Colors.red),
                        _buildActionButtonRow('Reset System Defaults', 'RESET', () {}, color: Colors.red),
                      ]
                    )
                  ),
                  const SizedBox(width: 40),
                  // Right Column
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        _buildSectionHeader('Automation & Script Tweaks'),
                        _buildDropdownRow('Trial Bypass Delay', ['5 Sec V', '7 Sec V', '10 Sec V']),
                        _buildStatusRow('Font Integrity Check (Edigna Medium)', 'INSTALLED', color: Colors.green),
                        _buildActionButtonRow('Kill CorelDRAW Process', 'STOP', () {
                           Process.run('taskkill', ['/F', '/IM', 'CorelDRW.exe']);
                        }, color: Colors.red),
                        const SizedBox(height: 30),
                        _buildSectionHeader('Calibration & Offset Adjustments'),
                        _buildOffsetCalibrationRow('Adjust X-Offset (mm)', viewModel, isX: true),
                        _buildOffsetCalibrationRow('Adjust Y-Offset (mm)', viewModel, isX: false),
                        _buildThemeSwitchRow(),
                      ]
                    )
                  )
                ]
              )
            )
          ],
        ),
      ),
    );
  }

  Widget _buildSectionHeader(String title) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Text(title, style: const TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Color(0xFF1E3A5F))),
    );
  }

  Widget _buildActionButtonRow(String label, String btnLabel, VoidCallback onPressed, {Color color = Colors.white}) {
    bool isColored = color != Colors.white;
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 16)),
          ElevatedButton(
            style: ElevatedButton.styleFrom(
              backgroundColor: isColored ? color : Colors.white,
              foregroundColor: isColored ? Colors.white : Colors.black,
              shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(20)),
              padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 12),
            ),
            onPressed: onPressed,
            child: Text(btnLabel, style: const TextStyle(fontWeight: FontWeight.bold)),
          )
        ],
      ),
    );
  }

  Widget _buildDropdownRow(String label, List<String> items) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 16)),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 15, vertical: 8),
            decoration: BoxDecoration(color: Colors.white, borderRadius: BorderRadius.circular(20)),
            child: Text(items[1], style: const TextStyle(fontWeight: FontWeight.bold)), // Hardcoded UI matcher for now
          )
        ],
      )
    );
  }

  Widget _buildStatusRow(String label, String status, {Color color = Colors.green}) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 16)),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 15, vertical: 8),
            decoration: BoxDecoration(color: color, borderRadius: BorderRadius.circular(20)),
            child: Text(status, style: const TextStyle(fontWeight: FontWeight.bold, color: Colors.white)), 
          )
        ],
      )
    );
  }

  Widget _buildOffsetCalibrationRow(String label, SettingsViewModel vm, {required bool isX}) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 16)),
          Container(
            decoration: BoxDecoration(color: Colors.white, borderRadius: BorderRadius.circular(20)),
            child: Row(
              children: [
                IconButton(icon: const Icon(Icons.remove, size: 18), onPressed: () {}),
                Text(isX ? vm.globalDxOffset.toString() : vm.globalDyOffset.toString(), style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16)),
                IconButton(icon: const Icon(Icons.add, size: 18), onPressed: () {}),
              ],
            )
          )
        ],
      )
    );
  }

  Widget _buildThemeSwitchRow() {
     return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          const Text('Dark Mode', style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold)),
          Container(
            decoration: BoxDecoration(color: Colors.white, borderRadius: BorderRadius.circular(20)),
            child: Row(
              children: [
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 8),
                  decoration: BoxDecoration(color: Colors.blueAccent, borderRadius: BorderRadius.circular(20)),
                  child: const Text("Light", style: TextStyle(color: Colors.white, fontWeight: FontWeight.bold))
                ),
                Container(
                  padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 8),
                  child: const Text("Dark", style: TextStyle(fontWeight: FontWeight.bold))
                )
              ]
            )
          )
        ],
      )
    );
  }
}
