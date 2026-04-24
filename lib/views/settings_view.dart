import 'dart:math';

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
            const Text(
              'LTO DIGITAL UV PRINTING SETTINGS',
              style: TextStyle(
                fontSize: 26,
                fontWeight: FontWeight.w900,
                color: Color(0xFF1E3A5F),
              ),
            ),
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
                        _buildActionButtonRow(
                          'Open Output Directory',
                          'OPEN',
                          () async {
                            final docs =
                                await getApplicationDocumentsDirectory();
                            Process.run('explorer', [
                              '${docs.path}\\PlakaMatik Files\\Outputs',
                            ]);
                          },
                        ),
                        _buildActionButtonRow(
                          'Open Logs Folder',
                          'OPEN',
                          () async {
                            final docs =
                                await getApplicationDocumentsDirectory();
                            Process.run('explorer', [
                              '${docs.path}\\PlakaMatik Files\\Logs',
                            ]);
                          },
                        ),
                        _buildActionButtonRow(
                          'Template Path Configuration',
                          'OPEN',
                          () async {
                            Process.run('explorer', [
                              '${Directory.current.path}\\python_engine\\Core\\CorelDRAW Templates',
                            ]);
                          },
                        ),
                        const SizedBox(height: 30),
                        _buildSectionHeader('Maintenance & Configuration'),
                        _buildActionButtonRow(
                          'Delete Temporary Files',
                          'DELETE',
                          () {},
                          color: Colors.red,
                        ),
                        _buildActionButtonRow(
                          'Save Configurations to Engine',
                          'SAVE',
                          () {
                            viewModel.exportJsonHandshake();
                            ScaffoldMessenger.of(context).showSnackBar(
                              const SnackBar(
                                content: Text(
                                  "Configurations pushed to Python Bridge!",
                                ),
                              ),
                            );
                          },
                          color: Colors.green,
                        ),
                        _buildActionButtonRow(
                          'Reset System Defaults',
                          'RESET',
                          () {
                            viewModel.resetJsonHandshake();
                            ScaffoldMessenger.of(context).showSnackBar(
                              const SnackBar(
                                content: Text("System Defaults restored."),
                              ),
                            );
                          },
                          color: Colors.red,
                        ),
                      ],
                    ),
                  ),
                  const SizedBox(width: 40),
                  // Right Column
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        _buildSectionHeader('Automation & Output Profiles'),
                        _buildPrinterDropdownRow(
                          'Hardware Output Queue',
                          viewModel,
                        ),
                        _buildSwitchRow(
                          'Force CMYK Hard-Separation',
                          viewModel.isCmyk,
                          (v) => viewModel.updateCmyk(v),
                        ),
                        _buildSwitchRow(
                          'Show Hidden Backend Execution Processes',
                          viewModel.isVisibleCorel,
                          (v) => viewModel.updateVisibleCorel(v),
                        ),
                        const SizedBox(height: 10),
                        _buildStatusRow(
                          'Font Integrity Check (Edigna Medium)',
                          'INSTALLED',
                          color: Colors.green,
                        ),
                        _buildActionButtonRow(
                          'Kill CorelDRAW Processes',
                          'STOP',
                          () {
                            Process.run('taskkill', [
                              '/F',
                              '/IM',
                              'CorelDRW.exe',
                            ]);
                          },
                          color: Colors.red,
                        ),
                        const SizedBox(height: 30),
                        _buildSectionHeader('Calibration & Offset Adjustments'),
                        _buildOffsetCalibrationRow(
                          'Adjust X-Offset (mm)',
                          viewModel,
                          isX: true,
                        ),
                        _buildOffsetCalibrationRow(
                          'Adjust Y-Offset (mm)',
                          viewModel,
                          isX: false,
                        ),
                        _buildThemeSwitchRow(),
                      ],
                    ),
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildSectionHeader(String title) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Text(
        title,
        style: const TextStyle(
          fontSize: 18,
          fontWeight: FontWeight.bold,
          color: Color(0xFF1E3A5F),
        ),
      ),
    );
  }

  Widget _buildActionButtonRow(
    String label,
    String btnLabel,
    VoidCallback onPressed, {
    Color color = Colors.white,
  }) {
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
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(20),
              ),
              padding: const EdgeInsets.symmetric(horizontal: 20, vertical: 12),
            ),
            onPressed: onPressed,
            child: Text(
              btnLabel,
              style: const TextStyle(fontWeight: FontWeight.bold),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildPrinterDropdownRow(String label, SettingsViewModel vm) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 16)),
          const SizedBox(width: 10),
          Expanded(
            child: Align(
              alignment: Alignment.centerRight,
              child: Container(
                padding: const EdgeInsets.symmetric(
                  horizontal: 10,
                  vertical: 0,
                ),
                decoration: BoxDecoration(
                  color: Colors.white,
                  borderRadius: BorderRadius.circular(10),
                ),
                child: DropdownButton<String>(
                  isExpanded: true,
                  value: vm.availablePrinters.contains(vm.selectedPrinter)
                      ? vm.selectedPrinter
                      : null,
                  underline: const SizedBox(),
                  items: vm.availablePrinters
                      .map(
                        (e) => DropdownMenuItem(
                          value: e,
                          child: Text(e, overflow: TextOverflow.ellipsis),
                        ),
                      )
                      .toList(),
                  onChanged: (val) {
                    if (val != null) vm.updatePrinter(val);
                  },
                ),
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildSwitchRow(
    String label,
    bool value,
    ValueChanged<bool> onChanged,
  ) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Expanded(child: Text(label, style: const TextStyle(fontSize: 16))),
          Switch(
            value: value,
            activeColor: const Color(0xFF1E3A5F),
            onChanged: onChanged,
          ),
        ],
      ),
    );
  }

  Widget _buildStatusRow(
    String label,
    String status, {
    Color color = Colors.green,
  }) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(
            label,
            style: const TextStyle(
              fontSize: 16,
              overflow: TextOverflow.ellipsis,
            ),
          ),
          Container(
            padding: const EdgeInsets.symmetric(horizontal: 15, vertical: 8),
            decoration: BoxDecoration(
              color: color,
              borderRadius: BorderRadius.circular(20),
            ),
            child: Text(
              status,
              style: const TextStyle(
                fontWeight: FontWeight.bold,
                color: Colors.white,
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildOffsetCalibrationRow(
    String label,
    SettingsViewModel vm, {
    required bool isX,
  }) {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          Text(label, style: const TextStyle(fontSize: 16)),
          Container(
            decoration: BoxDecoration(
              color: Colors.white,
              borderRadius: BorderRadius.circular(20),
            ),
            child: Row(
              children: [
                IconButton(
                  icon: const Icon(Icons.remove, size: 18),
                  onPressed: () => isX ? vm.adjustDx(-0.5) : vm.adjustDy(-0.5),
                ),
                Text(
                  isX
                      ? vm.globalDxOffset.toStringAsFixed(1)
                      : vm.globalDyOffset.toStringAsFixed(1),
                  style: const TextStyle(
                    fontWeight: FontWeight.bold,
                    fontSize: 16,
                  ),
                ),
                IconButton(
                  icon: const Icon(Icons.add, size: 18),
                  onPressed: () => isX ? vm.adjustDx(0.5) : vm.adjustDy(0.5),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildThemeSwitchRow() {
    return Padding(
      padding: const EdgeInsets.only(bottom: 15),
      child: Row(
        mainAxisAlignment: MainAxisAlignment.spaceBetween,
        children: [
          const Text(
            'Dark Mode',
            style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold),
          ),
          Container(
            decoration: BoxDecoration(
              color: Colors.white,
              borderRadius: BorderRadius.circular(20),
            ),
            child: Row(
              children: [
                Container(
                  padding: const EdgeInsets.symmetric(
                    horizontal: 20,
                    vertical: 8,
                  ),
                  decoration: BoxDecoration(
                    color: Colors.blueAccent,
                    borderRadius: BorderRadius.circular(20),
                  ),
                  child: const Text(
                    "Light",
                    style: TextStyle(
                      color: Colors.white,
                      fontWeight: FontWeight.bold,
                    ),
                  ),
                ),
                Container(
                  padding: const EdgeInsets.symmetric(
                    horizontal: 20,
                    vertical: 8,
                  ),
                  child: const Text(
                    "Dark",
                    style: TextStyle(fontWeight: FontWeight.bold),
                  ),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }
}
