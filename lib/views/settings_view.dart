import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../viewmodels/settings_viewmodel.dart';

class SettingsView extends StatelessWidget {
  const SettingsView({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<SettingsViewModel>(context);

    return Padding(
      padding: const EdgeInsets.all(25.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text('System Settings', style: TextStyle(fontSize: 26, fontWeight: FontWeight.w900, color: Color(0xFF1E3A5F))),
          const SizedBox(height: 20),
          Expanded(
            child: GridView.count(
              crossAxisCount: 2,
              crossAxisSpacing: 20,
              mainAxisSpacing: 20,
              childAspectRatio: 2.0,
              children: [
                _buildCard('Python Engine Path', Icons.folder, TextField(
                    controller: TextEditingController(text: viewModel.pythonEnginePath),
                    onChanged: (val) => viewModel.saveSettings(val, viewModel.isSimulationMode, viewModel.globalDxOffset, viewModel.globalDyOffset),
                    decoration: const InputDecoration(border: OutlineInputBorder(), hintText: 'Absolute path to python engine root')
                )),
                _buildCard('Automation Tweaks', Icons.memory, SwitchListTile(
                   title: const Text('Simulation Mode (No Print)'),
                   value: viewModel.isSimulationMode,
                   onChanged: (val) => viewModel.saveSettings(viewModel.pythonEnginePath, val, viewModel.globalDxOffset, viewModel.globalDyOffset),
                )),
                _buildCard('Global DX Offset Calibration', Icons.compare_arrows, Row(
                   mainAxisAlignment: MainAxisAlignment.center,
                   children: [
                     IconButton(icon: const Icon(Icons.remove_circle, color: Colors.red), onPressed: () => viewModel.saveSettings(viewModel.pythonEnginePath, viewModel.isSimulationMode, viewModel.globalDxOffset - 1.0, viewModel.globalDyOffset)),
                     Text('${viewModel.globalDxOffset.toStringAsFixed(1)} mm', style: const TextStyle(fontSize: 24, fontWeight: FontWeight.bold)),
                     IconButton(icon: const Icon(Icons.add_circle, color: Colors.green), onPressed: () => viewModel.saveSettings(viewModel.pythonEnginePath, viewModel.isSimulationMode, viewModel.globalDxOffset + 1.0, viewModel.globalDyOffset)),
                   ],
                )),
                _buildCard('Global DY Offset Calibration', Icons.height, Row(
                   mainAxisAlignment: MainAxisAlignment.center,
                   children: [
                     IconButton(icon: const Icon(Icons.remove_circle, color: Colors.red), onPressed: () => viewModel.saveSettings(viewModel.pythonEnginePath, viewModel.isSimulationMode, viewModel.globalDxOffset, viewModel.globalDyOffset - 1.0)),
                     Text('${viewModel.globalDyOffset.toStringAsFixed(1)} mm', style: const TextStyle(fontSize: 24, fontWeight: FontWeight.bold)),
                     IconButton(icon: const Icon(Icons.add_circle, color: Colors.green), onPressed: () => viewModel.saveSettings(viewModel.pythonEnginePath, viewModel.isSimulationMode, viewModel.globalDxOffset, viewModel.globalDyOffset + 1.0)),
                   ],
                )),
              ],
            ),
          )
        ],
      ),
    );
  }

  Widget _buildCard(String title, IconData icon, Widget child) {
    return Card(
      elevation: 2,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(15)),
      child: Padding(
        padding: const EdgeInsets.all(20.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(children: [Icon(icon, color: const Color(0xFF3B6B88)), const SizedBox(width: 10), Text(title, style: const TextStyle(fontSize: 16, fontWeight: FontWeight.bold, color: Color(0xFF3B6B88)))]),
            const SizedBox(height: 20),
            Expanded(child: Center(child: child)),
          ],
        ),
      ),
    );
  }
}
