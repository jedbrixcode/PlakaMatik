import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../viewmodels/batch_viewmodel.dart';
import '../viewmodels/settings_viewmodel.dart';

class BatchInputWidget extends StatefulWidget {
  const BatchInputWidget({super.key});

  @override
  _BatchInputWidgetState createState() => _BatchInputWidgetState();
}

class _BatchInputWidgetState extends State<BatchInputWidget> {
  final TextEditingController _idController = TextEditingController();
  final TextEditingController _desigController = TextEditingController();
  String _selectedType = 'MV';

  @override
  void dispose() {
    _idController.dispose();
    _desigController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<BatchViewModel>(context);
    final settings = Provider.of<SettingsViewModel>(context);

    return Container(
      padding: const EdgeInsets.all(20),
      decoration: BoxDecoration(
        color: Colors.white.withValues(alpha: 0.9),
        borderRadius: BorderRadius.circular(15),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.stretch,
        children: [
          const Text(
            'Plate Data Entry',
            style: TextStyle(
              fontWeight: FontWeight.bold,
              fontSize: 18,
              color: Color(0xFF3B6B88),
            ),
          ),
          const SizedBox(height: 15),
          if (_selectedType == 'MV')
            TextField(
              controller: _idController,
              decoration: InputDecoration(
                labelText: 'Plate Identifier',
                filled: true,
                fillColor: Colors.grey[100],
                border: OutlineInputBorder(
                  borderRadius: BorderRadius.circular(8),
                  borderSide: BorderSide.none,
                ),
              ),
            ),
          if (_selectedType == 'MV') const SizedBox(height: 15),
          TextField(
            controller: _desigController,
            decoration: InputDecoration(
              labelText: 'Designation / Region',
              filled: true,
              fillColor: Colors.grey[100],
              border: OutlineInputBorder(
                borderRadius: BorderRadius.circular(8),
                borderSide: BorderSide.none,
              ),
            ),
          ),
          const SizedBox(height: 15),
          DropdownButtonFormField<String>(
            value: _selectedType,
            decoration: InputDecoration(
              labelText: 'Plate Type',
              filled: true,
              fillColor: Colors.grey[100],
              border: OutlineInputBorder(
                borderRadius: BorderRadius.circular(8),
                borderSide: BorderSide.none,
              ),
            ),
            items: ['MV', 'MC']
                .map(
                  (String value) => DropdownMenuItem<String>(
                    value: value,
                    child: Text(value),
                  ),
                )
                .toList(),
            onChanged: (newValue) {
              if (newValue != null) {
                setState(() => _selectedType = newValue);
              }
            },
          ),
          const SizedBox(height: 20),
          ElevatedButton.icon(
            style: ElevatedButton.styleFrom(
              backgroundColor: const Color(0xFF1E6083),
              padding: const EdgeInsets.symmetric(vertical: 20),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(10),
              ),
            ),
            icon: const Icon(Icons.add, color: Colors.white),
            label: const Text(
              'Add to Queue',
              style: TextStyle(
                color: Colors.white,
                fontSize: 16,
                fontWeight: FontWeight.bold,
              ),
            ),
            onPressed: () {
              bool isValid = _desigController.text.isNotEmpty;
              if (_selectedType == 'MV' && _idController.text.isEmpty) {
                isValid = false;
              }
              if (isValid) {
                viewModel.addToQueue(
                  _idController.text,
                  _desigController.text,
                  _selectedType,
                );
                _idController.clear();
                // We purposefully leave _desigController uncleared for quick batching
              }
            },
          ),
          const SizedBox(height: 10),
          // PREVIEW EXPORT GENERATOR BUTTON
          ElevatedButton.icon(
            style: ElevatedButton.styleFrom(
              backgroundColor: const Color(0xFF4A90E2),
              padding: const EdgeInsets.symmetric(vertical: 25),
              shape: RoundedRectangleBorder(
                borderRadius: BorderRadius.circular(10),
              ),
            ),
            icon: viewModel.isProcessing
                ? const SizedBox(
                    width: 24,
                    height: 24,
                    child: CircularProgressIndicator(
                      color: Colors.white,
                      strokeWidth: 2,
                    ),
                  )
                : const Icon(Icons.build, color: Colors.white),
            label: Text(
              viewModel.isProcessing
                  ? 'Generating PDF Engine...'
                  : 'STAGE 1: GENERATE BATCH PREVIEW',
              style: const TextStyle(
                color: Colors.white,
                fontSize: 14,
                fontWeight: FontWeight.bold,
                letterSpacing: 1.2,
              ),
            ),
            onPressed:
                viewModel.isProcessing ||
                        viewModel.printQueue.isEmpty ||
                        viewModel.currentRunIndex >= viewModel.printQueue.length
                    ? null
                    : () => viewModel.generateNextPreviewChunk(settings),
          ),
        ],
      ),
    );
  }
}
