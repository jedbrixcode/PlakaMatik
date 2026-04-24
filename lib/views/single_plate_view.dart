import 'package:flutter/material.dart';
import 'dart:io';
import 'package:syncfusion_flutter_pdfviewer/pdfviewer.dart';
import 'package:path_provider/path_provider.dart';

import 'package:provider/provider.dart';
import '../viewmodels/settings_viewmodel.dart';
import '../services/cleanup_service.dart';
import '../widgets/console_log_widget.dart';

class SinglePlateView extends StatefulWidget {
  const SinglePlateView({super.key});

  @override
  State<SinglePlateView> createState() => _SinglePlateViewState();
}

class _SinglePlateViewState extends State<SinglePlateView> {
  final TextEditingController _middleController = TextEditingController();
  final TextEditingController _identifierController = TextEditingController();
  String _plateType = 'MV';

  bool _isProcessing = false;
  String? _previewPath;
  String? _errorMessage;

  Future<void> _generateSinglePlate(SettingsViewModel settings) async {
    // Only require identifier if it's an MV plate
    if (_middleController.text.isEmpty ||
        (_plateType == 'MV' && _identifierController.text.isEmpty)) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Please fill out all required fields.')),
      );
      return;
    }

    setState(() {
      _isProcessing = true;
      _previewPath = null;
      _errorMessage = null;
    });

    try {
      final docsDir = await getApplicationDocumentsDirectory();
      final plakamaticDir = '${docsDir.path}/PlakaMatik Files';
      final file = File('$plakamaticDir/csv/flutter_user_input.txt');

      if (!await file.parent.exists()) {
        await file.parent.create(recursive: true);
      }

      StringBuffer buffer = StringBuffer();
      buffer.writeln('MIDDLE    IDENTIFIER    TYPE');
      buffer.writeln(
        '${_middleController.text.trim()}    ${_identifierController.text.trim()}    $_plateType',
      );

      final bytes = [0xEF, 0xBB, 0xBF, ...buffer.toString().codeUnits];
      await file.writeAsBytes(bytes);

      final String projectRoot = Directory.current.path;
      final pythonScript = '$projectRoot/python_engine/Core/main.py';
      
      List<String> pyArgs = [pythonScript];
      pyArgs.add(settings.isCmyk ? '--cmyk' : '--rgb');
      pyArgs.add(settings.isVisibleCorel ? '--visible' : '--hidden');

      final process = await Process.run(
        'python',
        pyArgs,
        workingDirectory: '$projectRoot/python_engine/Core',
      ).timeout(const Duration(seconds: 45));

      if (!mounted) return;

      if (process.exitCode == 0) {
        // Find newest PREVIEW pdf in Outputs
        final outputsDir = Directory('$plakamaticDir/Outputs');
        final pdfs = outputsDir
            .listSync()
            .where((f) => f.path.endsWith('_PREVIEW.pdf'))
            .toList();

        if (pdfs.isNotEmpty) {
          // Sort to get the most recent just in case
          pdfs.sort(
            (a, b) => a.statSync().modified.compareTo(b.statSync().modified),
          );
          setState(() {
            _previewPath = pdfs.last.path;
          });
        } else {
          _errorMessage = "Processed, but PREVIEW PDF was not generated.";
        }
      } else {
        setState(() {
          _errorMessage = "Failed: ${process.stderr}\n${process.stdout}";
        });
      }
    } catch (e) {
      if (mounted) {
        setState(() {
          _errorMessage = e.toString();
        });
      }
    } finally {
      if (mounted) {
        setState(() {
          _isProcessing = false;
        });
      }
    }
  }

  Future<void> _printNow(SettingsViewModel settings) async {
    setState(() {
      _isProcessing = true;
    });
    
    try {
      final String projectRoot = Directory.current.path;
      final pythonScript = '$projectRoot/python_engine/Core/send_to_printer.py';
      
      final cleanupService = Provider.of<CleanupService>(context, listen: false);
      String targetPdf = '$projectRoot/PlakaMatik Files/Outputs/LTO_Batch_${cleanupService.currentSessionId}_PRINT.pdf';

      final process = await Process.run(
        'python',
        [pythonScript, targetPdf, settings.selectedPrinter, 'single'],
        workingDirectory: '$projectRoot/python_engine/Core',
      ).timeout(const Duration(seconds: 30));

      if (!mounted) return;

      if (process.exitCode == 0) {
         ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content: Text('Atomic Process Confirmed: Printed physically!')));
      } else {
         ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text('Failed: ${process.stderr}')));
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(const SnackBar(content: Text('Spooler Error: Python Engine detached.')));
      }
    } finally {
      if (mounted) {
        setState(() {
           _isProcessing = false;
        });
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    final settings = Provider.of<SettingsViewModel>(context);
    return Padding(
      padding: const EdgeInsets.all(25.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            'Single Plate Print',
            style: TextStyle(
              fontSize: 26,
              fontWeight: FontWeight.w900,
              color: Color(0xFF1E3A5F),
            ),
          ),
          const SizedBox(height: 20),

          Expanded(
            child: Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                // LEFT PANE
                Expanded(
                  flex: 1,
                  child: Column(
                    children: [
                      // INPUT CONTAINER
                      Container(
                        padding: const EdgeInsets.all(20),
                        decoration: BoxDecoration(
                          color: Colors.white.withOpacity(0.9),
                          borderRadius: BorderRadius.circular(15),
                        ),
                        child: Column(
                          children: [
                            const Text(
                              "Input Parameters",
                              style: TextStyle(
                                fontSize: 18,
                                fontWeight: FontWeight.bold,
                              ),
                            ),
                            const SizedBox(height: 20),
                            TextField(
                              controller: _middleController,
                              decoration: const InputDecoration(
                                labelText: 'Middle Value (e.g. 20TH CONGRESS)',
                                border: OutlineInputBorder(),
                              ),
                            ),
                            const SizedBox(height: 15),
                            if (_plateType ==
                                'MV') // Hide identifier for MC visually
                              TextField(
                                controller: _identifierController,
                                decoration: const InputDecoration(
                                  labelText: 'Identifier (e.g. BAM)',
                                  border: OutlineInputBorder(),
                                ),
                              ),
                            if (_plateType == 'MV') const SizedBox(height: 15),
                            DropdownButtonFormField<String>(
                              value: _plateType,
                              decoration: const InputDecoration(
                                labelText: 'Plate Type',
                                border: OutlineInputBorder(),
                              ),
                              items: ['MV', 'MC'].map((String type) {
                                return DropdownMenuItem<String>(
                                  value: type,
                                  child: Text(type),
                                );
                              }).toList(),
                              onChanged: (val) =>
                                  setState(() => _plateType = val!),
                            ),
                            const SizedBox(height: 25),
                            SizedBox(
                              width: double.infinity,
                              height: 50,
                              child: ElevatedButton.icon(
                                onPressed: _isProcessing
                                    ? null
                                    : () => _generateSinglePlate(settings),
                                icon: _isProcessing
                                    ? const SizedBox(
                                        height: 20,
                                        width: 20,
                                        child: CircularProgressIndicator(
                                          strokeWidth: 2,
                                        ),
                                      )
                                    : const Icon(Icons.build),
                                label: Text(
                                  _isProcessing
                                      ? 'Generating via CorelDRAW...'
                                      : 'Generate A3 Preview',
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(height: 15),
                      // LOGS CONTAINER
                      const Expanded(
                        child: ConsoleLogWidget(),
                      ),
                    ],
                  ),
                ),
                const SizedBox(width: 25),
                // RIGHT PANE (PDF PREVIEW)
                Expanded(
                  flex: 1,
                  child: Container(
                    decoration: BoxDecoration(
                      color: Colors.white.withOpacity(0.9),
                      borderRadius: BorderRadius.circular(15),
                    ),
                    child: _errorMessage != null
                        ? Padding(
                            padding: const EdgeInsets.all(15),
                            child: SingleChildScrollView(
                              child: Text(
                                _errorMessage!,
                                style: const TextStyle(color: Colors.red),
                              ),
                            ),
                          )
                        : _previewPath == null
                        ? const Center(
                            child: Text("Waiting for operator payload..."),
                          )
                        : Column(
                            crossAxisAlignment: CrossAxisAlignment.stretch,
                            children: [
                              Expanded(
                                child: ClipRRect(
                                  borderRadius: const BorderRadius.vertical(
                                    top: Radius.circular(15),
                                  ),
                                  child: SfPdfViewer.file(
                                    File(_previewPath!),
                                    canShowScrollHead: false,
                                    canShowScrollStatus: false,
                                  ),
                                ),
                              ),
                              Container(
                                color: Colors.white,
                                padding: const EdgeInsets.symmetric(
                                  horizontal: 20,
                                  vertical: 15,
                                ),
                                child: Row(
                                  mainAxisAlignment:
                                      MainAxisAlignment.spaceBetween,
                                  children: [
                                    const Text(
                                      "A3 Preview Linked",
                                      style: TextStyle(
                                        color: Colors.green,
                                        fontWeight: FontWeight.bold,
                                      ),
                                    ),
                                    FilledButton.icon(
                                      style: FilledButton.styleFrom(
                                        backgroundColor: const Color(
                                          0xFF1E3A5F,
                                        ),
                                        padding: const EdgeInsets.symmetric(
                                          horizontal: 30,
                                          vertical: 15,
                                        ),
                                      ),
                                      onPressed: () => _printNow(settings),
                                      icon: const Icon(Icons.print),
                                      label: const Text(
                                        "Confirm & Print",
                                        style: TextStyle(fontSize: 16),
                                      ),
                                    ),
                                  ],
                                ),
                              ),
                            ],
                          ),
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
