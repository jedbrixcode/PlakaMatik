import 'package:flutter/material.dart';
import 'dart:io';
import 'package:syncfusion_flutter_pdfviewer/pdfviewer.dart';
import 'package:path_provider/path_provider.dart';
import 'package:provider/provider.dart';
import '../services/backend_service.dart';
import '../viewmodels/settings_viewmodel.dart';
import '../widgets/console_log_widget.dart';
import '../widgets/print_countdown_dialog.dart';

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

  @override
  void dispose() {
    _middleController.dispose();
    _identifierController.dispose();
    super.dispose();
  }

  Future<void> _generateSinglePlate(SettingsViewModel settings) async {
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

      final exePath = BackendService.instance.executablePath;
      final configPath = '$plakamaticDir/config.json';

      final String s   = Platform.pathSeparator;
      final String exe  = exePath.replaceAll('/', s);
      final String cfg  = configPath.replaceAll('/', s);
      final String wDir = BackendService.instance.binDirPath.replaceAll('/', s);

      final process = await Process.run(
        exe,
        ['--config', cfg],
        workingDirectory: wDir,
        runInShell: false,
      ).timeout(const Duration(seconds: 120));

      if (!mounted) return;

      if (process.exitCode == 0) {
        final outputsDir = Directory('$plakamaticDir/Outputs');
        final pdfs = outputsDir
            .listSync()
            .where((f) => f.path.endsWith('_PREVIEW.pdf'))
            .toList();

        if (pdfs.isNotEmpty) {
          pdfs.sort(
            (a, b) => a.statSync().modified.compareTo(b.statSync().modified),
          );
          setState(() {
            _previewPath = pdfs.last.path;
          });
        } else {
          setState(() {
            _errorMessage = "Processed, but PREVIEW PDF was not generated.";
          });
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
                // ── LEFT PANE ────────────────────────────────────────────
                Expanded(
                  flex: 1,
                  child: Column(
                    children: [
                      Container(
                        padding: const EdgeInsets.all(20),
                        decoration: BoxDecoration(
                          color: Colors.white.withOpacity(0.9),
                          borderRadius: BorderRadius.circular(15),
                        ),
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.stretch,
                          children: [
                            const Text(
                              'Input Parameters',
                              style: TextStyle(
                                fontSize: 18,
                                fontWeight: FontWeight.bold,
                              ),
                            ),
                            const SizedBox(height: 16),
                            // Plate type dropdown
                            DropdownButtonFormField<String>(
                              value: _plateType,
                              decoration: const InputDecoration(
                                labelText: 'Plate Type',
                                border: OutlineInputBorder(),
                                isDense: true,
                              ),
                              items: ['MV', 'MC'].map((t) {
                                return DropdownMenuItem(
                                  value: t,
                                  child: Text(t),
                                );
                              }).toList(),
                              onChanged: (val) =>
                                  setState(() => _plateType = val!),
                            ),
                            const SizedBox(height: 12),
                            // Middle value
                            TextField(
                              controller: _middleController,
                              decoration: const InputDecoration(
                                labelText: 'Middle Value',
                                hintText: 'e.g. 20TH CONGRESS',
                                border: OutlineInputBorder(),
                                isDense: true,
                              ),
                            ),
                            // Identifier — MV only
                            if (_plateType == 'MV') ...[
                              const SizedBox(height: 12),
                              TextField(
                                controller: _identifierController,
                                decoration: const InputDecoration(
                                  labelText: 'Identifier',
                                  hintText: 'e.g. BAM',
                                  border: OutlineInputBorder(),
                                  isDense: true,
                                ),
                              ),
                            ],
                            const SizedBox(height: 20),
                            SizedBox(
                              height: 48,
                              child: ElevatedButton.icon(
                                onPressed: _isProcessing
                                    ? null
                                    : () => _generateSinglePlate(settings),
                                icon: _isProcessing
                                    ? const SizedBox(
                                        width: 18,
                                        height: 18,
                                        child: CircularProgressIndicator(
                                          strokeWidth: 2,
                                          color: Colors.white,
                                        ),
                                      )
                                    : const Icon(Icons.preview),
                                label: Text(
                                  _isProcessing
                                      ? 'GENERATING...'
                                      : 'GENERATE UV PLATE PREVIEW',
                                  style: const TextStyle(
                                      fontWeight: FontWeight.bold),
                                ),
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: const Color(0xFF4A90E2),
                                  foregroundColor: Colors.white,
                                  shape: RoundedRectangleBorder(
                                    borderRadius: BorderRadius.circular(10),
                                  ),
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                      const SizedBox(height: 15),
                      const Expanded(child: ConsoleLogWidget()),
                    ],
                  ),
                ),
                const SizedBox(width: 25),
                // ── RIGHT PANE (PDF PREVIEW) ──────────────────────────────
                Expanded(
                  flex: 1,
                  child: Container(
                    decoration: BoxDecoration(
                      color: Colors.white.withOpacity(0.9),
                      borderRadius: BorderRadius.circular(15),
                    ),
                    child: _buildPreviewPane(settings),
                  ),
                ),
              ],
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildPreviewPane(SettingsViewModel settings) {
    if (_errorMessage != null) {
      return Padding(
        padding: const EdgeInsets.all(15),
        child: SingleChildScrollView(
          child: Text(
            _errorMessage!,
            style: const TextStyle(color: Colors.red),
          ),
        ),
      );
    }

    if (_previewPath == null) {
      return const Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Icon(Icons.inbox_outlined, size: 48, color: Colors.grey),
            SizedBox(height: 12),
            Text(
              'No preview generated yet.',
              style: TextStyle(color: Colors.grey, fontSize: 14),
            ),
            SizedBox(height: 4),
            Text(
              'Fill in the form and press\nGENERATE UV PLATE PREVIEW.',
              style: TextStyle(color: Colors.grey, fontSize: 12),
              textAlign: TextAlign.center,
            ),
          ],
        ),
      );
    }

    return Column(
      crossAxisAlignment: CrossAxisAlignment.stretch,
      children: [
        Expanded(
          child: ClipRRect(
            borderRadius:
                const BorderRadius.vertical(top: Radius.circular(15)),
            child: SfPdfViewer.file(
              File(_previewPath!),
              key: ValueKey(
                  '${_previewPath!}_${File(_previewPath!).lastModifiedSync().millisecondsSinceEpoch}'),
              canShowScrollHead: false,
              canShowScrollStatus: false,
            ),
          ),
        ),
        // ── Bottom bar ───────────────────────────────────────────────────
        Container(
          decoration: const BoxDecoration(
            color: Color(0xFFEAF4EA),
            borderRadius:
                BorderRadius.vertical(bottom: Radius.circular(15)),
          ),
          padding:
              const EdgeInsets.symmetric(horizontal: 20, vertical: 14),
          child: Row(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  const Text(
                    'Print Preview for Single Plate',
                    style: TextStyle(
                      color: Colors.green,
                      fontWeight: FontWeight.bold,
                      fontSize: 13,
                    ),
                  ),
                  Text(
                    'Target: ${settings.selectedPrinter}',
                    style: const TextStyle(
                      fontSize: 11,
                      color: Colors.grey,
                    ),
                  ),
                ],
              ),
              FilledButton.icon(
                style: FilledButton.styleFrom(
                  backgroundColor: const Color(0xFF1E3A5F),
                  padding: const EdgeInsets.symmetric(
                      horizontal: 24, vertical: 14),
                ),
                onPressed: () => PrintCountdownDialog.show(
                  context,
                  settings: settings,
                  label: 'Single Plate',
                ),
                icon: const Icon(Icons.print),
                label: const Text(
                  'Confirm & Print',
                  style: TextStyle(fontSize: 15),
                ),
              ),
            ],
          ),
        ),
      ],
    );
  }
}
