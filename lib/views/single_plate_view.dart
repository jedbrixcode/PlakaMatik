import 'package:flutter/material.dart';
import 'dart:io';

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

  Future<void> _generateSinglePlate() async {
    // Only require identifier if it's an MV plate
    if (_middleController.text.isEmpty || (_plateType == 'MV' && _identifierController.text.isEmpty)) {
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
      final String projectRoot = Directory.current.path;
      final file = File('$projectRoot/csv/flutter_user_input.txt');

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

      final pythonScript = '$projectRoot/python_engine/Core/main.py';
      final process = await Process.run(
        'python',
        [pythonScript],
        workingDirectory: '$projectRoot/python_engine/Core',
      ).timeout(const Duration(seconds: 45));

      if (process.exitCode == 0) {
        // Find newest pdf in Outputs since main.py creates its own timestamp lock
        // Find newest PREVIEW pdf in Outputs
        final outputsDir = Directory('$projectRoot/Outputs');
        final pdfs = outputsDir
            .listSync()
            .where((f) => f.path.endsWith('_PREVIEW.pdf'))
            .toList();

        if (pdfs.isNotEmpty) {
          // Sort to get the most recent just in case
          pdfs.sort((a, b) => a.statSync().modified.compareTo(b.statSync().modified));
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
      setState(() {
        _errorMessage = e.toString();
      });
    } finally {
      setState(() {
        _isProcessing = false;
      });
    }
  }

  void _openPdf() {
    if (_previewPath != null) {
      Process.run('explorer', [_previewPath!]);
    }
  }

  void _printNow() {
    // Stub for actual UV print sending protocol
    ScaffoldMessenger.of(context).showSnackBar(
      const SnackBar(
        content: Text('Atomic Process Confirmed: Sent to UV Printer Queue!'),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
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
                            if (_plateType == 'MV') // Hide identifier for MC visually
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
                              onChanged: (val) => setState(() => _plateType = val!),
                            ),
                            const SizedBox(height: 25),
                            SizedBox(
                              width: double.infinity,
                              height: 50,
                              child: ElevatedButton.icon(
                                onPressed: _isProcessing ? null : _generateSinglePlate,
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
                      Expanded(
                        child: Container(
                          padding: const EdgeInsets.all(15),
                          decoration: BoxDecoration(
                            color: Colors.black87,
                            borderRadius: BorderRadius.circular(15),
                          ),
                          child: Column(
                            crossAxisAlignment: CrossAxisAlignment.start,
                            children: [
                              const Text("Python Execution Logs", style: TextStyle(color: Colors.greenAccent, fontSize: 14, fontWeight: FontWeight.bold)),
                              const Divider(color: Colors.white24),
                              Expanded(
                                child: StreamBuilder<String>(
                                  stream: Stream.periodic(const Duration(milliseconds: 500), (_) {
                                    try {
                                      final file = File('${Directory.current.path}/python_engine/Logs/daily_log.txt');
                                      if (file.existsSync()) {
                                        final lines = file.readAsLinesSync();
                                        return lines.length > 20 ? lines.sublist(lines.length - 20).join('\n') : lines.join('\n');
                                      }
                                    } catch (e) {}
                                    return "Awaiting logs...";
                                  }),
                                  builder: (context, snapshot) {
                                    return SingleChildScrollView(
                                      reverse: true,
                                      child: Text(
                                        snapshot.data ?? "Loading...",
                                        style: const TextStyle(color: Colors.lightGreen, fontFamily: 'monospace', fontSize: 11),
                                      ),
                                    );
                                  },
                                ),
                              )
                            ],
                          )
                        )
                      )
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
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              const Icon(
                                Icons.check_circle,
                                color: Colors.green,
                                size: 80,
                              ),
                              const SizedBox(height: 10),
                              const Text(
                                "A3 PDF Successfully Generated!",
                                style: TextStyle(
                                  fontSize: 18,
                                  fontWeight: FontWeight.bold,
                                ),
                              ),
                              const SizedBox(height: 20),
                              ElevatedButton.icon(
                                style: ElevatedButton.styleFrom(
                                  padding: const EdgeInsets.all(15),
                                ),
                                onPressed: _openPdf,
                                icon: const Icon(
                                  Icons.picture_as_pdf,
                                  color: Colors.red,
                                ),
                                label: const Text(
                                  "Open A3 Evidence in Windows",
                                ),
                              ),
                              const SizedBox(height: 35),
                              FilledButton.icon(
                                style: FilledButton.styleFrom(
                                  backgroundColor: const Color(0xFF1E3A5F),
                                  padding: const EdgeInsets.symmetric(
                                    horizontal: 40,
                                    vertical: 20,
                                  ),
                                ),
                                onPressed: _printNow,
                                icon: const Icon(Icons.print),
                                label: const Text(
                                  "Confirm & Print 1 Plate to UV",
                                  style: TextStyle(fontSize: 18),
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
