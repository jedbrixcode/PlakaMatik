import 'dart:async';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import '../services/backend_service.dart';
import '../viewmodels/settings_viewmodel.dart';

/// Universal print countdown dialog.
/// Finds the latest _PRINT.pdf in Outputs and spools it via orchestrator.exe.
class PrintCountdownDialog extends StatefulWidget {
  final SettingsViewModel settings;
  final String label;
  final VoidCallback? onSuccess;

  const PrintCountdownDialog({
    super.key,
    required this.settings,
    this.label = 'UV Plate',
    this.onSuccess,
  });

  static Future<void> show(
    BuildContext context, {
    required SettingsViewModel settings,
    String label = 'UV Plate',
    VoidCallback? onSuccess,
  }) {
    return showDialog(
      context: context,
      barrierDismissible: false,
      builder: (_) => PrintCountdownDialog(
        settings: settings,
        label: label,
        onSuccess: onSuccess,
      ),
    );
  }

  @override
  State<PrintCountdownDialog> createState() => _PrintCountdownDialogState();
}

class _PrintCountdownDialogState extends State<PrintCountdownDialog> {
  int _counter = 3;
  Timer? _timer;
  bool _spooling = false;
  String? _statusMessage;

  @override
  void initState() {
    super.initState();
    _timer = Timer.periodic(const Duration(seconds: 1), (timer) async {
      if (_counter > 1) {
        if (mounted) setState(() => _counter--);
      } else {
        timer.cancel();
        await _dispatch();
      }
    });
  }

  @override
  void dispose() {
    _timer?.cancel();
    super.dispose();
  }

  Future<void> _dispatch() async {
    if (!mounted) return;
    setState(() {
      _spooling = true;
      _statusMessage = 'Sending to spooler...';
    });

    try {
      final docsDir     = await getApplicationDocumentsDirectory();
      final outputsDir  = Directory('${docsDir.path}/PlakaMatik Files/Outputs');
      final configPath  = '${docsDir.path}/PlakaMatik Files/config.json';

      final exePath     = BackendService.instance.executablePath;

      // Find the latest _PRINT.pdf
      File? printFile;
      if (outputsDir.existsSync()) {
        final prints = outputsDir
            .listSync()
            .where((f) => f.path.endsWith('_PRINT.pdf'))
            .toList();
        if (prints.isNotEmpty) {
          prints.sort(
              (a, b) => a.statSync().modified.compareTo(b.statSync().modified));
          printFile = File(prints.last.path);
        }
      }

      if (printFile == null || !printFile.existsSync()) {
        if (mounted) {
          setState(() => _statusMessage = 'No PRINT PDF found in Outputs.');
        }
        await Future.delayed(const Duration(seconds: 2));
        if (mounted) Navigator.of(context).pop();
        return;
      }

      final String s    = Platform.pathSeparator;
      final String exe  = exePath.replaceAll('/', s);
      final String cfg  = configPath.replaceAll('/', s);
      final String pdf  = printFile.path.replaceAll('/', s);
      final String wDir = BackendService.instance.binDirPath.replaceAll('/', s);

      final process = await Process.run(
        exe,
        [
          '--config', cfg,
          '--action', 'spool',
          '--pdf', pdf,
        ],
        workingDirectory: wDir,
        runInShell: true,
      ).timeout(const Duration(seconds: 45));

      if (!mounted) return;

      if (process.exitCode == 0) {
        setState(() => _statusMessage = 'Successfully sent to print spooler!');
        widget.onSuccess?.call();
      } else {
        setState(() =>
            _statusMessage = 'Spool failed: ${process.stderr}\n${process.stdout}');
      }
    } catch (e) {
      if (mounted) setState(() => _statusMessage = 'Error: $e');
    }

    await Future.delayed(const Duration(seconds: 2));
    if (mounted) Navigator.of(context).pop();
  }

  @override
  Widget build(BuildContext context) {
    return AlertDialog(
      backgroundColor: const Color(0xFF1E3A5F),
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
      title: Text(
        _spooling ? 'TRANSMITTING TO HARDWARE' : 'INITIALIZING HARDWARE SPOOLER',
        style: const TextStyle(
          color: Colors.white,
          fontSize: 16,
          fontWeight: FontWeight.bold,
        ),
      ),
      content: Column(
        mainAxisSize: MainAxisSize.min,
        children: [
          Text(
            _spooling
                ? (_statusMessage ?? 'Sending...')
                : 'Committing ${widget.label} to designated UV plate queue...',
            style: const TextStyle(color: Colors.white70, fontSize: 14),
            textAlign: TextAlign.center,
          ),
          const SizedBox(height: 30),
          _spooling
              ? const CircularProgressIndicator(color: Colors.orange)
              : Text(
                  _counter.toString(),
                  style: const TextStyle(
                    fontSize: 70,
                    fontWeight: FontWeight.bold,
                    color: Colors.orange,
                  ),
                ),
        ],
      ),
    );
  }
}
