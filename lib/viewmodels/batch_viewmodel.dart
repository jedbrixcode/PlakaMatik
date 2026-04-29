import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:shared_preferences/shared_preferences.dart';
import '../services/backend_service.dart';
import '../utils/input_sanitizer.dart';
import '../services/cleanup_service.dart';
import '../viewmodels/settings_viewmodel.dart';

class PlateData {
  String identifier;
  String designation;
  String plateType;
  double? dxOffset;
  double? dyOffset;
  String? previewPath;

  PlateData(this.identifier, this.designation, this.plateType);

  String toPythonPayload() {
    final cleanId = InputSanitizer.sanitize(identifier);
    final cleanDesig = InputSanitizer.sanitize(designation);
    return '$cleanId    $cleanDesig    $plateType';
  }
}

class BatchViewModel extends ChangeNotifier {
  List<PlateData> printQueue = [];
  bool isProcessing = false;
  String? errorMessage;
  String? latestBatchPreviewPath;
  int currentRunIndex = 0;
  final int platesPerRun = 2;
  bool requiresNextRunPrompt = false;
  String? lastSuccessfullyPrintedId;

  final CleanupService cleanupService;

  BatchViewModel(this.cleanupService) {
    _loadBookmark();
  }

  Future<void> _loadBookmark() async {
    final prefs = await SharedPreferences.getInstance();
    lastSuccessfullyPrintedId = prefs.getString('last_printed_id');
    notifyListeners();
  }

  void addToQueue(String id, String desig, String type) {
    printQueue.add(PlateData(id, desig, type));
    notifyListeners();
  }

  void removeFromQueue(int index) {
    printQueue.removeAt(index);
    notifyListeners();
  }

  void clearQueue() {
    printQueue.clear();
    currentRunIndex = 0;
    errorMessage = null;
    requiresNextRunPrompt = false;
    latestBatchPreviewPath = null;
    notifyListeners();
  }

  Future<void> generateNextPreviewChunk(SettingsViewModel settings) async {
    if (currentRunIndex >= printQueue.length) return;

    isProcessing = true;
    errorMessage = null;
    requiresNextRunPrompt = false;
    notifyListeners();

    int endIndex = (currentRunIndex + platesPerRun > printQueue.length)
        ? printQueue.length
        : currentRunIndex + platesPerRun;

    List<PlateData> currentChunk = printQueue.sublist(
      currentRunIndex,
      endIndex,
    );

    try {
      // Use path_provider to securely reach PlakaMatik Files mapping
      final docsDir = await getApplicationDocumentsDirectory();
      final plakamaticDir = '${docsDir.path}/PlakaMatik Files';
      final file = File('$plakamaticDir/csv/flutter_user_input.txt');

      if (!await file.parent.exists()) {
        await file.parent.create(recursive: true);
      }

      StringBuffer buffer = StringBuffer();
      // Write the explicit header so data_processor.py processes properly
      buffer.writeln('MIDDLE    IDENTIFIER    TYPE');
      for (var plate in currentChunk) {
        buffer.writeln(plate.toPythonPayload());
      }

      // utf-8-sig encoding for Philippine 'Ñ' parsing correctly in COM
      final bytes = [0xEF, 0xBB, 0xBF, ...utf8.encode(buffer.toString())];
      await file.writeAsBytes(bytes);

      final exePath = BackendService.instance.executablePath;
      final configPath = '${docsDir.path}/PlakaMatik Files/config.json';

      final String s    = Platform.pathSeparator;
      final String exe   = exePath.replaceAll('/', s);
      final String cfg   = configPath.replaceAll('/', s);
      final String wDir  = BackendService.instance.binDirPath.replaceAll('/', s);

      final List<String> pyArgs = ['--config', cfg];

      final process = await Process.run(
        exe,
        pyArgs,
        workingDirectory: wDir,
        runInShell: false,
      ).timeout(const Duration(seconds: 120));

      if (process.exitCode == 0) {
        // Scan for the most recently modified *_PREVIEW.pdf in the Outputs folder.
        // We scan rather than using a fixed name because the Python engine's
        // session timestamp may differ by a few seconds from Flutter's.
        final outputsDir = Directory('${docsDir.path}/PlakaMatik Files/Outputs');
        List<FileSystemEntity> previews = [];
        if (outputsDir.existsSync()) {
          previews = outputsDir
              .listSync()
              .where((f) => f is File && f.path.endsWith('_PREVIEW.pdf'))
              .toList()
            ..sort((a, b) =>
                b.statSync().modified.compareTo(a.statSync().modified));
        }

        if (previews.isNotEmpty) {
          final latestPreview = previews.first.path;
          latestBatchPreviewPath = latestPreview;
          // Assign the same combined A3 preview to all plates in this chunk
          for (int i = 0; i < currentChunk.length; i++) {
            final globalIndex = currentRunIndex + i;
            printQueue[globalIndex].previewPath = latestPreview;
          }
        } else {
          errorMessage = 'PDF Export failed. No output detected.';
          latestBatchPreviewPath = null;
        }

        currentRunIndex = endIndex;


        // Save bookmark of the last printed plate from this chunk
        if (currentChunk.isNotEmpty) {
          lastSuccessfullyPrintedId = currentChunk.last.identifier;
          final prefs = await SharedPreferences.getInstance();
          await prefs.setString(
            'last_printed_id',
            currentChunk.last.identifier,
          );
        }

        if (currentRunIndex < printQueue.length) {
          requiresNextRunPrompt = true;
        }
      } else {
        errorMessage =
            'Batch Failed at index ${currentRunIndex + 1}: ${process.stderr.toString().trim()}';
        if (errorMessage!.isEmpty &&
            process.stdout.toString().trim().isNotEmpty) {
          errorMessage = 'Batch Failed: ${process.stdout.toString().trim()}';
        }
      }
    } catch (e) {
      errorMessage =
          'System Communication Timeout. Please verify Python backend.';
    } finally {
      isProcessing = false;
      notifyListeners();
    }
  }

  Future<bool> dispatchToSpooler(SettingsViewModel settings) async {
    isProcessing = true;
    notifyListeners();
    try {
      final exePath = BackendService.instance.executablePath;
      final docsDir = await getApplicationDocumentsDirectory();

      String targetPdf =
          '${docsDir.path}/PlakaMatik Files/Outputs/${cleanupService.currentSessionId}_PRINT.pdf';
      final configPath = '${docsDir.path}/PlakaMatik Files/config.json';

      final String s       = Platform.pathSeparator;
      final String exe      = exePath.replaceAll('/', s);
      final String cfg      = configPath.replaceAll('/', s);
      final String pdf      = targetPdf.replaceAll('/', s);
      final String wDir     = BackendService.instance.binDirPath.replaceAll('/', s);

      final process = await Process.run(
        exe,
        ['--config', cfg, '--action', 'print_corel', '--pdf', pdf],
        workingDirectory: wDir,
        runInShell: false,
      ).timeout(const Duration(seconds: 120));

      if (process.exitCode == 0) {
        requiresNextRunPrompt = true;
        isProcessing = false;
        notifyListeners();
        return true;
      } else {
        errorMessage = 'CorelDRAW Printing Error: ${process.stderr.toString().trim()}';
        isProcessing = false;
        notifyListeners();
        return false;
      }
    } catch (e) {
      errorMessage = 'CorelDRAW printer automation disconnected: $e';
      isProcessing = false;
      notifyListeners();
      return false;
    }
  }
}
