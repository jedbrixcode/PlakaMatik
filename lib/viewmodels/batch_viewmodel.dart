import 'dart:convert';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:shared_preferences/shared_preferences.dart';
import '../utils/input_sanitizer.dart';
import '../services/cleanup_service.dart';

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

  Future<void> executeNextChunk() async {
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
      // Use the strict project root instead of system documents logic
      final String projectRoot = Directory.current.path;
      final file = File('$projectRoot/csv/flutter_user_input.txt');

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

      // We now call the python script natively leveraging the installed compiler
      final pythonScript = '$projectRoot/python_engine/Core/main.py';

      final process = await Process.run('python', [
        pythonScript,
      ], workingDirectory: '$projectRoot/python_engine/Core').timeout(const Duration(seconds: 45));

      if (process.exitCode == 0) {
        try {
          final response = jsonDecode(process.stdout.toString());
          for (int i = 0; i < currentChunk.length; i++) {
            final globalIndex = currentRunIndex + i;
            if (response['data'] != null && response['data'].length > i) {
              printQueue[globalIndex].dxOffset = response['data'][i]['dx'];
              printQueue[globalIndex].dyOffset = response['data'][i]['dy'];
              printQueue[globalIndex].previewPath =
                  response['data'][i]['preview_path'];
            }
          }
        } catch (e) {
          print("JSON Parsing or format unexpected: $e");
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
}
