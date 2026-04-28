import 'dart:io';
import 'package:flutter/services.dart';
import 'package:path_provider/path_provider.dart';

/// BackendService: extracts the bundled orchestrator.exe from Flutter assets
/// into Documents/PlakaMatik Files/bin/ on first run (or whenever it changes).
/// All Process.run calls should use [executablePath] instead of the dev dist path.
class BackendService {
  static BackendService? _instance;
  static BackendService get instance => _instance ??= BackendService._();
  BackendService._();

  String? _executablePath;

  /// The resolved path to orchestrator.exe.
  /// Call [initialize] before accessing this.
  String get executablePath {
    if (_executablePath == null) {
      throw StateError('BackendService not initialized. Call initialize() first.');
    }
    return _executablePath!;
  }

  /// Extracts orchestrator.exe from assets to a writable Documents subfolder.
  /// Safe to call multiple times — only re-extracts if the asset has changed.
  Future<void> initialize() async {
    try {
      final docsDir   = await getApplicationDocumentsDirectory();
      final binDir    = Directory('${docsDir.path}/PlakaMatik Files/bin');
      final destFile  = File('${binDir.path}/orchestrator.exe');

      if (!binDir.existsSync()) {
        binDir.createSync(recursive: true);
      }

      // Load the asset bytes
      final ByteData assetData = await rootBundle.load('assets/orchestrator.exe');
      final List<int> assetBytes = assetData.buffer.asUint8List(
        assetData.offsetInBytes,
        assetData.lengthInBytes,
      );

      // Only overwrite if file is missing or has a different size (quick change check)
      bool needsWrite = !destFile.existsSync() ||
          destFile.lengthSync() != assetBytes.length;

      if (needsWrite) {
        await destFile.writeAsBytes(assetBytes, flush: true);
        // ignore: avoid_print
        print('[BackendService] orchestrator.exe extracted to ${destFile.path}');
      } else {
        // ignore: avoid_print
        print('[BackendService] orchestrator.exe already up-to-date.');
      }

      _executablePath = destFile.path;
    } catch (e) {
      // Fallback to dev path if asset extraction fails (dev builds without asset)
      final String projectRoot = Directory.current.path;
      _executablePath =
          '$projectRoot/python_engine/Core/dist/orchestrator.exe';
      // ignore: avoid_print
      print('[BackendService] Asset extraction failed, using dev path. Error: $e');
    }
  }

  /// Convenience: runs the orchestrator with --config pointing to PlakaMatik Files/config.json.
  /// Returns the ProcessResult.
  Future<ProcessResult> runOrchestrator({
    required String plakamaticDir,
    Duration timeout = const Duration(seconds: 120),
    String workingDirectory = '',
  }) async {
    final configPath = '$plakamaticDir/config.json';
    final cwd = workingDirectory.isEmpty
        ? File(executablePath).parent.path
        : workingDirectory;

    return Process.run(
      executablePath,
      ['--config', configPath],
      workingDirectory: cwd,
    ).timeout(timeout);
  }
}
