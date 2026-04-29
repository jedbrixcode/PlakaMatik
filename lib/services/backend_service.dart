import 'dart:io';
import 'package:flutter/services.dart';
import 'package:path/path.dart' as p;
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

  /// The parent directory (bin/) of the orchestrator.exe.
  /// Use this as the workingDirectory in Process.run to guarantee the shell
  /// context is correct even when the path contains spaces.
  String get binDirPath => File(executablePath).parent.path;

  /// Extracts orchestrator.exe from assets to a writable Documents subfolder.
  /// Safe to call multiple times — only re-extracts if the asset has changed.
  Future<void> initialize() async {
    try {
      final docsDir  = await getApplicationDocumentsDirectory();
      // Use native path separators so Windows never sees a mixed-slash path.
      final binDir   = Directory(p.join(docsDir.path, 'PlakaMatik Files', 'bin'));
      final destFile = File(p.join(binDir.path, 'orchestrator.exe'));

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

      // Normalize to native Windows backslashes
      _executablePath = destFile.path.replaceAll('/', Platform.pathSeparator);
    } catch (e) {
      // Fallback to dev path if asset extraction fails (dev builds without asset)
      final String projectRoot = Directory.current.path;
      _executablePath = p.join(
        projectRoot, 'python_engine', 'Core', 'dist', 'orchestrator.exe'
      );
      // ignore: avoid_print
      print('[BackendService] Asset extraction failed, using dev path. Error: $e');
    }
  }

  /// Self-healing template guard.
  ///
  /// Checks that MV_PLATE.cdr and MC_PLATE.cdr are present in
  ///   Documents\PlakaMatik Files\CorelDRAW Templates\Main Templates\
  /// If either file is missing it is extracted from the bundled Flutter
  /// assets — this matches the exact path the Python config.py expects.
  static const List<String> _requiredTemplates = [
    'MV_PLATE.cdr',
    'MC_PLATE.cdr',
  ];

  Future<void> ensureTemplates() async {
    try {
      final docsDir = await getApplicationDocumentsDirectory();
      final String s = Platform.pathSeparator;
      final templateDir = Directory(
        p.join(docsDir.path, 'PlakaMatik Files', 'CorelDRAW Templates', 'Main Templates'),
      );

      if (!templateDir.existsSync()) {
        templateDir.createSync(recursive: true);
        // ignore: avoid_print
        print('[BackendService] Created template directory: ${templateDir.path}');
      }

      for (final filename in _requiredTemplates) {
        final destFile = File(p.join(templateDir.path, filename));
        if (!destFile.existsSync()) {
          // ignore: avoid_print
          print('[BackendService] Missing template: $filename — restoring from assets...');
          final ByteData data = await rootBundle.load(
            'assets/Templates/Main Templates/$filename',
          );
          final bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
          await destFile.writeAsBytes(bytes, flush: true);
          // ignore: avoid_print
          print('[BackendService] Restored: ${destFile.path}');
        } else {
          // ignore: avoid_print
          print('[BackendService] Template OK: $filename');
        }
      }
    } catch (e) {
      // Non-fatal — log but don't crash startup
      // ignore: avoid_print
      print('[BackendService] ensureTemplates error: $e');
    }
  }

  /// Convenience: runs the orchestrator with --config pointing to PlakaMatik Files/config.json.
  /// Returns the ProcessResult.
  Future<ProcessResult> runOrchestrator({
    required String plakamaticDir,
    Duration timeout = const Duration(seconds: 120),
  }) async {
    final String s = Platform.pathSeparator;

    // Build every path segment with Platform.pathSeparator — no mixed slashes.
    final String normalizedExe    = executablePath.replaceAll('/', s);
    final String normalizedConfig = p.join(plakamaticDir, 'config.json')
        .replaceAll('/', s);
    // workingDirectory = the bin folder — shell is already "inside" it.
    final String workDir = binDirPath.replaceAll('/', s);

    // Clean List<String> args — Flutter + runInShell handles space-quoting.
    return Process.run(
      normalizedExe,
      ['--config', normalizedConfig],
      workingDirectory: workDir,
      runInShell: false,
    ).timeout(timeout);
  }
}
