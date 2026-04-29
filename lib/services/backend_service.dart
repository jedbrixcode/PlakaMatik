import 'dart:io';
import 'package:flutter/services.dart';
import 'package:path/path.dart' as p;
import 'path_service.dart';

/// Manages the lifecycle of the bundled Python orchestrator.
///
/// Responsibilities:
///   • Extract orchestrator.exe from Flutter assets → Documents/bin/ on first run.
///   • Self-heal missing CDR templates from bundled assets.
///   • Expose [runOrchestrator] for callers — all paths are pre-normalized.
class BackendService {
  BackendService._();
  static final BackendService instance = BackendService._();

  String? _executablePath;

  /// Normalized path to orchestrator.exe. Throws if [initialize] was not called.
  String get executablePath => _executablePath ??
      (throw StateError('BackendService.initialize() must be called first.'));

  /// Parent bin/ directory — use as workingDirectory in Process.run.
  String get binDirPath => File(executablePath).parent.path;

  // ── Initialization ──────────────────────────────────────────────────────────

  /// Extracts the orchestrator binary from assets to a writable Documents folder.
  /// Safe to call on every launch — only re-writes when the binary has changed.
  Future<void> initialize() async {
    try {
      final paths   = await PathService.resolve();
      final binDir  = Directory(paths.binDir);
      final dest    = File(paths.orchestratorExe);

      if (!binDir.existsSync()) binDir.createSync(recursive: true);

      final data  = await rootBundle.load('assets/orchestrator.exe');
      final bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);

      if (!dest.existsSync() || dest.lengthSync() != bytes.length) {
        await dest.writeAsBytes(bytes, flush: true);
        // ignore: avoid_print
        print('[BackendService] orchestrator.exe extracted → ${dest.path}');
      } else {
        // ignore: avoid_print
        print('[BackendService] orchestrator.exe up-to-date.');
      }

      _executablePath = PlakaMatikPaths.norm(dest.path);
    } catch (e) {
      // Dev fallback — runs when asset is not bundled (flutter run from source)
      _executablePath = p.join(
        Directory.current.path,
        'python_engine', 'Core', 'dist', 'orchestrator.exe',
      );
      // ignore: avoid_print
      print('[BackendService] Asset extraction failed, using dev path. $e');
    }
  }

  // ── Template Guard ──────────────────────────────────────────────────────────

  static const _requiredTemplates = ['MV_PLATE.cdr', 'MC_PLATE.cdr'];

  /// Verifies CDR templates exist in Documents and restores them from assets
  /// if any are missing. Non-fatal — logs and continues on any error.
  Future<void> ensureTemplates() async {
    try {
      final paths  = await PathService.resolve();
      final tmplDir = Directory(paths.templateDir);
      if (!tmplDir.existsSync()) tmplDir.createSync(recursive: true);

      for (final name in _requiredTemplates) {
        final dest = File(p.join(paths.templateDir, name));
        if (!dest.existsSync()) {
          // ignore: avoid_print
          print('[BackendService] Restoring missing template: $name');
          final data  = await rootBundle.load('assets/Templates/Main Templates/$name');
          final bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
          await dest.writeAsBytes(bytes, flush: true);
          // ignore: avoid_print
          print('[BackendService] Restored → ${dest.path}');
        } else {
          // ignore: avoid_print
          print('[BackendService] Template OK: $name');
        }
      }
    } catch (e) {
      // ignore: avoid_print
      print('[BackendService] ensureTemplates error (non-fatal): $e');
    }
  }

  /// Force-repairs all assets (orchestrator + templates) regardless of current state.
  /// Called by the Troubleshooting Hub's "Asset Repair" button.
  Future<String> repairAllAssets() async {
    final log = StringBuffer();
    try {
      final paths  = await PathService.resolve();

      // 1. Re-extract orchestrator
      final binDir = Directory(paths.binDir);
      final dest   = File(paths.orchestratorExe);
      if (!binDir.existsSync()) binDir.createSync(recursive: true);
      final data  = await rootBundle.load('assets/orchestrator.exe');
      final bytes = data.buffer.asUint8List(data.offsetInBytes, data.lengthInBytes);
      await dest.writeAsBytes(bytes, flush: true);
      log.writeln('✓ orchestrator.exe re-extracted → ${dest.path}');
      _executablePath = PlakaMatikPaths.norm(dest.path);

      // 2. Re-extract templates
      final tmplDir = Directory(paths.templateDir);
      if (!tmplDir.existsSync()) tmplDir.createSync(recursive: true);
      for (final name in _requiredTemplates) {
        final tDest  = File(p.join(paths.templateDir, name));
        final tData  = await rootBundle.load('assets/Templates/Main Templates/$name');
        final tBytes = tData.buffer.asUint8List(tData.offsetInBytes, tData.lengthInBytes);
        await tDest.writeAsBytes(tBytes, flush: true);
        log.writeln('✓ $name restored → ${tDest.path}');
      }
    } catch (e) {
      log.writeln('✗ Repair error: $e');
    }
    return log.toString().trim();
  }

  // ── Process Execution ───────────────────────────────────────────────────────

  /// Runs the orchestrator with [--config <configJson>].
  /// Uses [runInShell: false] to call CreateProcess directly — bypasses
  /// cmd.exe quote-stripping that breaks paths with spaces (e.g. Win10 PRO).
  Future<ProcessResult> runOrchestrator({
    required String plakamaticDir,
    Duration timeout = const Duration(seconds: 120),
  }) async {
    final exe    = PlakaMatikPaths.norm(executablePath);
    final config = PlakaMatikPaths.norm(p.join(plakamaticDir, 'config.json'));
    final wDir   = PlakaMatikPaths.norm(binDirPath);

    return Process.run(
      exe,
      ['--config', config],
      workingDirectory: wDir,
      runInShell: false,
    ).timeout(timeout);
  }
}
