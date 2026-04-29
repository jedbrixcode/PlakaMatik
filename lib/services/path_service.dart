import 'dart:io';
import 'package:path/path.dart' as p;
import 'package:path_provider/path_provider.dart';

/// Centralized resolver for all PlakaMatik runtime paths.
///
/// All paths are built with [p.join] (native separators) so they work
/// correctly on Windows even when the username contains spaces.
///
/// Usage:
///   final paths = await PathService.resolve();
///   print(paths.configJson);
class PlakaMatikPaths {
  const PlakaMatikPaths._({
    required this.root,
    required this.binDir,
    required this.orchestratorExe,
    required this.configJson,
    required this.outputsDir,
    required this.logsDir,
    required this.templateDir,
    required this.tempPreviewsDir,
    required this.inputTxt,
  });

  /// Documents\PlakaMatik Files
  final String root;

  /// Documents\PlakaMatik Files\bin
  final String binDir;

  /// Documents\PlakaMatik Files\bin\orchestrator.exe
  final String orchestratorExe;

  /// Documents\PlakaMatik Files\config.json
  final String configJson;

  /// Documents\PlakaMatik Files\Outputs
  final String outputsDir;

  /// Documents\PlakaMatik Files\Logs
  final String logsDir;

  /// Documents\PlakaMatik Files\CorelDRAW Templates\Main Templates
  final String templateDir;

  /// Documents\PlakaMatik Files\temp_previews
  final String tempPreviewsDir;

  /// Documents\PlakaMatik Files\input.txt
  final String inputTxt;

  /// Convenience: normalize a path to native Windows backslashes.
  static String norm(String path) =>
      path.replaceAll('/', Platform.pathSeparator);
}

class PathService {
  PathService._();

  /// Resolve all runtime paths from the current Documents directory.
  /// Always call this before accessing paths — do NOT cache across restarts.
  static Future<PlakaMatikPaths> resolve() async {
    final docs = await getApplicationDocumentsDirectory();
    final root = p.join(docs.path, 'PlakaMatik Files');

    return PlakaMatikPaths._(
      root:           root,
      binDir:         p.join(root, 'bin'),
      orchestratorExe: p.join(root, 'bin', 'orchestrator.exe'),
      configJson:     p.join(root, 'config.json'),
      outputsDir:     p.join(root, 'Outputs'),
      logsDir:        p.join(root, 'Logs'),
      templateDir:    p.join(root, 'CorelDRAW Templates', 'Main Templates'),
      tempPreviewsDir: p.join(root, 'temp_previews'),
      inputTxt:       p.join(root, 'input.txt'),
    );
  }
}
