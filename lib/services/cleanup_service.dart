import 'dart:io';
import 'package:path_provider/path_provider.dart';

class CleanupService {
  late String currentSessionId;

  CleanupService() {
    final now = DateTime.now();
    currentSessionId =
        "${now.year}${now.month.toString().padLeft(2, '0')}${now.day.toString().padLeft(2, '0')}_${now.hour.toString().padLeft(2, '0')}${now.minute.toString().padLeft(2, '0')}";
    
    _performCleanup(); // Fire and forget background deletion lock
  }

  Future<void> _performCleanup() async {

    final appDocDir = await getApplicationDocumentsDirectory();
    final tempPreviewsPath = '${appDocDir.path}/PlakaMatik Files/temp_previews';
    
    final tempPreviewsDir = Directory(tempPreviewsPath);
    
    if (await tempPreviewsDir.exists()) {
      final subDirs = tempPreviewsDir.listSync().whereType<Directory>();
      for (var dir in subDirs) {
        final dirName = dir.path.split(Platform.pathSeparator).last;
        if (dirName != currentSessionId) {
          try {
             await dir.delete(recursive: true);
          } catch(e) {
             print("Cleanup folder failed: $e");
          }
        }
      }
    } else {
      await tempPreviewsDir.create(recursive: true);
    }
    
    final newSessionDir = Directory('$tempPreviewsPath/$currentSessionId');
    if (!(await newSessionDir.exists())) {
      await newSessionDir.create(recursive: true);
    }
  }
}
