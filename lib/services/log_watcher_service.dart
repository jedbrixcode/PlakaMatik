import 'dart:async';
import 'dart:io';
import 'package:path_provider/path_provider.dart';

class LogWatcherService {
  Stream<List<String>> get logStream async* {
    List<String> logs = [];
    final appDocDir = await getApplicationDocumentsDirectory();
    final logFile = File('${appDocDir.path}/Logs/daily_log.txt');
    
    if (!(await logFile.parent.exists())) {
      await logFile.parent.create(recursive: true);
      await logFile.writeAsString("--- LTO PlakaMatik Log Started ---\n");
    }
    
    while(true) {
        if(await logFile.exists()) {
            try {
                final lines = await logFile.readAsLines();
                if(lines.length > 100) {
                     logs = lines.sublist(lines.length - 100).reversed.toList();
                } else {
                     logs = lines.reversed.toList();
                }
                yield logs;
            } catch(e) {
                // Ignore file lock issues while Python writes
            }
        }
        await Future.delayed(const Duration(milliseconds: 1500));
    }
  }
}
