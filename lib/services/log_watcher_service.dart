import 'dart:async';
import 'dart:io';
import 'package:path_provider/path_provider.dart';

class LogWatcherService {
  Stream<List<String>> get logStream async* {
    List<String> logs = [];
    final appDocDir = await getApplicationDocumentsDirectory();
    final logDir = Directory('${appDocDir.path}/PlakaMatik Files/Logs');
    
    if (!(await logDir.exists())) {
      await logDir.create(recursive: true);
    }
    
    while(true) {
        if(await logDir.exists()) {
            final files = logDir.listSync().where((f) => f.path.endsWith('.txt')).toList();
            if (files.isNotEmpty) {
                files.sort((a, b) => a.statSync().modified.compareTo(b.statSync().modified));
                final logFile = File(files.last.path);
                try {
                    final lines = await logFile.readAsLines();
                    if(lines.length > 100) {
                        logs = lines.sublist(lines.length - 100).toList();
                    } else {
                        logs = lines.toList();
                    }
                    yield logs;
                } catch(e) {
                    // Ignore file lock issues while Python writes
                }
            }
        }
        await Future.delayed(const Duration(milliseconds: 1500));
    }
  }
}
