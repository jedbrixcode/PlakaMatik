import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../services/log_watcher_service.dart';

/// Terminal-style log console that tails the Python engine's daily log.
///
/// Key features:
///   • Auto-scrolls to the latest entry (reverse scroll).
///   • SelectableText — operator can copy long Windows paths directly.
///   • softWrap: true — no horizontal clip on paths like C:\Program Files (x86)\...
class ConsoleLogWidget extends StatefulWidget {
  const ConsoleLogWidget({super.key});

  @override
  State<ConsoleLogWidget> createState() => _ConsoleLogWidgetState();
}

class _ConsoleLogWidgetState extends State<ConsoleLogWidget> {
  late final Stream<List<String>> _logStream;

  @override
  void initState() {
    super.initState();
    // Cache the broadcast stream to avoid re-subscribing on every rebuild.
    _logStream = Provider.of<LogWatcherService>(context, listen: false)
        .logStream
        .asBroadcastStream();
  }

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: const EdgeInsets.all(15),
      decoration: BoxDecoration(
        color: Colors.black87,
        borderRadius: BorderRadius.circular(15),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            'Python Execution Logs',
            style: TextStyle(
              color: Colors.greenAccent,
              fontSize: 14,
              fontWeight: FontWeight.bold,
            ),
          ),
          const Divider(color: Colors.white24),
          Expanded(
            child: StreamBuilder<List<String>>(
              stream: _logStream,
              builder: (context, snapshot) {
                final logs = snapshot.data ?? [];

                if (logs.isEmpty) {
                  return const Text(
                    'Awaiting logs...',
                    style: TextStyle(
                      color: Colors.white38,
                      fontSize: 11,
                      fontFamily: 'monospace',
                    ),
                  );
                }

                return SingleChildScrollView(
                  reverse: true, // pins latest log at the bottom
                  child: SelectableText(
                    logs.join('\n'),
                    style: const TextStyle(
                      color: Colors.lightGreen,
                      fontFamily: 'monospace',
                      fontSize: 11,
                      height: 1.5,
                    ),
                  ),
                );
              },
            ),
          ),
        ],
      ),
    );
  }
}
