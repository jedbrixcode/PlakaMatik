import 'dart:async';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../services/log_watcher_service.dart';

class ConsoleLogWidget extends StatefulWidget {
  const ConsoleLogWidget({super.key});

  @override
  _ConsoleLogWidgetState createState() => _ConsoleLogWidgetState();
}

class _ConsoleLogWidgetState extends State<ConsoleLogWidget> {
  late Stream<List<String>> _logStream;

  @override
  void initState() {
    super.initState();
    // Cache the stream securely to prevent Flutter from spawning a new
    // while(true) loop every time the parent widget rebuilds.
    _logStream = Provider.of<LogWatcherService>(context, listen: false).logStream.asBroadcastStream();
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
            "Python Execution Logs",
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
                if (!snapshot.hasData || snapshot.data!.isEmpty) {
                  return const Text(
                    'Awaiting logs...',
                    style: TextStyle(
                      color: Colors.lightGreen,
                      fontSize: 11,
                      fontFamily: 'monospace',
                    ),
                  );
                }

                final logs = snapshot.data!;
                return SingleChildScrollView(
                  reverse: true, // Auto-scrolls to latest logs
                  child: SingleChildScrollView(
                    scrollDirection: Axis.horizontal,
                    child: Text(
                      logs.join('\n'),
                      style: const TextStyle(
                        color: Colors.lightGreen,
                        fontFamily: 'monospace',
                        fontSize: 11,
                      ),
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
