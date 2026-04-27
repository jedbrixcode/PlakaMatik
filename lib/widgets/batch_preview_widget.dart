import 'dart:async';
import 'dart:io';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'package:syncfusion_flutter_pdfviewer/pdfviewer.dart';
import '../viewmodels/batch_viewmodel.dart';
import '../viewmodels/settings_viewmodel.dart';

class BatchPreviewWidget extends StatelessWidget {
  final VoidCallback onInterlockTriggered;

  const BatchPreviewWidget({super.key, required this.onInterlockTriggered});

  void _triggerPrintCountdown(
    BuildContext context,
    BatchViewModel viewModel,
    SettingsViewModel settings,
  ) {
    int _counter = 3;
    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (context) {
        return StatefulBuilder(
          builder: (context, setDialogState) {
            Timer.periodic(const Duration(seconds: 1), (timer) async {
              if (_counter > 1) {
                setDialogState(() {
                  _counter--;
                });
              } else {
                timer.cancel();
                Navigator.pop(context); // Close countdown

                // Trigger Hardware abstraction
                bool success = await viewModel.dispatchToSpooler(settings);

                if (success) {
                  onInterlockTriggered(); // Raise Gateway in parent view
                }
              }
            });
            return AlertDialog(
              backgroundColor: const Color(0xFF1E3A5F),
              title: const Text(
                "INITIALIZING HARDWARE SPOOLER",
                style: TextStyle(color: Colors.white),
              ),
              content: Column(
                mainAxisSize: MainAxisSize.min,
                children: [
                  const Text(
                    "Committing chunks to designated UV plate queue...",
                    style: TextStyle(color: Colors.white70),
                  ),
                  const SizedBox(height: 30),
                  Text(
                    _counter.toString(),
                    style: const TextStyle(
                      fontSize: 70,
                      fontWeight: FontWeight.bold,
                      color: Colors.orange,
                    ),
                  ),
                ],
              ),
            );
          },
        );
      },
    );
  }

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<BatchViewModel>(context);
    final settings = Provider.of<SettingsViewModel>(context);

    final String docsPath = Platform.environment['USERPROFILE'] ?? '';
    final outputsDir = Directory('$docsPath/Documents/PlakaMatik Files/Outputs');

    File? latestPreview;
    if (outputsDir.existsSync()) {
      final previews = outputsDir
          .listSync()
          .where((f) => f.path.endsWith('_PREVIEW.pdf'))
          .toList();
      if (previews.isNotEmpty) {
        previews.sort((a, b) => a.statSync().modified.compareTo(b.statSync().modified));
        latestPreview = File(previews.last.path);
      }
    }

    final bool hasPreviewLink = latestPreview != null && latestPreview.existsSync();

    return Container(
      decoration: BoxDecoration(
        color: Colors.white.withValues(alpha: 0.9),
        borderRadius: BorderRadius.circular(15),
      ),
      child: !hasPreviewLink
          ? const Center(
              child: Text(
                "No Preview Generated",
                style: TextStyle(
                  color: Colors.grey,
                  fontWeight: FontWeight.bold,
                  fontSize: 18,
                ),
              ),
            )
          : Column(
              crossAxisAlignment: CrossAxisAlignment.stretch,
              children: [
                Expanded(
                  child: ClipRRect(
                    borderRadius: const BorderRadius.vertical(
                      top: Radius.circular(15),
                    ),
                    child: SfPdfViewer.file(
                      latestPreview!,
                      key: ValueKey(latestPreview.lastModifiedSync().millisecondsSinceEpoch),
                      canShowScrollHead: false,
                      canShowScrollStatus: false,
                    ),
                  ),
                ),
                Container(
                  color: Colors.green.shade50,
                  padding: const EdgeInsets.symmetric(
                    horizontal: 20,
                    vertical: 15,
                  ),
                  child: Row(
                    mainAxisAlignment: MainAxisAlignment.spaceBetween,
                    children: [
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          const Text(
                            "Visual Map Verified",
                            style: TextStyle(
                              color: Colors.green,
                              fontWeight: FontWeight.bold,
                            ),
                          ),
                          Text(
                            "Target: ${settings.selectedPrinter}",
                            style: const TextStyle(
                              fontSize: 10,
                              color: Colors.grey,
                            ),
                          ),
                        ],
                      ),
                      FilledButton.icon(
                        style: FilledButton.styleFrom(
                          backgroundColor: const Color(0xFF1E3A5F),
                          padding: const EdgeInsets.symmetric(
                            horizontal: 30,
                            vertical: 15,
                          ),
                        ),
                        onPressed: () => _triggerPrintCountdown(
                          context,
                          viewModel,
                          settings,
                        ),
                        icon: const Icon(Icons.print),
                        label: const Text(
                          "Confirm & Print",
                          style: TextStyle(fontSize: 16),
                        ),
                      ),
                    ],
                  ),
                ),
              ],
            ),
    );
  }
}
