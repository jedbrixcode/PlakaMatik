import 'dart:io';
import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../viewmodels/batch_viewmodel.dart';
import '../viewmodels/settings_viewmodel.dart';
import '../widgets/print_countdown_dialog.dart';
import 'package:syncfusion_flutter_pdfviewer/pdfviewer.dart';

class BatchPreviewWidget extends StatelessWidget {
  final VoidCallback onInterlockTriggered;

  const BatchPreviewWidget({super.key, required this.onInterlockTriggered});

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<BatchViewModel>(context);
    final settings  = Provider.of<SettingsViewModel>(context);

    String? previewPath;
    // Check if the current chunk has a preview path assigned from a successful python execution
    if (viewModel.printQueue.isNotEmpty &&
        viewModel.currentRunIndex < viewModel.printQueue.length) {
      previewPath = viewModel.printQueue[viewModel.currentRunIndex].previewPath;
    }

    File? latestPreview = previewPath != null ? File(previewPath) : null;

    final bool hasPreview = latestPreview != null && latestPreview.existsSync();
    final bool moreRemaining = viewModel.currentRunIndex < viewModel.printQueue.length;

    return Container(
      decoration: BoxDecoration(
        color: Colors.white.withValues(alpha: 0.9),
        borderRadius: BorderRadius.circular(15),
      ),
      child: !hasPreview
          ? const Center(
              child: Text(
                'No Preview Generated',
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
                      key: ValueKey(
                          latestPreview.lastModifiedSync().millisecondsSinceEpoch),
                      canShowScrollHead: false,
                      canShowScrollStatus: false,
                    ),
                  ),
                ),
                Container(
                  decoration: BoxDecoration(
                    color: Colors.green.shade50,
                    borderRadius: const BorderRadius.vertical(
                        bottom: Radius.circular(15)),
                  ),
                  padding: const EdgeInsets.symmetric(
                      horizontal: 20, vertical: 15),
                  child: Row(
                    mainAxisAlignment: MainAxisAlignment.spaceBetween,
                    children: [
                      Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          const Text(
                            'Print Preview for Batch Plate',
                            style: TextStyle(
                              color: Colors.green,
                              fontWeight: FontWeight.bold,
                            ),
                          ),
                          Text(
                            'Target: ${settings.selectedPrinter}',
                            style: const TextStyle(
                                fontSize: 10, color: Colors.grey),
                          ),
                          if (moreRemaining)
                            Text(
                              '${viewModel.printQueue.length - viewModel.currentRunIndex} plate(s) still in queue',
                              style: const TextStyle(
                                  fontSize: 10, color: Colors.orange),
                            ),
                        ],
                      ),
                      FilledButton.icon(
                        style: FilledButton.styleFrom(
                          backgroundColor: const Color(0xFF1E3A5F),
                          padding: const EdgeInsets.symmetric(
                              horizontal: 30, vertical: 15),
                        ),
                        onPressed: () => PrintCountdownDialog.show(
                          context,
                          settings: settings,
                          label: 'Batch Plate',
                          onSuccess: () {
                            // After print is dispatched:
                            // Check if more plates remain in the queue
                            if (moreRemaining) {
                              _showBedClearanceDialog(
                                  context, viewModel, settings);
                            } else {
                              onInterlockTriggered();
                            }
                          },
                        ),
                        icon: const Icon(Icons.print),
                        label: const Text(
                          'Confirm & Print',
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

  /// Step 6.1 — Show "Is the printer bed clear?" dialog.
  void _showBedClearanceDialog(
    BuildContext context,
    BatchViewModel viewModel,
    SettingsViewModel settings,
  ) {
    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (ctx) => AlertDialog(
        backgroundColor: const Color(0xFF1B2A3B),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
        title: const Row(
          children: [
            Icon(Icons.cleaning_services, color: Colors.orange),
            SizedBox(width: 10),
            Text(
              'PRINTER BED CHECK',
              style: TextStyle(
                color: Colors.white,
                fontSize: 16,
                fontWeight: FontWeight.bold,
              ),
            ),
          ],
        ),
        content: const Text(
          'Has the printer finished and is the printer bed clear and ready for the next batch?',
          style: TextStyle(color: Colors.white70, fontSize: 14),
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(ctx).pop(),
            child: const Text('Not yet',
                style: TextStyle(color: Colors.white54)),
          ),
          FilledButton(
            style: FilledButton.styleFrom(
                backgroundColor: Colors.orange),
            onPressed: () {
              Navigator.of(ctx).pop();
              _showContinueQueueDialog(context, viewModel, settings);
            },
            child: const Text('Bed is Clear'),
          ),
        ],
      ),
    );
  }

  /// Step 6.1.1 — Ask if the user wants to process the next batch.
  void _showContinueQueueDialog(
    BuildContext context,
    BatchViewModel viewModel,
    SettingsViewModel settings,
  ) {
    final remaining =
        viewModel.printQueue.length - viewModel.currentRunIndex;

    showDialog(
      context: context,
      barrierDismissible: false,
      builder: (ctx) => AlertDialog(
        backgroundColor: const Color(0xFF1B2A3B),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
        title: const Row(
          children: [
            Icon(Icons.playlist_play, color: Colors.green),
            SizedBox(width: 10),
            Text(
              'CONTINUE QUEUE',
              style: TextStyle(
                color: Colors.white,
                fontSize: 16,
                fontWeight: FontWeight.bold,
              ),
            ),
          ],
        ),
        content: Text(
          'There are $remaining plate(s) remaining in the queue.\nProceed with the next batch?',
          style: const TextStyle(color: Colors.white70, fontSize: 14),
        ),
        actions: [
          TextButton(
            onPressed: () {
              Navigator.of(ctx).pop();
              onInterlockTriggered(); // Exit to queue view
            },
            child: const Text('Stop Queue',
                style: TextStyle(color: Colors.white54)),
          ),
          FilledButton(
            style: FilledButton.styleFrom(
                backgroundColor: const Color(0xFF1E3A5F)),
            onPressed: () {
              Navigator.of(ctx).pop();
              // Process next chunk — this triggers the rebuild
              // which will show the new preview once done
              viewModel.generateNextPreviewChunk(settings);
            },
            child: const Text('Process Next Batch'),
          ),
        ],
      ),
    );
  }
}
