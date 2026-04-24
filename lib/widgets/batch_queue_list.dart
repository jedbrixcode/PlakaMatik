import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../viewmodels/batch_viewmodel.dart';

class BatchQueueList extends StatelessWidget {
  const BatchQueueList({super.key});

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<BatchViewModel>(context);

    return Container(
      padding: const EdgeInsets.all(10),
      decoration: BoxDecoration(
        color: const Color(0xFFF0F4F8),
        borderRadius: BorderRadius.circular(15),
        border: Border.all(color: Colors.grey.shade300),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              const Text(
                'To Print Queue',
                style: TextStyle(
                  fontSize: 16,
                  fontWeight: FontWeight.bold,
                  color: Color(0xFF1E3A5F),
                ),
              ),
              Text(
                '${viewModel.printQueue.length - viewModel.currentRunIndex} remaining',
                style: TextStyle(
                  color: Colors.grey.shade600,
                  fontWeight: FontWeight.bold,
                ),
              ),
            ],
          ),
          const SizedBox(height: 5),
          Expanded(
            child: ListView.builder(
              itemCount: viewModel.printQueue.length,
              itemBuilder: (context, index) {
                final plate = viewModel.printQueue[index];
                final isPrinted = index < viewModel.currentRunIndex;
                final isCurrentChunk =
                    index >= viewModel.currentRunIndex &&
                    index < viewModel.currentRunIndex + viewModel.platesPerRun;
                    
                return Card(
                  elevation: isCurrentChunk ? 4 : 1,
                  color:
                      isPrinted
                          ? Colors.green.shade50
                          : (isCurrentChunk
                              ? Colors.white
                              : Colors.grey.shade50),
                  shape: RoundedRectangleBorder(
                    borderRadius: BorderRadius.circular(10),
                    side: BorderSide(
                      color:
                          isCurrentChunk
                              ? const Color(0xFF4A90E2)
                              : Colors.transparent,
                      width: 2,
                    ),
                  ),
                  margin: const EdgeInsets.only(bottom: 6),
                  child: ListTile(
                    leading: CircleAvatar(
                      backgroundColor:
                          isPrinted
                              ? Colors.green
                              : (isCurrentChunk
                                  ? const Color(0xFF4A90E2)
                                  : Colors.grey),
                      radius: 12,
                      child: Icon(
                        isPrinted ? Icons.check : Icons.format_list_numbered,
                        color: Colors.white,
                        size: 12,
                      ),
                    ),
                    title: Text(
                      '${plate.identifier} | ${plate.plateType}',
                      style: const TextStyle(
                        fontWeight: FontWeight.bold,
                        fontSize: 13,
                      ),
                    ),
                    subtitle: Text(
                      plate.designation,
                      style: const TextStyle(fontSize: 11),
                    ),
                    trailing:
                        isPrinted
                            ? null
                            : IconButton(
                                icon: const Icon(
                                  Icons.delete_outline,
                                  color: Colors.redAccent,
                                  size: 18,
                                ),
                                onPressed: () => viewModel.removeFromQueue(index),
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
