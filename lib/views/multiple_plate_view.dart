import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../viewmodels/batch_viewmodel.dart';
import '../widgets/batch_input_widget.dart';
import '../widgets/batch_preview_widget.dart';
import '../widgets/console_log_widget.dart';
import '../widgets/batch_queue_list.dart';

class MultiplePlateView extends StatefulWidget {
  const MultiplePlateView({super.key});

  @override
  _MultiplePlateViewState createState() => _MultiplePlateViewState();
}

class _MultiplePlateViewState extends State<MultiplePlateView> {
  bool _hardwareInterlockActive = false;

  void _triggerInterlockGateway() {
    setState(() {
      _hardwareInterlockActive = true;
    });
  }

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<BatchViewModel>(context);

    return LayoutBuilder(
      builder: (context, constraints) {
        return SingleChildScrollView(
          scrollDirection: Axis.horizontal,
          child: SingleChildScrollView(
            scrollDirection: Axis.vertical,
            child: ConstrainedBox(
              constraints: BoxConstraints(
                minWidth: constraints.maxWidth < 1000 ? 1000 : constraints.maxWidth,
                maxWidth: constraints.maxWidth < 1000 ? 1000 : constraints.maxWidth,
                minHeight: constraints.maxHeight < 800 ? 800 : constraints.maxHeight,
                maxHeight: constraints.maxHeight < 800 ? 800 : constraints.maxHeight,
              ),
              child: Padding(
                padding: const EdgeInsets.all(25.0),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    const Text(
                      'Batch Queue Operator Dashboard',
                      style: TextStyle(
                        fontSize: 26,
                        fontWeight: FontWeight.w900,
                        color: Color(0xFF1E3A5F),
                      ),
                    ),
                    if (viewModel.errorMessage != null)
                      Container(
                        margin: const EdgeInsets.only(top: 15),
                        padding: const EdgeInsets.all(15),
                        decoration: BoxDecoration(
                          color: Colors.red.shade50,
                          borderRadius: BorderRadius.circular(8),
                          border: Border.all(color: Colors.redAccent),
                        ),
                        child: Row(
                          children: [
                            const Icon(Icons.error_outline, color: Colors.red),
                            const SizedBox(width: 10),
                            Expanded(
                              child: Text(
                                viewModel.errorMessage!,
                                style: const TextStyle(
                                  color: Colors.red,
                                  fontWeight: FontWeight.w500,
                                ),
                              ),
                            ),
                          ],
                        ),
                      ),
                    const SizedBox(height: 25),
                    
                    // Hardware Interlock Gateway - Overtakes screen when active
                    if (_hardwareInterlockActive)
                      Expanded(
                        child: Container(
                          width: double.infinity,
                          padding: const EdgeInsets.all(40),
                          decoration: BoxDecoration(
                            color: Colors.orange.shade50,
                            borderRadius: BorderRadius.circular(20),
                            border: Border.all(color: Colors.orange, width: 4),
                          ),
                          child: Column(
                            mainAxisAlignment: MainAxisAlignment.center,
                            children: [
                              const Icon(
                                Icons.warning_amber_rounded,
                                size: 80,
                                color: Colors.orange,
                              ),
                              const SizedBox(height: 20),
                              const Text(
                                "Run Complete. Please clear the printing bed.",
                                style: TextStyle(
                                  fontSize: 26,
                                  fontWeight: FontWeight.bold,
                                  color: Colors.orange,
                                ),
                              ),
                              const SizedBox(height: 10),
                              const Text(
                                "The hardware queue requires physical bed clearing confirmation before proceeding with the next batch limits.",
                                textAlign: TextAlign.center,
                                style: TextStyle(fontSize: 16),
                              ),
                              const SizedBox(height: 40),
                              ElevatedButton(
                                style: ElevatedButton.styleFrom(
                                  backgroundColor: Colors.orange,
                                  padding: const EdgeInsets.symmetric(
                                    horizontal: 50,
                                    vertical: 25,
                                  ),
                                ),
                                onPressed: () {
                                  setState(() {
                                    viewModel.requiresNextRunPrompt = false;
                                    _hardwareInterlockActive = false;
                                  });
                                },
                                child: const Text(
                                  "BED IS CLEAR",
                                  style: TextStyle(
                                    color: Colors.white,
                                    fontSize: 20,
                                    fontWeight: FontWeight.bold,
                                  ),
                                ),
                              ),
                            ],
                          ),
                        ),
                      )
                    else
                      Expanded(
                        child: Column(
                          crossAxisAlignment: CrossAxisAlignment.stretch,
                          children: [
                            // ROW 1: [INPUT] | [PREVIEW]
                            Expanded(
                              flex: 6,
                              child: Row(
                                children: [
                                  // 1. INPUT PANE
                                  const Expanded(
                                    flex: 4,
                                    child: BatchInputWidget(),
                                  ),
                                  const SizedBox(width: 15),

                                  // 2. PREVIEW PANE
                                  Expanded(
                                    flex: 5,
                                    child: BatchPreviewWidget(
                                      onInterlockTriggered: _triggerInterlockGateway,
                                    ),
                                  ),
                                ],
                              ),
                            ),
                            const SizedBox(height: 15),

                            // ROW 2: [CONSOLE] | [QUEUE]
                            const Expanded(
                              flex: 4,
                              child: Row(
                                children: [
                                  // 3. CONSOLE PANE
                                  Expanded(
                                    flex: 4,
                                    child: ConsoleLogWidget(),
                                  ),
                                  SizedBox(width: 15),

                                  // 4. QUEUE PANE
                                  Expanded(
                                    flex: 5,
                                    child: BatchQueueList(),
                                  ),
                                ],
                              ),
                            ),
                          ],
                        ),
                      ),
                  ],
                ),
              ),
            ),
          ),
        );
      },
    );
  }
}
