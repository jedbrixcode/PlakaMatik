import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import '../viewmodels/batch_viewmodel.dart';
import '../services/log_watcher_service.dart';

class MultiplePlateView extends StatefulWidget {
  @override
  _MultiplePlateViewState createState() => _MultiplePlateViewState();
}

class _MultiplePlateViewState extends State<MultiplePlateView> {
  final TextEditingController _idController = TextEditingController();
  final TextEditingController _desigController = TextEditingController();
  String _selectedType = 'MV';

  @override
  Widget build(BuildContext context) {
    final viewModel = Provider.of<BatchViewModel>(context);

    // Bookmarking banner
    final String lastPrinted = viewModel.lastSuccessfullyPrintedId ?? 'None';

    return Padding(
      padding: const EdgeInsets.all(25.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            mainAxisAlignment: MainAxisAlignment.spaceBetween,
            children: [
              const Text('Multiple Plate Automation Job', style: TextStyle(fontSize: 26, fontWeight: FontWeight.w900, color: Color(0xFF1E3A5F))),
              Container(
                padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 8),
                decoration: BoxDecoration(
                  color: Colors.green.shade50,
                  borderRadius: BorderRadius.circular(20),
                  border: Border.all(color: Colors.green.shade200)
                ),
                child: Row(
                  children: [
                    const Icon(Icons.bookmark, color: Colors.green, size: 20),
                    const SizedBox(width: 8),
                    Text('Last Printed ID: $lastPrinted', style: const TextStyle(color: Colors.green, fontWeight: FontWeight.bold)),
                  ],
                ),
              )
            ],
          ),
          if (viewModel.errorMessage != null)
            Container(
              margin: const EdgeInsets.only(top: 15),
              padding: const EdgeInsets.all(15),
              decoration: BoxDecoration(color: Colors.red.shade50, borderRadius: BorderRadius.circular(8), border: Border.all(color: Colors.redAccent)),
              child: Row(
                children: [
                  const Icon(Icons.error_outline, color: Colors.red),
                  const SizedBox(width: 10),
                  Expanded(child: Text(viewModel.errorMessage!, style: const TextStyle(color: Colors.red, fontWeight: FontWeight.w500))),
                ],
              ),
            ),
          const SizedBox(height: 25),
          Expanded(
            child: Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                // Input Section
                Expanded(
                  flex: 4,
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.stretch,
                    children: [
                      const Text('Plate Data Entry', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 18, color: Color(0xFF3B6B88))),
                      const SizedBox(height: 15),
                      TextField(
                        controller: _idController, 
                        decoration: InputDecoration(
                          labelText: 'Plate Identifier', 
                          filled: true, 
                          fillColor: Colors.grey[100],
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8), borderSide: BorderSide.none)
                        )
                      ),
                      const SizedBox(height: 15),
                      TextField(
                        controller: _desigController, 
                        decoration: InputDecoration(
                          labelText: 'Designation / Region', 
                          filled: true, 
                          fillColor: Colors.grey[100],
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8), borderSide: BorderSide.none)
                        )
                      ),
                      const SizedBox(height: 15),
                      DropdownButtonFormField<String>(
                        value: _selectedType,
                        decoration: InputDecoration(
                          labelText: 'Plate Type', 
                          filled: true, 
                          fillColor: Colors.grey[100],
                          border: OutlineInputBorder(borderRadius: BorderRadius.circular(8), borderSide: BorderSide.none)
                        ),
                        items: ['MV', 'MC'].map((String value) {
                          return DropdownMenuItem<String>(
                            value: value,
                            child: Text(value),
                          );
                        }).toList(),
                        onChanged: (newValue) {
                          if(newValue != null) setState(() => _selectedType = newValue);
                        },
                      ),
                      const SizedBox(height: 20),
                      ElevatedButton.icon(
                        style: ElevatedButton.styleFrom(
                          backgroundColor: const Color(0xFF1E6083),
                          padding: const EdgeInsets.symmetric(vertical: 20),
                          shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10))
                        ),
                        icon: const Icon(Icons.add, color: Colors.white),
                        label: const Text('Add to Queue', style: TextStyle(color: Colors.white, fontSize: 16, fontWeight: FontWeight.bold)),
                        onPressed: () {
                          if(_idController.text.isNotEmpty && _desigController.text.isNotEmpty) {
                             viewModel.addToQueue(_idController.text, _desigController.text, _selectedType);
                             _idController.clear();
                             _desigController.clear();
                          }
                        },
                      ),
                      const Spacer(),
                      // Execution Section
                      if (viewModel.requiresNextRunPrompt)
                        ElevatedButton.icon(
                          style: ElevatedButton.styleFrom(
                            backgroundColor: Colors.orange.shade600,
                            padding: const EdgeInsets.symmetric(vertical: 25),
                            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10))
                          ),
                          icon: const Icon(Icons.cleaning_services, color: Colors.white),
                          label: const Text('PRINTER BED CLEAR: START NEXT RUN', style: TextStyle(color: Colors.white, fontSize: 16, fontWeight: FontWeight.bold, letterSpacing: 1.2)),
                          onPressed: viewModel.executeNextChunk,
                        )
                      else
                        ElevatedButton.icon(
                          style: ElevatedButton.styleFrom(
                            backgroundColor: const Color(0xFF4A90E2),
                            padding: const EdgeInsets.symmetric(vertical: 25),
                            shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10))
                          ),
                          icon: viewModel.isProcessing ? const SizedBox(width:24, height:24, child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2)) : const Icon(Icons.print, color: Colors.white),
                          label: Text(viewModel.isProcessing ? 'Processing Chunk...' : 'EXECUTE PRINT QUEUE', style: const TextStyle(color: Colors.white, fontSize: 16, fontWeight: FontWeight.bold, letterSpacing: 1.2)),
                          onPressed: viewModel.isProcessing || viewModel.printQueue.isEmpty || viewModel.currentRunIndex >= viewModel.printQueue.length ? null : viewModel.executeNextChunk,
                        ),
                    ],
                  ),
                ),
                const SizedBox(width: 30),
                // Queue Section
                Expanded(
                  flex: 5,
                  child: Container(
                    padding: const EdgeInsets.all(20),
                    decoration: BoxDecoration(color: const Color(0xFFF0F4F8), borderRadius: BorderRadius.circular(15), border: Border.all(color: Colors.grey.shade300)),
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Row(
                          mainAxisAlignment: MainAxisAlignment.spaceBetween,
                          children: [
                            const Text('To Print Queue', style: TextStyle(fontSize: 18, fontWeight: FontWeight.bold, color: Color(0xFF1E3A5F))),
                            Text('${viewModel.printQueue.length - viewModel.currentRunIndex} remaining', style: TextStyle(color: Colors.grey.shade600, fontWeight: FontWeight.bold)),
                          ],
                        ),
                        const SizedBox(height: 15),
                        Expanded(
                          child: ListView.builder(
                            itemCount: viewModel.printQueue.length,
                            itemBuilder: (context, index) {
                              final plate = viewModel.printQueue[index];
                              final isPrinted = index < viewModel.currentRunIndex;
                              final isCurrentChunk = index >= viewModel.currentRunIndex && index < viewModel.currentRunIndex + viewModel.platesPerRun;
                              
                              return Card(
                                elevation: isCurrentChunk ? 4 : 1,
                                color: isPrinted ? Colors.green.shade50 : (isCurrentChunk ? Colors.white : Colors.grey.shade50),
                                shape: RoundedRectangleBorder(
                                  borderRadius: BorderRadius.circular(10),
                                  side: BorderSide(color: isCurrentChunk ? const Color(0xFF4A90E2) : Colors.transparent, width: 2)
                                ),
                                margin: const EdgeInsets.only(bottom: 10),
                                child: ExpansionTile(
                                  leading: CircleAvatar(
                                    backgroundColor: isPrinted ? Colors.green : (isCurrentChunk ? const Color(0xFF4A90E2) : Colors.grey),
                                    child: Icon(isPrinted ? Icons.check : Icons.format_list_numbered, color: Colors.white, size: 18),
                                  ),
                                  title: Text('${plate.identifier} | ${plate.plateType}', style: const TextStyle(fontWeight: FontWeight.bold)),
                                  subtitle: Text(plate.designation),
                                  trailing: isPrinted ? null : IconButton(
                                    icon: const Icon(Icons.delete_outline, color: Colors.redAccent),
                                    onPressed: () => viewModel.removeFromQueue(index),
                                  ),
                                  children: [
                                    Container(
                                      padding: const EdgeInsets.all(15),
                                      color: Colors.grey.shade100,
                                      width: double.infinity,
                                      child: Column(
                                        crossAxisAlignment: CrossAxisAlignment.start,
                                        children: [
                                          const Text('Technical Data', style: TextStyle(fontWeight: FontWeight.bold, fontSize: 12, color: Colors.grey)),
                                          const SizedBox(height: 5),
                                          Text('Δx (Offset): ${plate.dxOffset ?? 'Pending'}'),
                                          Text('Δy (Offset): ${plate.dyOffset ?? 'Pending'}'),
                                          if (plate.previewPath != null) Text('Preview: ${plate.previewPath}'),
                                        ],
                                      ),
                                    )
                                  ]
                                ),
                              );
                            },
                          ),
                        ),
                      ],
                    ),
                  ),
                ),
              ],
            ),
          ),
          const SizedBox(height: 20),
          // Audit Log Console
          Container(
             height: 120,
             padding: const EdgeInsets.all(10),
             decoration: BoxDecoration(color: Colors.black87, borderRadius: BorderRadius.circular(10)),
             child: Column(
               crossAxisAlignment: CrossAxisAlignment.stretch,
               children: [
                 const Text('Audit Log Stream', style: TextStyle(color: Colors.greenAccent, fontSize: 12, fontWeight: FontWeight.bold)),
                 const Divider(color: Colors.white24),
                 Expanded(
                   child: StreamBuilder<List<String>>(
                     stream: Provider.of<LogWatcherService>(context, listen: false).logStream,
                     builder: (context, snapshot) {
                       if (!snapshot.hasData || snapshot.data!.isEmpty) {
                         return const Text('Waiting for logs...', style: TextStyle(color: Colors.white54, fontSize: 12, fontFamily: 'Courier'));
                       }
                       final logs = snapshot.data!;
                       return ListView.builder(
                         itemCount: logs.length,
                         itemBuilder: (context, index) {
                           return Text(logs[index], style: const TextStyle(color: Colors.white70, fontSize: 12, fontFamily: 'Courier'));
                         },
                       );
                     },
                   ),
                 )
               ],
             )
          )
        ],
      ),
    );
  }
}
