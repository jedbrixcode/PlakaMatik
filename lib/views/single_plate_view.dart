import 'package:flutter/material.dart';

class SinglePlateView extends StatelessWidget {
  const SinglePlateView({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.all(25.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text('Single Plate Print', style: TextStyle(fontSize: 26, fontWeight: FontWeight.w900, color: Color(0xFF1E3A5F))),
          const SizedBox(height: 20),
          Expanded(
             child: Center(
                child: Text('This view acts as a dedicated quick print interface for replacing missing individual plates outside of the normal batch process.\n\nLogic operates symmetrically to Batch View but skips the chunking algorithm.', textAlign: TextAlign.center, style: TextStyle(fontSize: 16, color: Colors.grey.shade600))
             )
          )
        ],
      )
    );
  }
}
