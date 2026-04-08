import 'package:flutter/material.dart';

class TroubleshootingView extends StatelessWidget {
  const TroubleshootingView({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.all(25.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text('Troubleshooting & Diagnostics', style: TextStyle(fontSize: 26, fontWeight: FontWeight.w900, color: Color(0xFF1E3A5F))),
          const SizedBox(height: 20),
          Expanded(
            child: ListView(
              children: [
                _buildIssue(
                  'Application Freezes During Print', 
                  'This occurs when CorelDRAW encounters a modal dialog or trial screen. Restart the Python Engine and ensure the bypass_trial_screen automation is active. Close any hidden dialog boxes.'
                ),
                _buildIssue(
                  'Com Error: Rpc Server is Unavailable', 
                  'The user or automation engine closed CorelDRAW prematurely. You must wait for the pdf export to finish. Force quit CorelDRAW in Task Manager and restart the job.'
                ),
                _buildIssue(
                  'Offsets Ignored in Output', 
                  'Ensure that the active DX/DY fields in the Template are actually text fields binded to Print Merge and are not converted to curves.'
                ),
                _buildIssue(
                  'Python Engine Timeout', 
                  'The backend execution exceeded 45 seconds. This could happen if the template takes too long to load or export. Try testing with a smaller print chunk or reducing template graphical complexity.'
                ),
              ],
            ),
          )
        ],
      )
    );
  }

  Widget _buildIssue(String title, String solution) {
    return Card(
      elevation: 2,
      margin: const EdgeInsets.only(bottom: 10),
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(10)),
      child: ExpansionTile(
        leading: const Icon(Icons.warning_amber_rounded, color: Colors.orange),
        title: Text(title, style: const TextStyle(fontWeight: FontWeight.bold)),
        children: [
          Container(
             padding: const EdgeInsets.all(16),
             width: double.infinity,
             color: Colors.orange.shade50,
             child: Column(
               crossAxisAlignment: CrossAxisAlignment.start,
               children: [
                 const Text('Resolution:', style: TextStyle(fontWeight: FontWeight.bold, color: Colors.deepOrange)),
                 const SizedBox(height: 5),
                 Text(solution),
               ],
             ),
          )
        ],
      ),
    );
  }
}
