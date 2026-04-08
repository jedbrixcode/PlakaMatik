import 'package:flutter/material.dart';

class InformationView extends StatelessWidget {
  const InformationView({super.key});

  @override
  Widget build(BuildContext context) {
    return Padding(
      padding: const EdgeInsets.all(25.0),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          const Text(
            'Information & Guides',
            style: TextStyle(
              fontSize: 26,
              fontWeight: FontWeight.w900,
              color: Color(0xFF1E3A5F),
            ),
          ),
          const SizedBox(height: 20),
          Expanded(
            child: DefaultTabController(
              length: 3,
              child: Column(
                children: [
                  Container(
                    decoration: BoxDecoration(
                      color: Colors.grey.shade200,
                      borderRadius: BorderRadius.circular(10),
                    ),
                    child: TabBar(
                      padding: EdgeInsets.symmetric(horizontal: 1.0),
                      labelColor: Colors.white,
                      unselectedLabelColor: const Color(0xFF1E3A5F),
                      indicatorColor: Colors.transparent,
                      indicator: BoxDecoration(
                        color: const Color(0xFF4A90E2),
                        borderRadius: BorderRadius.circular(10),
                      ),

                      tabs: [
                        Tab(text: "Program Description"),
                        Tab(text: "Quick Start Guide"),
                        Tab(text: "Essential Setup"),
                      ],
                    ),
                  ),
                  const SizedBox(height: 20),
                  Expanded(
                    child: TabBarView(
                      children: [
                        _buildScrollableCard(
                          'Program Description',
                          'PlakaMatik is a high-performance automation bridge for the LTO Plate Making Plant. It shifts production from manual silk-screening to a Direct-to-UV Printing pipeline using CorelDraw automation running silently via Python.\n\nThe system enforces chunking logic to protect the physical printer bed limits.',
                        ),
                        _buildScrollableCard(
                          'Quick Start Guide',
                          '1. Go to "Batch Print".\n2. Enter Identifiers and Designations.\n3. Verify the queue list on the right.\n4. Click "Execute Print Queue".\n5. Wait for the Python engine to execute the batch dynamically.\n6. Clear the real printer bed when prompted.',
                        ),
                        _buildScrollableCard(
                          'Essential Setup',
                          '1. Ensure Python Engine is installed and linked in the Settings View.\n2. Ensure directory paths for Csv and temp_previews are valid and writeable.\n3. Make sure the LTO Template Corel Draw files are placed in "CorelDRAW Templates" directory within the project root.',
                        ),
                      ],
                    ),
                  ),
                ],
              ),
            ),
          ),
        ],
      ),
    );
  }

  Widget _buildScrollableCard(String title, String content) {
    return Card(
      elevation: 2,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(15)),
      child: Padding(
        padding: const EdgeInsets.all(20.0),
        child: SingleChildScrollView(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(
                title,
                style: const TextStyle(
                  fontSize: 20,
                  fontWeight: FontWeight.bold,
                  color: Color(0xFF3B6B88),
                ),
              ),
              const Divider(thickness: 1, height: 20),
              Text(content, style: const TextStyle(fontSize: 16, height: 1.5)),
            ],
          ),
        ),
      ),
    );
  }
}
