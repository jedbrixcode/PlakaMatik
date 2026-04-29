import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'dart:io';
import '../viewmodels/navigation_viewmodel.dart';
import '../viewmodels/settings_viewmodel.dart';
import 'multiple_plate_view.dart';
import 'single_plate_view.dart';
import 'information_view.dart';
import 'settings_view.dart';
import 'troubleshooting_view.dart';
import '../services/backend_service.dart';

class MainLayout extends StatelessWidget {
  const MainLayout({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Stack(
        children: [
          // Background Layer
          Positioned.fill(
            child: Consumer<SettingsViewModel>(
              builder: (context, settings, child) {
                return Image.asset(
                  settings.isDarkMode ? 'assets/BACKGROUND_DARKMODE.png' : 'assets/BACKGROUND.png',
                  fit: BoxFit.cover,
                  errorBuilder: (context, error, stackTrace) =>
                      Container(color: Colors.blueGrey[50]),
                );
              },
            ),
          ),

          // Main Content Layer
          Positioned(
            left: 270, // Sidebar width + offset
            top: 20,
            bottom: 20,
            right: 20,
            child: ClipRRect(
              borderRadius: BorderRadius.circular(20),
              child: Container(
                color: Colors.white.withOpacity(
                  0.7,
                ), // Glassmorphism-like solid backdrop
                child: Consumer<NavigationViewModel>(
                  builder: (context, navModel, child) {
                    switch (navModel.currentIndex) {
                      case 0:
                        return SinglePlateView();
                      case 1:
                        return const MultiplePlateView();
                      case 2:
                        return const InformationView();
                      case 3:
                        return const SettingsView();
                      case 4:
                        return const TroubleshootingView();
                      default:
                        return MultiplePlateView();
                    }
                  },
                ),
              ),
            ),
          ),

          // Translucent Debug Button
          Positioned(
            top: 40,
            right: 40,
            child: Opacity(
              opacity: 0.6,
              child: FilledButton.icon(
                onPressed: () async {
                  final exePath = BackendService.instance.executablePath;
                  print("Triggering Manual Orchestrator Debug Button...");
                  Process.run(
                    exePath,
                    [],
                    workingDirectory: File(exePath).parent.path,
                    runInShell: true,
                  ).then((result) {
                    print("Orchestrator Terminated. Code: ${result.exitCode}");
                    print("Logs Context: ${result.stdout}");
                    if (result.stderr.toString().isNotEmpty) {
                      print("Errors: ${result.stderr}");
                    }
                  });
                },
                icon: const Icon(Icons.bug_report_outlined),
                label: const Text('DEBUG: Open Python Orchestrator'),
                style: FilledButton.styleFrom(
                  backgroundColor: Colors.redAccent.shade700,
                  foregroundColor: Colors.white,
                  padding: const EdgeInsets.symmetric(
                    horizontal: 20,
                    vertical: 15,
                  ),
                ),
              ),
            ),
          ),

          // Sidebar Layer
          Positioned(
            left: 0,
            top: 0,
            bottom: 0,
            width: 250,
            child: Container(
              decoration: const BoxDecoration(
                color: Color(0xFF1E3A5F),
                boxShadow: [
                  BoxShadow(
                    color: Colors.black38,
                    offset: Offset(5, 0),
                    blurRadius: 15,
                  ),
                ],
              ),
              child: Column(
                children: [
                  Padding(
                    padding: const EdgeInsets.symmetric(vertical: 40.0),
                    child: Image.asset(
                      'assets/LTO_LOGO.png',
                      height: 100,
                      errorBuilder: (context, error, stackTrace) => const Image(
                        image: AssetImage('assets/LTO_LOGO.png'),
                        height: 100,
                      ),
                    ),
                  ),
                  const Text(
                    "PLAKAMATIK",
                    style: TextStyle(
                      color: Colors.white,
                      fontSize: 24,
                      fontWeight: FontWeight.bold,
                      letterSpacing: 2,
                    ),
                  ),
                  const SizedBox(height: 10),
                  const Divider(
                    color: Colors.white24,
                    thickness: 1,
                    endIndent: 20,
                    indent: 20,
                  ),
                  const _NavButton(
                    index: 0,
                    title: "Single Print",
                    icon: Icons.print,
                  ),
                  const _NavButton(
                    index: 1,
                    title: "Batch Print",
                    icon: Icons.view_list,
                  ),
                  const _NavButton(
                    index: 2,
                    title: "Information",
                    icon: Icons.info,
                  ),
                  const _NavButton(
                    index: 3,
                    title: "Settings",
                    icon: Icons.settings,
                  ),
                  const _NavButton(
                    index: 4,
                    title: "Troubleshoot",
                    icon: Icons.build,
                  ),
                ],
              ),
            ),
          ),
        ],
      ),
    );
  }
}

class _NavButton extends StatelessWidget {
  final int index;
  final String title;
  final IconData icon;

  const _NavButton({
    required this.index,
    required this.title,
    required this.icon,
  });

  @override
  Widget build(BuildContext context) {
    final navModel = Provider.of<NavigationViewModel>(context);
    final isSelected = navModel.currentIndex == index;

    return Padding(
      padding: const EdgeInsets.symmetric(horizontal: 10.0, vertical: 5.0),
      child: Material(
        color: Colors.transparent,
        child: InkWell(
          borderRadius: BorderRadius.circular(10),
          onTap: () => navModel.setIndex(index),
          hoverColor: Colors.white10,
          child: Container(
            padding: const EdgeInsets.symmetric(
              vertical: 15.0,
              horizontal: 20.0,
            ),
            decoration: BoxDecoration(
              color: isSelected ? const Color(0xFF4A90E2) : Colors.transparent,
              borderRadius: BorderRadius.circular(10),
            ),
            child: Row(
              children: [
                Icon(icon, color: Colors.white, size: 24),
                const SizedBox(width: 15),
                Text(
                  title,
                  style: const TextStyle(color: Colors.white, fontSize: 16),
                ),
              ],
            ),
          ),
        ),
      ),
    );
  }
}
