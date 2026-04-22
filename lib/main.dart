import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'package:window_manager/window_manager.dart';
import 'services/cleanup_service.dart';
import 'services/log_watcher_service.dart';
import 'viewmodels/batch_viewmodel.dart';
import 'viewmodels/navigation_viewmodel.dart';
import 'viewmodels/settings_viewmodel.dart';
import 'views/main_layout.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  await windowManager.ensureInitialized();
  final cleanupService = CleanupService();


  WindowOptions windowOptions = const WindowOptions(
    size: Size(1280, 720),
    minimumSize: Size(1200, 600),
    center: true,
    title: 'Plakamatik - Automatic UV Printed Protocol Plates',
  );

  windowManager.waitUntilReadyToShow(windowOptions, () async {
    await windowManager.show();
    await windowManager.focus();
  });

  runApp(
    MultiProvider(
      providers: [
        Provider<CleanupService>.value(value: cleanupService),
        Provider<LogWatcherService>(create: (_) => LogWatcherService()),
        ChangeNotifierProvider(create: (_) => NavigationViewModel()),
        ChangeNotifierProvider(create: (_) => SettingsViewModel()),
        ChangeNotifierProvider(create: (_) => BatchViewModel(cleanupService)),
      ],
      child: const PlakaMatikApp(),
    ),
  );
}

class PlakaMatikApp extends StatelessWidget {
  const PlakaMatikApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'PlakaMatik',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        fontFamily: 'Segoe UI',
        useMaterial3: true,
        scaffoldBackgroundColor:
            Colors.transparent, // Background handled by MainLayout stack
      ),
      home: MainLayout(),
    );
  }
}
