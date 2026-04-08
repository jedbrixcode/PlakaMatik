import 'package:flutter/material.dart';
import 'package:provider/provider.dart';
import 'services/cleanup_service.dart';
import 'services/log_watcher_service.dart';
import 'viewmodels/batch_viewmodel.dart';
import 'viewmodels/navigation_viewmodel.dart';
import 'viewmodels/settings_viewmodel.dart';
import 'views/main_layout.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();

  final cleanupService = CleanupService();
  await cleanupService.initialize();

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
  const PlakaMatikApp({Key? key}) : super(key: key);

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'PlakaMatik',
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        fontFamily: 'Segoe UI',
        useMaterial3: true,
        scaffoldBackgroundColor: Colors.transparent, // Background handled by MainLayout stack
      ),
      home: MainLayout(),
    );
  }
}
