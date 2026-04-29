import 'dart:io';
import 'package:flutter/material.dart';
import '../services/backend_service.dart';

// ── Issue data model ─────────────────────────────────────────────────────────

class _Issue {
  const _Issue({
    required this.category,
    required this.title,
    required this.resolution,
    this.icon = Icons.warning_amber_rounded,
    this.iconColor = const Color(0xFFFFA726),
  });

  final String category;
  final String title;
  final String resolution;
  final IconData icon;
  final Color iconColor;
}

const _issues = <_Issue>[
  // ── Hardware ──────────────────────────────────────────────────────────────
  _Issue(
    category: 'Hardware',
    title: 'Blank Page Printed — Canon iX6700',
    resolution:
        'The Canon iX6700 is a GDI inkjet printer. It cannot interpret raw '
        'PDF bytes. PlakaMatik uses SumatraPDF.exe to render through the Windows '
        'GDI pipeline. Ensure SumatraPDF.exe is present in:\n'
        '  Documents\\PlakaMatik Files\\bin\\\n\n'
        'Download the portable build from sumatrapdfreader.org and place it there, '
        'then use the "Asset Repair" button below to verify the folder.',
    icon: Icons.print_disabled,
    iconColor: Color(0xFFF44336),
  ),
  _Issue(
    category: 'Hardware',
    title: 'Printer Spooler — Access is Denied (Error 5)',
    resolution:
        'Right-click on the Canon iX6700 in Windows Printers & Scanners and '
        'enable "Share this printer." PlakaMatik requires the printer to be '
        'accessible as a shared device for the win32print API to open a handle.\n\n'
        'Alternatively, use the "Flush Spooler" button below to clear any stuck jobs '
        'and restart the Windows Print Spooler service.',
    icon: Icons.lock_outline,
    iconColor: Color(0xFFF44336),
  ),
  _Issue(
    category: 'Hardware',
    title: 'Printer Offline / Not Responding',
    resolution:
        'Check that the Canon iX6700 is:\n'
        '  1. Powered ON\n'
        '  2. Connected via USB or network\n'
        '  3. Set as the default printer in Windows\n\n'
        'In Windows Settings → Printers, right-click and select "See what\'s printing." '
        'Cancel all pending jobs and restart the printer.',
  ),
  _Issue(
    category: 'Hardware',
    title: 'Paper Tray Mismatch',
    resolution:
        'The Canon driver may reject the job if the loaded paper size does not '
        'match the document. PlakaMatik outputs A3 (420 × 297 mm) landscape PDFs. '
        'Ensure A3 paper is loaded and the Canon driver paper size is set to A3 '
        'in Canon iX6700 Printing Preferences.',
  ),

  // ── Engine ────────────────────────────────────────────────────────────────
  _Issue(
    category: 'Engine',
    title: 'Python Engine Timeout (120 s exceeded)',
    resolution:
        'The orchestrator did not finish within 120 seconds. Common causes:\n'
        '  • CorelDRAW opened a modal dialog (trial screen, save prompt)\n'
        '  • A very complex CDR template slows the COM export\n\n'
        'Resolution: Force-quit CorelDRAW in Task Manager, then re-generate. '
        'Increase the trial bypass delay in Settings if the trial dialog is still appearing.',
    icon: Icons.timer_off,
    iconColor: Color(0xFFFFA726),
  ),
  _Issue(
    category: 'Engine',
    title: 'COM Error: RPC Server is Unavailable',
    resolution:
        'CorelDRAW was closed while the Python engine was still connected via COM. '
        'The COM channel was severed mid-operation.\n\n'
        'Resolution:\n'
        '  1. Open Task Manager → End Task on any CorelDRAW.exe processes.\n'
        '  2. Wait 5 seconds for COM handles to release.\n'
        '  3. Re-trigger the Generate button.',
    icon: Icons.cable,
    iconColor: Color(0xFFF44336),
  ),
  _Issue(
    category: 'Engine',
    title: 'CorelDRAW Trial Screen Not Bypassed',
    resolution:
        'The trial bypass (Alt+Z) fires AFTER the COM dispatch call. If CorelDRAW '
        'is slow to load on this machine, the bypass executes before the dialog appears.\n\n'
        'Resolution: In Settings → Trial Bypass Delay, increase the value to 7–10 seconds '
        'to give CorelDRAW more time to show the trial screen before Alt+Z is sent.',
    icon: Icons.hourglass_bottom,
    iconColor: Color(0xFFFFA726),
  ),
  _Issue(
    category: 'Engine',
    title: 'DX/DY Offsets Ignored in Output',
    resolution:
        'Text offset fields are only applied if the CorelDRAW shapes in the "Print Layer" '
        'are genuine Artistic Text objects — NOT curves. If you converted text to curves '
        'in the template, the COM text-replacement API cannot target them.\n\n'
        'Resolution: Open the .cdr template, select the offset text object, and confirm '
        'the status bar shows "Artistic Text." If it shows "Curve," undo the conversion.',
  ),
  _Issue(
    category: 'Engine',
    title: '"Invalid Directory Name" ProcessException',
    resolution:
        'This occurs when a path containing spaces is not normalized before '
        'being passed to Process.run. PlakaMatik uses runInShell: false and '
        'Platform.pathSeparator normalization to prevent this.\n\n'
        'If this appears in the log, use the "Asset Repair" button to re-extract '
        'the orchestrator to ensure it is in the correct Documents path.',
    icon: Icons.folder_off,
    iconColor: Color(0xFFF44336),
  ),

  // ── Automation ────────────────────────────────────────────────────────────
  _Issue(
    category: 'Automation',
    title: 'Template Not Found (MV_PLATE / MC_PLATE)',
    resolution:
        'The Python engine looks for templates in:\n'
        '  Documents\\PlakaMatik Files\\CorelDRAW Templates\\Main Templates\\\n\n'
        'If the files were accidentally deleted, click "Asset Repair" below to '
        'restore both MV_PLATE.cdr and MC_PLATE.cdr from the bundled assets.',
    icon: Icons.description_outlined,
    iconColor: Color(0xFFF44336),
  ),
  _Issue(
    category: 'Automation',
    title: 'Record Parsing Error — "Ñ" Characters',
    resolution:
        'Philippine characters (Ñ, ñ) require UTF-8-BOM encoding in the input '
        'payload. PlakaMatik writes the config with a BOM header (0xEF 0xBB 0xBF) '
        'automatically. If the Python engine reports a UnicodeDecodeError, the '
        'orchestrator.exe may be outdated.\n\n'
        'Resolution: Use "Asset Repair" to re-extract the latest orchestrator.',
    icon: Icons.translate,
    iconColor: Color(0xFFFFA726),
  ),
  _Issue(
    category: 'Automation',
    title: 'PDF Export: 0-Byte Output File',
    resolution:
        'A 0-byte PDF means CorelDRAW\'s PublishToPDF call failed silently. '
        'Common causes:\n'
        '  • The output Outputs/ folder is write-protected\n'
        '  • The destination path contains unsupported characters\n'
        '  • CorelDRAW ran out of virtual memory during export\n\n'
        'Resolution: Restart CorelDRAW and regenerate. Ensure PlakaMatik is '
        'run as Administrator if the folder permission issue persists.',
    icon: Icons.picture_as_pdf,
    iconColor: Color(0xFFF44336),
  ),
];

// ── View ─────────────────────────────────────────────────────────────────────

class TroubleshootingView extends StatefulWidget {
  const TroubleshootingView({super.key});

  @override
  State<TroubleshootingView> createState() => _TroubleshootingViewState();
}

class _TroubleshootingViewState extends State<TroubleshootingView> {
  String _selectedCategory = 'All';
  bool _isSpoolerFlushing = false;
  bool _isRepairing = false;
  String _actionLog = '';

  static const _categories = ['All', 'Hardware', 'Engine', 'Automation'];

  static const _bg = Color(0xFF1A1A2E);
  static const _surface = Color(0xFF16213E);
  static const _card = Color(0xFF1F2B47);
  static const _accent = Color(0xFF4A90E2);
  static const _textPrimary = Color(0xFFE0E6F0);
  static const _textMuted = Color(0xFF8899AA);
  static const _radius = 12.0;
  static const _pad = 15.0;

  // ── Actions ────────────────────────────────────────────────────────────────

  Future<void> _flushSpooler() async {
    setState(() {
      _isSpoolerFlushing = true;
      _actionLog = '';
    });
    final log = StringBuffer('[Spooler] Stopping Windows Print Spooler...\n');
    try {
      final stop = await Process.run('net', [
        'stop',
        'spooler',
      ], runInShell: false).timeout(const Duration(seconds: 15));
      log.writeln(stop.stdout.toString().trim());

      final start = await Process.run('net', [
        'start',
        'spooler',
      ], runInShell: false).timeout(const Duration(seconds: 15));
      log.writeln(start.stdout.toString().trim());
      log.writeln('[Spooler] ✓ Print Spooler restarted successfully.');
    } catch (e) {
      log.writeln('[Spooler] ✗ Error: $e');
      log.writeln('[Spooler] Try running PlakaMatik as Administrator.');
    }
    setState(() {
      _isSpoolerFlushing = false;
      _actionLog = log.toString().trim();
    });
  }

  Future<void> _repairAssets() async {
    setState(() {
      _isRepairing = true;
      _actionLog = '[Repair] Starting asset verification...\n';
    });
    try {
      final result = await BackendService.instance.repairAllAssets();
      setState(() {
        _actionLog = result;
      });
    } catch (e) {
      setState(() {
        _actionLog = '[Repair] ✗ Unexpected error: $e';
      });
    } finally {
      setState(() => _isRepairing = false);
    }
  }

  // ── Build ──────────────────────────────────────────────────────────────────

  @override
  Widget build(BuildContext context) {
    final filtered = _selectedCategory == 'All'
        ? _issues
        : _issues.where((i) => i.category == _selectedCategory).toList();

    return Container(
      color: _bg,
      padding: const EdgeInsets.all(_pad),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          _buildHeader(),
          const SizedBox(height: 16),
          _buildCategoryFilter(),
          const SizedBox(height: 12),
          _buildDebugTools(),
          if (_actionLog.isNotEmpty) ...[
            const SizedBox(height: 10),
            _buildActionLog(),
          ],
          const SizedBox(height: 12),
          Expanded(child: _buildIssueList(filtered)),
        ],
      ),
    );
  }

  Widget _buildHeader() => Column(
    crossAxisAlignment: CrossAxisAlignment.start,
    children: [
      Row(
        children: [
          const Icon(Icons.build_circle_outlined, color: _accent, size: 28),
          const SizedBox(width: 10),
          const Text(
            'Troubleshooting & Diagnostics',
            style: TextStyle(
              fontSize: 22,
              fontWeight: FontWeight.w900,
              color: _textPrimary,
              letterSpacing: 0.5,
            ),
          ),
        ],
      ),
      const SizedBox(height: 4),
      const Text(
        'Expand an issue for step-by-step resolution guidance.',
        style: TextStyle(color: _textMuted, fontSize: 12),
      ),
    ],
  );

  Widget _buildCategoryFilter() => SingleChildScrollView(
    scrollDirection: Axis.horizontal,
    child: Row(
      children: _categories.map((cat) {
        final active = cat == _selectedCategory;
        return Padding(
          padding: const EdgeInsets.only(right: 8),
          child: AnimatedContainer(
            duration: const Duration(milliseconds: 180),
            child: GestureDetector(
              onTap: () => setState(() => _selectedCategory = cat),
              child: Container(
                padding: const EdgeInsets.symmetric(
                  horizontal: 16,
                  vertical: 7,
                ),
                decoration: BoxDecoration(
                  color: active ? _accent : _card,
                  borderRadius: BorderRadius.circular(20),
                  border: Border.all(color: active ? _accent : Colors.white12),
                ),
                child: Text(
                  cat,
                  style: TextStyle(
                    color: active ? Colors.white : _textMuted,
                    fontWeight: active ? FontWeight.bold : FontWeight.normal,
                    fontSize: 13,
                  ),
                ),
              ),
            ),
          ),
        );
      }).toList(),
    ),
  );

  Widget _buildDebugTools() => Container(
    padding: const EdgeInsets.all(_pad),
    decoration: BoxDecoration(
      color: _surface,
      borderRadius: BorderRadius.circular(_radius),
      border: Border.all(color: Colors.white10),
    ),
    child: Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        const Text(
          'Interactive Debug Tools',
          style: TextStyle(
            color: _accent,
            fontWeight: FontWeight.bold,
            fontSize: 13,
            letterSpacing: 0.5,
          ),
        ),
        const SizedBox(height: 12),
        Row(
          children: [
            _buildToolButton(
              label: 'Flush Spooler',
              icon: Icons.refresh,
              color: const Color(0xFFFFA726),
              loading: _isSpoolerFlushing,
              tooltip:
                  'Stops and restarts the Windows Print Spooler service.\n'
                  'Clears stuck print jobs that cause "Access Denied" errors.\n'
                  'Requires PlakaMatik to run as Administrator.',
              onPressed: _flushSpooler,
            ),
            const SizedBox(width: 10),
            _buildToolButton(
              label: 'Asset Repair',
              icon: Icons.download_for_offline_outlined,
              color: const Color(0xFF4CAF50),
              loading: _isRepairing,
              tooltip:
                  'Force re-extracts orchestrator.exe and both CDR templates\n'
                  'from the bundled Flutter assets to Documents\\PlakaMatik Files.\n'
                  'Use this if templates were accidentally deleted or corrupted.',
              onPressed: _repairAssets,
            ),
          ],
        ),
      ],
    ),
  );

  Widget _buildToolButton({
    required String label,
    required IconData icon,
    required Color color,
    required bool loading,
    required String tooltip,
    required VoidCallback onPressed,
  }) => Tooltip(
    message: tooltip,
    preferBelow: false,
    child: FilledButton.icon(
      onPressed: loading ? null : onPressed,
      icon: loading
          ? const SizedBox(
              width: 16,
              height: 16,
              child: CircularProgressIndicator(
                strokeWidth: 2,
                color: Colors.white,
              ),
            )
          : Icon(icon, size: 18),
      label: Text(label, style: const TextStyle(fontSize: 13)),
      style: FilledButton.styleFrom(
        backgroundColor: color,
        foregroundColor: Colors.white,
        disabledBackgroundColor: color.withOpacity(0.4),
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 12),
        shape: RoundedRectangleBorder(
          borderRadius: BorderRadius.circular(_radius),
        ),
      ),
    ),
  );

  Widget _buildActionLog() => Container(
    width: double.infinity,
    padding: const EdgeInsets.all(12),
    decoration: BoxDecoration(
      color: Colors.black87,
      borderRadius: BorderRadius.circular(_radius),
      border: Border.all(color: Colors.white10),
    ),
    child: SelectableText(
      _actionLog,
      style: const TextStyle(
        color: Colors.lightGreen,
        fontFamily: 'monospace',
        fontSize: 11,
        height: 1.5,
      ),
    ),
  );

  Widget _buildIssueList(List<_Issue> items) => ListView.builder(
    itemCount: items.length,
    itemBuilder: (context, index) => _IssueCard(issue: items[index]),
  );
}

// ── Issue Card ────────────────────────────────────────────────────────────────

class _IssueCard extends StatelessWidget {
  const _IssueCard({required this.issue});
  final _Issue issue;

  static const _card = Color(0xFF1F2B47);
  static const _textPrimary = Color(0xFFE0E6F0);
  static const _textMuted = Color(0xFF8899AA);
  static const _radius = 10.0;

  @override
  Widget build(BuildContext context) {
    return Container(
      margin: const EdgeInsets.only(bottom: 8),
      decoration: BoxDecoration(
        color: _card,
        borderRadius: BorderRadius.circular(_radius),
        border: Border.all(color: Colors.white10),
      ),
      child: Theme(
        data: Theme.of(context).copyWith(dividerColor: Colors.transparent),
        child: ExpansionTile(
          leading: Icon(issue.icon, color: issue.iconColor, size: 22),
          title: Text(
            issue.title,
            style: const TextStyle(
              color: _textPrimary,
              fontWeight: FontWeight.w600,
              fontSize: 13,
            ),
          ),
          subtitle: Text(
            issue.category,
            style: const TextStyle(color: _textMuted, fontSize: 11),
          ),
          iconColor: _textMuted,
          collapsedIconColor: _textMuted,
          tilePadding: const EdgeInsets.symmetric(horizontal: 15, vertical: 4),
          childrenPadding: EdgeInsets.zero,
          children: [
            Container(
              width: double.infinity,
              padding: const EdgeInsets.fromLTRB(15, 0, 15, 15),
              decoration: const BoxDecoration(
                color: Color(0xFF0D1421),
                borderRadius: BorderRadius.vertical(
                  bottom: Radius.circular(_radius),
                ),
              ),
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  const Divider(color: Colors.white10, height: 1),
                  const SizedBox(height: 10),
                  const Text(
                    'RESOLUTION',
                    style: TextStyle(
                      color: Color(0xFF4A90E2),
                      fontSize: 10,
                      fontWeight: FontWeight.bold,
                      letterSpacing: 1.2,
                    ),
                  ),
                  const SizedBox(height: 8),
                  SelectableText(
                    issue.resolution,
                    style: const TextStyle(
                      color: Color(0xFFB0BEC5),
                      fontSize: 12,
                      height: 1.6,
                    ),
                  ),
                ],
              ),
            ),
          ],
        ),
      ),
    );
  }
}
