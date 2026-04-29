import 'package:flutter/material.dart';

/// Shared design tokens for the PlakaMatik UI.
///
/// All views read from this class to stay visually consistent.
/// Dark-mode variants are in the [dark] namespace.
abstract class AppTheme {
  // ── Brand colours ─────────────────────────────────────────────────────────
  static const Color navyBlue     = Color(0xFF1E3A5F);
  static const Color accent        = Color(0xFF4A90E2);
  static const Color accentLight   = Color(0xFF6AABFF);
  static const Color success       = Color(0xFF4CAF50);
  static const Color warning       = Color(0xFFFFA726);
  static const Color danger        = Color(0xFFF44336);

  // ── Light-mode surfaces ───────────────────────────────────────────────────
  static const Color lightBg       = Colors.transparent;  // glass backdrop
  static const Color lightCard     = Color(0xE6FFFFFF);   // white 90% opacity
  static const Color lightCardAlt  = Color(0xFFF0F4F8);
  static const Color lightText     = Color(0xFF1E3A5F);
  static const Color lightTextSub  = Color(0xFF607088);
  static const Color lightDivider  = Color(0xFFDDE4EE);

  // ── Dark-mode surfaces ────────────────────────────────────────────────────
  static const Color darkBg        = Color(0xFF1A1A2E);
  static const Color darkSurface   = Color(0xFF16213E);
  static const Color darkCard      = Color(0xFF1F2B47);
  static const Color darkCardAlt   = Color(0xFF0D1421);
  static const Color darkText      = Color(0xFFE0E6F0);
  static const Color darkTextSub   = Color(0xFF8899AA);
  static const Color darkDivider   = Color(0xFF2A3A55);

  // ── Shared geometry ───────────────────────────────────────────────────────
  static const double radiusCard   = 15.0;
  static const double radiusBtn    = 10.0;
  static const double paddingPage  = 25.0;
  static const double paddingCard  = 20.0;

  // ── Typography ────────────────────────────────────────────────────────────
  static const TextStyle pageTitle = TextStyle(
    fontSize: 26,
    fontWeight: FontWeight.w900,
    letterSpacing: 0.3,
  );

  static const TextStyle cardTitle = TextStyle(
    fontSize: 17,
    fontWeight: FontWeight.bold,
  );

  static const TextStyle body = TextStyle(
    fontSize: 13,
    height: 1.65,
  );

  static const TextStyle label = TextStyle(
    fontSize: 11,
    fontWeight: FontWeight.w600,
    letterSpacing: 1.0,
  );

  // ── Helpers ───────────────────────────────────────────────────────────────

  /// Returns the correct card background for the current mode.
  static Color cardColor(bool dark) => dark ? darkCard : lightCard;

  /// Returns the primary text color for the current mode.
  static Color textColor(bool dark) => dark ? darkText : lightText;

  /// Returns the muted/subtitle text color for the current mode.
  static Color textSubColor(bool dark) => dark ? darkTextSub : lightTextSub;

  /// Returns the page background color for the current mode.
  static Color bgColor(bool dark) => dark ? darkBg : lightBg;

  /// Returns the divider color for the current mode.
  static Color dividerColor(bool dark) => dark ? darkDivider : lightDivider;
}
