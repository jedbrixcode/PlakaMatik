class InputSanitizer {
  static String sanitize(String input) {
    if (input.isEmpty) return input;
    return input.replaceAll(RegExp(r'\s{2,}'), ' ').trim();
  }
}
