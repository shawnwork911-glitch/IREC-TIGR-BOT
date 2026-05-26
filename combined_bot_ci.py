"""
combined_bot_ci.py
==================
CI-safe wrapper around combined_bot.py for GitHub Actions.

The original combined_bot.py has two things that hang in a non-interactive
environment (no keyboard):

  1.  input("Press Enter to close this window...")  — in main()
  2.  input("Press Enter to close this window...")  — in _fatal()

This wrapper patches both before importing, so the rest of the code is
100% unchanged. You do NOT need to edit combined_bot.py at all.
"""

import builtins
import sys

# ── Patch input() globally so it never blocks ────────────────────────────────
# In CI there is no terminal, so input() would hang forever.
# We replace it with a no-op that just prints the prompt and returns "".
_real_input = builtins.input

def _ci_input(prompt=""):
    if prompt:
        print(f"[CI] Skipping input prompt: {prompt}")
    return ""

builtins.input = _ci_input

# ── Now import and run the real bot ──────────────────────────────────────────
# We import main() from combined_bot and call it directly.
# All the bot logic (IREC, TIGR, combine) runs exactly as normal.

# Add current directory to path so combined_bot.py is importable
sys.path.insert(0, ".")

try:
    from combined_bot import main
    main()
except SystemExit as e:
    # combined_bot calls sys.exit(0) on success, sys.exit(1) on failure.
    # Re-raise so GitHub Actions sees the correct exit code.
    sys.exit(e.code)
except Exception as e:
    print(f"\n[CI ERROR] Bot crashed with unhandled exception: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
