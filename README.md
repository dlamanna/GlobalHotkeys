# GlobalHotkeys

## Environment variables

- `GLOBALHOTKEYS_PERF` (optional): enables performance timing for audio session refresh.
  - Set to `1` (or any non-empty, non-`0`/`false` value) to enable.

## Setting environment variables (Windows)

### Current PowerShell session
- `$env:GLOBALHOTKEYS_PERF = "1"`

### Persist for future terminals
- `setx GLOBALHOTKEYS_PERF "1"`

After using `setx`, open a new terminal session for the value to take effect.
