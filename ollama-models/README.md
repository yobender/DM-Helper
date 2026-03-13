## PF2e Ollama Models

These files let you rebuild the local PF2e models used by DM Helper on another PC.

### What to install first

1. Install Ollama.
2. Make sure Ollama is running.
3. Pull the public base models:
   - `ollama pull gpt-oss:20b`
   - `ollama pull qwen2.5-coder:1.5b-base`
   - `ollama pull qwen2.5:3b`

### Build the custom models

From the app repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup-ollama-models.ps1
```

That builds:

- `gpt-oss-20b-optimized:latest`
- `lorebound-pf2e:latest`
- `lorebound-pf2e-fast:latest`
- `lorebound-pf2e-ultra-fast:latest`
- `lorebound-pf2e-pure:latest`

### Recommended use

- `lorebound-pf2e:latest`
  - Best quality for campaign prep and PDF-grounded answers
  - Heaviest and slowest
- `lorebound-pf2e-fast:latest`
  - Better speed / quality balance
- `lorebound-pf2e-ultra-fast:latest`
  - Very fast fallback on weaker hardware
- `lorebound-pf2e-pure:latest`
  - Lightweight PF2e-only helper

### Hardware note

The 20b models are easy to recreate, but performance depends on the other PC.

- Stronger GPU / more RAM: use `lorebound-pf2e:latest`
- Midrange machine: try `lorebound-pf2e-fast:latest`
- Weak laptop / CPU-heavy box: use `lorebound-pf2e-ultra-fast:latest`
