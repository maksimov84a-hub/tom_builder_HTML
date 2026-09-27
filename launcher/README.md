# DeepSeek Harness — launcher

Desktop shortcut for the local DeepSeek Harness web UI.

Official site: <https://www.deepseek.com/harness/en/>

## What it does

```
npx --yes @deepseek-ai/dsh web
```

Starts the local DSH web server and hands the UI to the default browser.
A console window stays open; **closing it stops the server.**

## Files

| File | Purpose |
| --- | --- |
| `deepseek-harness.cmd` | Launcher the shortcut points at. Forwards extra arguments to `dsh web`. |
| `deepseek-harness.ico` | Official whale mark on brand-blue tile. 16–256 px, multi-resolution. |
| `deepseek-harness-mark.ico` | Alternative: bare whale mark, transparent background. |
| `deepseek-harness-icon-512.png` | 512 px icon (docs, store listings). |
| `deepseek-harness-logo.png` | Official whale + `deepseek` wordmark lockup. |
| `deepseek-harness-mark.svg` | Vector source of the mark. |
| `make-shortcut.ps1` | Creates the `.lnk` with icon, name and tooltip. |
| `install-shortcut.cmd` | Double-click helper: installs to Desktop **and** Start Menu. |

## Shortcut properties

- **Target** — `launcher\deepseek-harness.cmd`
- **Start in** — `I:\Progi\tom_builder_HTML` (this folder becomes the agent's workspace)
- **Icon** — `launcher\deepseek-harness.ico,0`
- **Name** — `DeepSeek Harness`

Change **Start in** in the shortcut's Properties to launch against a different project folder.

## Useful flags

```
deepseek-harness.cmd --help             # list dsh web flags
deepseek-harness.cmd --port 8080        # serve on another port
deepseek-harness.cmd --no-open          # do not open a browser
```

## Recreating the shortcut

Double-click `install-shortcut.cmd`, or run:

```powershell
powershell -ExecutionPolicy Bypass -File launcher\make-shortcut.ps1 `
  -Destination "$env:USERPROFILE\Desktop"
```

## Branding source

The whale mark is the official `FISH_LOGO_PATH` from
`@deepseek-ai/dsh-web-frontend`; the wordmark is the official
`BrandWordmark` geometry from the same package. Brand colour `#4D6BFE`.
Nothing is redrawn by hand — the paths are used verbatim.
