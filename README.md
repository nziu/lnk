# lnk

[![Go Reference](https://pkg.go.dev/badge/github.com/nziu/lnk.svg)](https://pkg.go.dev/github.com/nziu/lnk)

Create and read Windows shortcut (`.lnk`) files in Go. Windows only, Go 1.22+.

```bash
go get github.com/nziu/lnk
```

## Quick Start

```go
package main

import (
	"fmt"
	"log"

	"github.com/nziu/lnk"
)

func main() {
	shortcut := lnk.Shortcut{
		TargetPath:       "C:\\Windows\\System32\\notepad.exe",
		Description:      "Notepad Shortcut",
		Hotkey:           "Alt+Ctrl+T",
		WorkingDirectory: "C:\\Windows\\System32",
	}

	if err := lnk.Write("notepad.lnk", shortcut); err != nil {
		log.Fatalf("Failed to create shortcut: %v", err)
	}

	shortcut, err := lnk.Read("notepad.lnk")
	if err != nil {
		log.Fatalf("Failed to read shortcut: %v", err)
	}
	fmt.Printf("%+v\n", shortcut)
}
```

`Write` replaces the whole shortcut; it is not a partial update. Empty optional fields clear previous values, except `IconLocation` and `WindowStyle`, which fall back to `DefaultIconLocation` and `WindowStyleNormal`.

## Shortcut Properties

| Property           | Description                                                                | Default                               |
| ------------------ | -------------------------------------------------------------------------- | ------------------------------------- |
| `TargetPath`       | Path to the target file, folder, or URL                                    | required                              |
| `Arguments`        | Command-line arguments                                                     | -                                     |
| `Description`      | Shortcut description                                                       | -                                     |
| `WorkingDirectory` | Initial working directory                                                  | -                                     |
| `IconLocation`     | Icon file and index (e.g., `"shell32.dll,0"`)                              | `%SystemRoot%\System32\SHELL32.dll,0` |
| `WindowStyle`      | Window display style: `"1"` (normal), `"3"` (maximized), `"7"` (minimized) | `"1"`                                 |
| `Hotkey`           | Keyboard shortcut (e.g., `"Ctrl+Alt+M"`); Windows may reorder modifiers    | -                                     |
