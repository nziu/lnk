# lnk

[![Go Reference](https://pkg.go.dev/badge/github.com/nziu/lnk.svg)](https://pkg.go.dev/github.com/nziu/lnk)

Create and read Windows shortcut (.lnk) files in Go.

## Features

- Create Windows shortcuts with full property control
- Read existing shortcut properties
- Pure Go bindings (no cgo) to Windows COM automation

## Installation

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

	// Create a new shortcut
	if err := lnk.Write("notepad.lnk", shortcut); err != nil {
		log.Fatalf("Failed to create shortcut: %v", err)
	}

	// Read and display shortcut properties
	shortcut, err := lnk.Read("notepad.lnk")
	if err != nil {
		log.Fatalf("Failed to read shortcut: %v", err)
	}
	fmt.Printf("%+v\n", shortcut)
}
```

## Shortcut Properties

| Property           | Description                                                                | Default                               |
| ------------------ | -------------------------------------------------------------------------- | ------------------------------------- |
| `TargetPath`       | Path to the target file, folder, or URL                                    | -                                     |
| `Arguments`        | Command-line arguments                                                     | -                                     |
| `Description`      | Shortcut description                                                       | -                                     |
| `WorkingDirectory` | Initial working directory                                                  | -                                     |
| `IconLocation`     | Icon file and index (e.g., `"shell32.dll,0"`)                              | `%SystemRoot%\System32\SHELL32.dll,0` |
| `WindowStyle`      | Window display style: `"1"` (normal), `"3"` (maximized), `"7"` (minimized) | `"1"`                                 |
| `Hotkey`           | Keyboard shortcut (e.g., `"Ctrl+Alt+M"`)                                   | -                                     |
