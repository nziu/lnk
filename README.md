# lnk

[![Go Reference](https://pkg.go.dev/badge/github.com/nziu/lnk.svg)](https://pkg.go.dev/github.com/nziu/lnk)
[![Go Report Card](https://goreportcard.com/badge/github.com/nziu/lnk)](https://goreportcard.com/report/github.com/nziu/lnk)

A simple Go library for creating and reading Windows shortcut (`.lnk`) files.

## Features

- Create Windows shortcuts with full property control
- Read existing shortcut properties
- Pure Go implementation using COM automation

## Installation

```bash
go get github.com/nziu/lnk
```

## Quick Start

Here's a simple example to create and read a Windows shortcut:

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
		Arguments:        "",
		Description:      "Notepad Shortcut",
		Hotkey:           "Alt+Ctrl+T",
		WorkingDirectory: "C:\\Windows\\System32",
		IconLocation:     ",0",
		WindowStyle:      "1",
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
	fmt.Printf("Target Path:       %s\n", shortcut.TargetPath)
	fmt.Printf("Description:       %s\n", shortcut.Description)
	fmt.Printf("Working Directory: %s\n", shortcut.WorkingDirectory)
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
