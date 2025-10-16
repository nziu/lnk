# lnk

[![Go Reference](https://pkg.go.dev/badge/github.com/nziu/lnk.svg)](https://pkg.go.dev/github.com/nziu/lnk)
[![Go Report Card](https://goreportcard.com/badge/github.com/nziu/lnk)](https://goreportcard.com/report/github.com/nziu/lnk)

A simple Go library for creating and reading Windows shortcut (`.lnk`) files.

## Features

- Create Windows shortcuts with full property control
- Read existing shortcut properties
- Pure Go implementation using COM automation
- Windows-only (uses `WScript.Shell` COM object)

## Installation

```bash
go get github.com/nziu/lnk
```

## Usage

### Creating a Shortcut

```go
package main

import (
    "log"
    "github.com/nziu/lnk"
)

func main() {
    shortcut := lnk.Shortcut{
        TargetPath:       "C:\\Windows\\System32\\notepad.exe",
        Arguments:        "my.txt",
        Description:      "My Application",
        Hotkey:           "Ctrl+Alt+M",
        WorkingDirectory: "C:\\Windows\\System32",
        IconLocation:     lnk.DefaultIconLocation,
        WindowStyle:      lnk.DefaultWindowStyle,
    }

    err := lnk.Make("C:\\Users\\Public\\Desktop\\MyApp.lnk", shortcut)
    if err != nil {
        log.Fatal(err)
    }
}
```

### Reading a Shortcut

```go
package main

import (
    "fmt"
    "log"
    "github.com/nziu/lnk"
)

func main() {
    shortcut, err := lnk.Read("C:\\Users\\Public\\Desktop\\MyApp.lnk")
    if err != nil {
        log.Fatal(err)
    }

    fmt.Printf("Target: %s\n", shortcut.TargetPath)
    fmt.Printf("Arguments: %s\n", shortcut.Arguments)
    fmt.Printf("Description: %s\n", shortcut.Description)
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
