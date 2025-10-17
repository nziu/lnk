package lnk

import (
	"os"
	"path/filepath"
	"reflect"
	"runtime"
	"testing"
)

// TestMakeAndRead tests creating and reading a shortcut
func TestMakeAndRead(t *testing.T) {
	// Skips the test if not running on Windows
	if runtime.GOOS != "windows" {
		t.Skip("skipping test on non-Windows platform")
	}

	tempDir := t.TempDir()
	shortcutPath := filepath.Join(tempDir, "test_shortcut.lnk")

	// Define expected shortcut properties
	want := Shortcut{
		TargetPath:       "C:\\Windows\\System32\\notepad.exe",
		Arguments:        "test.txt",
		Description:      "Test Shortcut",
		Hotkey:           "Alt+Ctrl+T", // Windows standardizes hotkey order
		WorkingDirectory: "C:\\Windows\\System32",
		WindowStyle:      DefaultWindowStyle,
		IconLocation:     DefaultIconLocation,
	}

	// Create shortcut
	if err := Make(shortcutPath, want); err != nil {
		t.Fatalf("Make() failed: %v", err)
	}

	// Verify file exists
	if _, err := os.Stat(shortcutPath); os.IsNotExist(err) {
		t.Fatal("shortcut file was not created")
	}

	// Read shortcut
	got, err := Read(shortcutPath)
	if err != nil {
		t.Fatalf("Read() failed: %v", err)
	}

	// Compare all properties using reflect.DeepEqual
	if !reflect.DeepEqual(want, got) {
		t.Errorf("Shortcut mismatch:\nWant: %+v\nGot:  %+v", want, got)
	}
}
