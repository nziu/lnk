package lnk

import (
	"os"
	"path/filepath"
	"reflect"
	"testing"
)

// TestRoundTrip tests that a shortcut survives a write/read round-trip
func TestRoundTrip(t *testing.T) {
	tempDir := t.TempDir()
	shortcutPath := filepath.Join(tempDir, "test_shortcut.lnk")

	want := Shortcut{
		TargetPath:       "C:\\Windows\\System32\\notepad.exe",
		Arguments:        "test.txt",
		Description:      "Test Shortcut",
		Hotkey:           "Alt+Ctrl+T", // Windows standardizes hotkey order
		WorkingDirectory: "C:\\Windows\\System32",
		WindowStyle:      WindowStyleNormal,
		IconLocation:     DefaultIconLocation,
	}

	if err := Write(shortcutPath, want); err != nil {
		t.Fatalf("Write() failed: %v", err)
	}

	if _, err := os.Stat(shortcutPath); os.IsNotExist(err) {
		t.Fatal("shortcut file was not created")
	}

	got, err := Read(shortcutPath)
	if err != nil {
		t.Fatalf("Read() failed: %v", err)
	}

	if !reflect.DeepEqual(want, got) {
		t.Errorf("Shortcut mismatch:\nWant: %+v\nGot:  %+v", want, got)
	}
}
