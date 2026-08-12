//go:build windows

package lnk

import (
	"errors"
	"io/fs"
	"path/filepath"
	"testing"
)

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

	got, err := Read(shortcutPath)
	if err != nil {
		t.Fatalf("Read() failed: %v", err)
	}

	if want != got {
		t.Errorf("Shortcut mismatch:\nWant: %+v\nGot:  %+v", want, got)
	}
}

func TestReadNonexistentFile(t *testing.T) {
	_, err := Read(filepath.Join(t.TempDir(), "missing.lnk"))
	if !errors.Is(err, fs.ErrNotExist) {
		t.Fatalf("Read() error = %v, want fs.ErrNotExist", err)
	}
}

func TestWriteEmptyTargetPath(t *testing.T) {
	lnkPath := filepath.Join(t.TempDir(), "test.lnk")
	for _, target := range []string{"", "   "} {
		if err := Write(lnkPath, Shortcut{TargetPath: target}); !errors.Is(err, ErrNoTargetPath) {
			t.Errorf("Write(TargetPath=%q) = %v, want ErrNoTargetPath", target, err)
		}
	}
}

func TestValidatePath(t *testing.T) {
	tests := []struct {
		name    string
		path    string
		wantErr error
	}{
		{"valid", "shortcut.lnk", nil},
		{"uppercase extension", "shortcut.LNK", nil},
		{"empty", "", ErrEmptyPath},
		{"whitespace only", "   ", ErrEmptyPath},
		{"no extension", "shortcut", ErrInvalidExtension},
		{"wrong extension", "shortcut.exe", ErrInvalidExtension},
		{"suffix in directory", "dir.lnk/file", ErrInvalidExtension},
		{"double extension", "x.lnk.bak", ErrInvalidExtension},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if err := validatePath(tt.path); !errors.Is(err, tt.wantErr) {
				t.Errorf("validatePath(%q) error = %v, want %v", tt.path, err, tt.wantErr)
			}
		})
	}
}
