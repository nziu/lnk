//go:build windows

package lnk

import (
	"errors"
	"io/fs"
	"path/filepath"
	"testing"
)

func TestRoundTrip(t *testing.T) {
	path := filepath.Join(t.TempDir(), "test_shortcut.lnk")

	want := Shortcut{
		TargetPath:       "C:\\Windows\\System32\\notepad.exe",
		Arguments:        "test.txt",
		Description:      "Test Shortcut",
		Hotkey:           "Alt+Ctrl+T", // Read may reorder modifiers; this is canonical.
		WorkingDirectory: "C:\\Windows\\System32",
		WindowStyle:      WindowStyleNormal,
		IconLocation:     DefaultIconLocation,
	}

	if err := Write(path, want); err != nil {
		t.Fatalf("Write() failed: %v", err)
	}

	got, err := Read(path)
	if err != nil {
		t.Fatalf("Read() failed: %v", err)
	}

	if got != want {
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
	path := filepath.Join(t.TempDir(), "test.lnk")
	for _, target := range []string{"", "   "} {
		if err := Write(path, Shortcut{TargetPath: target}); !errors.Is(err, ErrEmptyTargetPath) {
			t.Errorf("Write(TargetPath=%q) = %v, want ErrEmptyTargetPath", target, err)
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

func TestWriteDefaults(t *testing.T) {
	path := filepath.Join(t.TempDir(), "defaults.lnk")
	target := "C:\\Windows\\System32\\notepad.exe"
	if err := Write(path, Shortcut{TargetPath: target}); err != nil {
		t.Fatalf("Write() failed: %v", err)
	}
	got, err := Read(path)
	if err != nil {
		t.Fatalf("Read() failed: %v", err)
	}
	want := Shortcut{
		TargetPath:   target,
		IconLocation: DefaultIconLocation,
		WindowStyle:  WindowStyleNormal,
	}
	if got != want {
		t.Errorf("got %+v, want %+v", got, want)
	}
}

func TestWriteReplaces(t *testing.T) {
	path := filepath.Join(t.TempDir(), "replace.lnk")
	target := "C:\\Windows\\System32\\notepad.exe"
	first := Shortcut{
		TargetPath:       target,
		Arguments:        "a.txt",
		Description:      "old",
		Hotkey:           "Ctrl+Alt+M",
		WorkingDirectory: "C:\\Windows\\System32",
	}
	if err := Write(path, first); err != nil {
		t.Fatalf("Write(first) failed: %v", err)
	}
	if err := Write(path, Shortcut{TargetPath: target}); err != nil {
		t.Fatalf("Write(replace) failed: %v", err)
	}
	got, err := Read(path)
	if err != nil {
		t.Fatalf("Read() failed: %v", err)
	}
	want := Shortcut{
		TargetPath:   target,
		IconLocation: DefaultIconLocation,
		WindowStyle:  WindowStyleNormal,
	}
	if got != want {
		t.Errorf("got %+v, want %+v", got, want)
	}
}
