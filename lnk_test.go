package lnk_test

import (
	"os"
	"os/user"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/nziu/lnk"
)

func TestLnk(t *testing.T) {
	user, _ := user.Current()
	tempDir := filepath.Join(user.HomeDir, "/AppData/Local/Temp/")

	tmpfile, err := os.CreateTemp(tempDir, "temp")
	if err != nil {
		t.Fatalf("could not create temp file %v", err)
	}
	defer os.Remove(tmpfile.Name())

	got := lnk.Shortcut{
		TargetPath:       tmpfile.Name(),
		Description:      "tmpfile",
		IconLocation:     "%SystemRoot%\\System32\\SHELL32.dll,0",
		WindowStyle:      "",
		WorkingDirectory: filepath.Dir(tmpfile.Name()),
	}

	tempLnk := tmpfile.Name() + ".lnk"
	if err := lnk.Make(tempLnk, got); err != nil {
		t.Fatalf("unable to make shortcut %v", err)
	}
	defer os.Remove(tempLnk)

	want, err := lnk.Read(tempLnk)
	if err != nil {
		t.Fatalf("unable to read shortcut %v", err)
	}

	if !reflect.DeepEqual(got, want) {
		t.Errorf("got %#v want %#v", got, want)
	}
}
