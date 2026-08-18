//go:build windows

// Package lnk creates and reads Windows shortcut (.lnk) files via WScript.Shell.
//
// Write replaces the whole shortcut; it is not a partial update.
// Empty optional fields clear previous values, except IconLocation and
// WindowStyle, which fall back to DefaultIconLocation and WindowStyleNormal.
//
// Example:
//
//	shortcut := lnk.Shortcut{
//		TargetPath:       "C:\\Windows\\System32\\notepad.exe",
//		Description:      "Notepad Shortcut",
//		Hotkey:           "Alt+Ctrl+T",
//		WorkingDirectory: "C:\\Windows\\System32",
//	}
//	err := lnk.Write("notepad.lnk", shortcut)
//
//	shortcut, err := lnk.Read("notepad.lnk")
package lnk

import (
	"cmp"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// DefaultIconLocation is used by Write when IconLocation is empty.
const DefaultIconLocation = "%SystemRoot%\\System32\\SHELL32.dll,0"

const (
	WindowStyleNormal    = "1"
	WindowStyleMaximized = "3"
	WindowStyleMinimized = "7"
)

var (
	ErrEmptyPath        = errors.New("lnk: path cannot be empty")
	ErrInvalidExtension = errors.New("lnk: path must have .lnk extension")
	ErrEmptyTargetPath  = errors.New("lnk: TargetPath cannot be empty")
)

var errCOMApartment = errors.New("lnk: COM already initialized with a different concurrency model")

// Shortcut holds WSH properties of a .lnk file.
// TargetPath is required by Write; all other fields are optional.
type Shortcut struct {
	TargetPath  string
	Arguments   string
	Description string

	// Windows may reorder modifiers, so Read can differ from Write.
	Hotkey string

	// Empty uses DefaultIconLocation, not the target's own icon.
	// Format is "path,index" (e.g. "shell32.dll,0").
	IconLocation string

	// Use WindowStyle* constants; other values are passed through to WSH.
	WindowStyle string

	WorkingDirectory string
}

type propBinding struct {
	name  string
	field *string
}

// propBindings keeps Read and Write on the same WSH properties and order.
func (s *Shortcut) propBindings() []propBinding {
	return []propBinding{
		{"TargetPath", &s.TargetPath},
		{"Arguments", &s.Arguments},
		{"Description", &s.Description},
		{"Hotkey", &s.Hotkey},
		{"IconLocation", &s.IconLocation},
		{"WindowStyle", &s.WindowStyle},
		{"WorkingDirectory", &s.WorkingDirectory},
	}
}

// Read returns the shortcut at path. A missing file satisfies fs.ErrNotExist.
func Read(path string) (Shortcut, error) {
	var sc Shortcut
	if err := validatePath(path); err != nil {
		return sc, err
	}

	// CreateShortcut opens-or-creates and never reports a missing file.
	if _, err := os.Stat(path); err != nil {
		return sc, fmt.Errorf("lnk: %w", err)
	}

	sh, err := newWshShell()
	if err != nil {
		return sc, err
	}
	defer sh.Close()

	disp, err := sh.createShortcut(path)
	if err != nil {
		return sc, err
	}
	defer disp.Release()

	for _, binding := range sc.propBindings() {
		v, err := oleutil.GetProperty(disp, binding.name)
		if err != nil {
			clearVariant(v)
			return sc, fmt.Errorf("lnk: failed to get property %s: %w", binding.name, err)
		}

		// VT_EMPTY/VT_NULL skip the switch and stay the empty string.
		switch v.VT {
		case ole.VT_BSTR:
			*binding.field = v.ToString()
		case ole.VT_I4:
			*binding.field = fmt.Sprintf("%d", v.Value())
		}

		// Clear now; a deferred Clear would keep every BSTR until Read returns.
		v.Clear()
	}

	return sc, nil
}

// Write creates or replaces the .lnk at path. Every field is written.
// Empty optional fields clear prior values, except IconLocation and
// WindowStyle, which use DefaultIconLocation and WindowStyleNormal.
// ErrEmptyTargetPath is returned when TargetPath is blank.
func Write(path string, sc Shortcut) error {
	if err := validatePath(path); err != nil {
		return err
	}
	if strings.TrimSpace(sc.TargetPath) == "" {
		return ErrEmptyTargetPath
	}

	sc.IconLocation = cmp.Or(sc.IconLocation, DefaultIconLocation)
	sc.WindowStyle = cmp.Or(sc.WindowStyle, WindowStyleNormal)

	sh, err := newWshShell()
	if err != nil {
		return err
	}
	defer sh.Close()

	disp, err := sh.createShortcut(path)
	if err != nil {
		return err
	}
	defer disp.Release()

	for _, binding := range sc.propBindings() {
		v, err := oleutil.PutProperty(disp, binding.name, *binding.field)
		if err != nil {
			clearVariant(v)
			return fmt.Errorf("lnk: failed to set property %s: %w", binding.name, err)
		}
		v.Clear()
	}

	v, err := oleutil.CallMethod(disp, "Save")
	if err != nil {
		clearVariant(v)
		return fmt.Errorf("lnk: failed to save shortcut: %w", err)
	}
	v.Clear()

	return nil
}

func validatePath(path string) error {
	if strings.TrimSpace(path) == "" {
		return ErrEmptyPath
	}
	// Ext, not a suffix check, so dir.lnk/file and x.lnk.bak are rejected.
	if !strings.EqualFold(filepath.Ext(path), ".lnk") {
		return ErrInvalidExtension
	}
	return nil
}

func clearVariant(v *ole.VARIANT) {
	if v != nil {
		v.Clear()
	}
}

const (
	// S_FALSE: already STA; the init count still increments.
	sFalse = 0x1
	// RPC_E_CHANGED_MODE: not STA; must not call CoUninitialize.
	rpcEChangedMode = 0x80010106
)

// wshShell pins COM to one OS thread so Close can CoUninitialize safely.
type wshShell struct {
	dispatch *ole.IDispatch
	// Set on S_OK/S_FALSE so Close does not CoUninit a failed init.
	coInitialized bool
	// Held with LockOSThread so repeated Close is a no-op.
	threadLocked bool
}

func newWshShell() (*wshShell, error) {
	sh := &wshShell{threadLocked: true}
	runtime.LockOSThread()
	if err := sh.init(); err != nil {
		sh.Close()
		return nil, err
	}
	return sh, nil
}

func (sh *wshShell) init() error {
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY); err != nil {
		var oleErr *ole.OleError
		if !errors.As(err, &oleErr) {
			return fmt.Errorf("lnk: failed to initialize COM: %w", err)
		}
		switch uint32(oleErr.Code()) {
		case sFalse:
			// Init count still increased; Close must pair it with CoUninitialize.
		case rpcEChangedMode:
			return errCOMApartment
		default:
			return fmt.Errorf("lnk: failed to initialize COM: %w", err)
		}
	}
	sh.coInitialized = true

	unknown, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		return fmt.Errorf("lnk: failed to create WScript.Shell object: %w", err)
	}
	// QI AddRef'd IDispatch separately; release IUnknown before Close.
	defer unknown.Release()

	dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return fmt.Errorf("lnk: failed to query interface: %w", err)
	}
	sh.dispatch = dispatch
	return nil
}

func (sh *wshShell) Close() {
	if sh.dispatch != nil {
		sh.dispatch.Release()
		sh.dispatch = nil
	}
	if sh.coInitialized {
		ole.CoUninitialize()
		sh.coInitialized = false
	}
	if sh.threadLocked {
		runtime.UnlockOSThread()
		sh.threadLocked = false
	}
}

// createShortcut transfers the dispatch to the caller.
// Do not Clear the VARIANT after a successful ToIDispatch (double-Release).
func (sh *wshShell) createShortcut(path string) (*ole.IDispatch, error) {
	v, err := oleutil.CallMethod(sh.dispatch, "CreateShortcut", path)
	if err != nil {
		clearVariant(v)
		return nil, fmt.Errorf("lnk: failed to create shortcut object: %w", err)
	}
	disp := v.ToIDispatch()
	if disp == nil {
		v.Clear()
		return nil, errors.New("lnk: CreateShortcut returned nil dispatch")
	}
	return disp, nil
}
