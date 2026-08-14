//go:build windows

// Package lnk provides functionality to create and read Windows shortcut (.lnk) files.
// It uses the WScript.Shell COM object from Windows Script Host to interact with shortcut files.
//
// Example usage:
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

const (
	DefaultIconLocation = "%SystemRoot%\\System32\\SHELL32.dll,0"

	WindowStyleNormal    = "1"
	WindowStyleMaximized = "3"
	WindowStyleMinimized = "7"
)

// Validation errors returned by Read or Write.
var (
	ErrEmptyPath        = errors.New("lnk: path cannot be empty")
	ErrInvalidExtension = errors.New("lnk: path must have .lnk extension")
	ErrNoTargetPath     = errors.New("lnk: shortcut TargetPath is required")
)

// Shortcut represents a Windows shortcut (.lnk file) and its properties.
// TargetPath is required by Write; all other fields are optional.
type Shortcut struct {
	// TargetPath is the path to the target file, folder, or URL that the shortcut points to
	TargetPath string

	// Arguments are command-line arguments to pass to the target when executed
	Arguments string

	// Description is a human-readable description of the shortcut
	Description string

	// Hotkey is the keyboard shortcut to activate this shortcut (e.g., "Ctrl+Alt+M").
	// Windows normalizes the key order, so Read may return a different order than Write.
	Hotkey string

	// IconLocation is the icon file and 0-based index, "path,index"
	// (e.g., "shell32.dll,0"). For the target's own icon, use the target
	// path (e.g., "C:\\app.exe,0"). If empty, defaults to DefaultIconLocation.
	IconLocation string

	// WindowStyle is the target window's display mode; use the WindowStyle* constants.
	WindowStyle string

	// WorkingDirectory is the initial working directory when the target is launched
	WorkingDirectory string
}

type propBinding struct {
	name  string
	field *string
}

// propBindings defines the single shared field order for Read and Write.
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

// Read reads a .lnk file's shortcut properties; missing files satisfy fs.ErrNotExist.
func Read(path string) (Shortcut, error) {
	var sc Shortcut

	if err := validatePath(path); err != nil {
		return sc, err
	}

	// WSH's CreateShortcut opens-or-creates silently; it never reports a missing file.
	if _, err := os.Stat(path); err != nil {
		return sc, fmt.Errorf("lnk: %w", err)
	}

	sh, err := newWshShell()
	if err != nil {
		return sc, fmt.Errorf("failed to initialize shell: %w", err)
	}
	defer sh.Close()

	disp, err := sh.createShortcut(path)
	if err != nil {
		return sc, err
	}
	defer disp.Release()

	for _, binding := range sc.propBindings() {
		prop, err := oleutil.GetProperty(disp, binding.name)
		if err != nil {
			return sc, fmt.Errorf("failed to get property %s: %w", binding.name, err)
		}

		// Unset properties (VT_EMPTY/VT_NULL) hit no case and stay empty.
		switch prop.VT {
		case ole.VT_BSTR:
			*binding.field = prop.ToString()
		case ole.VT_I4:
			*binding.field = fmt.Sprintf("%d", prop.Value())
		}

		// Clear in the loop; a deferred Clear would hold every BSTR until
		// Read returns.
		prop.Clear()
	}

	return sc, nil
}

// Write creates a .lnk file; ErrNoTargetPath is returned when TargetPath is blank.
func Write(path string, sc Shortcut) error {
	if err := validatePath(path); err != nil {
		return err
	}
	// A shortcut without a target is invalid on Windows.
	if strings.TrimSpace(sc.TargetPath) == "" {
		return ErrNoTargetPath
	}

	sc.IconLocation = cmp.Or(sc.IconLocation, DefaultIconLocation)
	sc.WindowStyle = cmp.Or(sc.WindowStyle, WindowStyleNormal)

	sh, err := newWshShell()
	if err != nil {
		return fmt.Errorf("failed to initialize shell: %w", err)
	}
	defer sh.Close()

	disp, err := sh.createShortcut(path)
	if err != nil {
		return err
	}
	defer disp.Release()

	for _, binding := range sc.propBindings() {
		res, err := oleutil.PutProperty(disp, binding.name, *binding.field)
		if err != nil {
			return fmt.Errorf("failed to set property %s: %w", binding.name, err)
		}
		res.Clear()
	}

	res, err := oleutil.CallMethod(disp, "Save")
	if err != nil {
		return fmt.Errorf("failed to save shortcut: %w", err)
	}
	res.Clear()

	return nil
}

// validatePath checks the path first; WSH never reports a bad path itself.
func validatePath(path string) error {
	if strings.TrimSpace(path) == "" {
		return ErrEmptyPath
	}
	// filepath.Ext (not a suffix check) rejects "dir.lnk/file" and "x.lnk.bak".
	if !strings.EqualFold(filepath.Ext(path), ".lnk") {
		return ErrInvalidExtension
	}
	return nil
}

// HRESULT values go-ole does not export.
const (
	// sFalse: CoInitializeEx succeeded; thread was already STA, count still incremented.
	sFalse = 0x1
	// rpcEChangedMode: thread already initialized with a different concurrency model (MTA).
	rpcEChangedMode = 0x80010106
)

// wshShell pins the WScript.Shell COM session and its cleanup to one OS
// thread, so CoUninitialize and UnlockOSThread run on the same thread.
type wshShell struct {
	dispatch *ole.IDispatch
	// coInitialized tracks whether CoInitializeEx actually succeeded on this
	// thread. CoUninitialize must only be called when true; RPC_E_CHANGED_MODE
	// means initialization failed and there is no reference count to release.
	coInitialized bool
	// threadLocked tracks whether LockOSThread is in effect, keeping repeated
	// Close calls a true no-op on every Go version.
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

// init initializes COM and the WScript.Shell dispatch.
func (sh *wshShell) init() error {
	coInitialized := true
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY); err != nil {
		var oleErr *ole.OleError
		if !errors.As(err, &oleErr) {
			return fmt.Errorf("failed to initialize COM: %w", err)
		}

		switch oleErr.Code() {
		case sFalse:
			// S_FALSE still increments the reference count, pairing is required.
		case rpcEChangedMode:
			// Initialization failed; must NOT call CoUninitialize later.
			coInitialized = false
		default:
			return fmt.Errorf("failed to initialize COM: %w", err)
		}
	}
	sh.coInitialized = coInitialized

	unknown, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		return fmt.Errorf("failed to create WScript.Shell object: %w", err)
	}
	// QI AddRef'ed the IDispatch, an independent reference; the IUnknown is
	// always released here, before the caller's Close can CoUninitialize.
	defer unknown.Release()

	dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return fmt.Errorf("failed to query interface: %w", err)
	}
	sh.dispatch = dispatch
	return nil
}

// Close must be called once per successful newWshShell; repeated calls are no-ops.
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

// createShortcut returns the WshShortcut object; Write's Save creates the .lnk file.
func (sh *wshShell) createShortcut(path string) (*ole.IDispatch, error) {
	res, err := oleutil.CallMethod(sh.dispatch, "CreateShortcut", path)
	if err != nil {
		return nil, fmt.Errorf("failed to create shortcut object: %w", err)
	}
	// res's dispatch reference is transferred to the caller; never res.Clear().
	return res.ToIDispatch(), nil
}
