//go:build windows

// Package lnk provides functionality to create and read Windows shortcut (.lnk) files.
// It uses Windows Script Shell COM object to interact with shortcut files.
//
// Example usage:
//
//	// Create a new shortcut
//	shortcut := lnk.Shortcut{
//		TargetPath:       "C:\\Program Files\\MyApp\\app.exe",
//		Description:      "My Application",
//		WorkingDirectory: "C:\\Program Files\\MyApp",
//	}
//	err := lnk.Write("C:\\Users\\Desktop\\MyApp.lnk", shortcut)
//
//	// Read an existing shortcut
//	shortcut, err := lnk.Read("C:\\Users\\Desktop\\MyApp.lnk")
package lnk

import (
	"errors"
	"fmt"
	"runtime"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Constants for default values
const (
	DefaultIconLocation = "%SystemRoot%\\System32\\SHELL32.dll,0"

	WindowStyleNormal    = "1"
	WindowStyleMaximized = "3"
	WindowStyleMinimized = "7"
)

// Shortcut represents a Windows shortcut (.lnk file) and its properties.
// All fields are optional when creating a shortcut, but TargetPath is typically required.
type Shortcut struct {
	// TargetPath is the path to the target file, folder, or URL that the shortcut points to
	TargetPath string

	// Arguments are command-line arguments to pass to the target when executed
	Arguments string

	// Description is a human-readable description of the shortcut
	Description string

	// Hotkey is the keyboard shortcut to activate this shortcut (e.g., "Ctrl+Alt+M")
	Hotkey string

	// IconLocation specifies the icon file and index (e.g., "shell32.dll,0")
	// Format: "path,index" where index is the icon number (0-based).
	// To use the target file's own icon, set path to the target
	// (e.g., "C:\\app.exe,0").
	// If empty, defaults to DefaultIconLocation
	IconLocation string

	// WindowStyle controls how the target window is displayed:
	// "1" (default) - normal window
	// "3" - maximized window
	// "7" - minimized window
	// Use the provided constants: WindowStyleNormal, WindowStyleMaximized,
	// WindowStyleMinimized
	WindowStyle string

	// WorkingDirectory is the initial working directory when the target is launched
	WorkingDirectory string
}

// propBinding binds a COM property name to its corresponding Shortcut field
type propBinding struct {
	name  string
	field *string
}

// propBindings returns an ordered slice binding each COM property name to its
// corresponding field pointer. This method provides a centralized way to
// iterate over all shortcut properties in a deterministic order.
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

// Read reads shortcut properties from a .lnk file
func Read(path string) (Shortcut, error) {
	var sc Shortcut

	// Validate input path
	if err := validatePath(path); err != nil {
		return sc, err
	}

	sh, err := newWshShell()
	if err != nil {
		return sc, fmt.Errorf("failed to initialize shell: %w", err)
	}
	defer sh.Close()

	disp, err := sh.shortcut(path)
	if err != nil {
		return sc, err
	}
	defer disp.Release()

	// Read all properties via the binding table
	for _, binding := range sc.propBindings() {
		prop, err := oleutil.GetProperty(disp, binding.name)
		if err != nil {
			return sc, fmt.Errorf("failed to get property %s: %w", binding.name, err)
		}

		// Convert COM Variant to string. Unset properties (VT_EMPTY/VT_NULL)
		// match no case and leave the field as the empty string.
		switch prop.VT {
		case ole.VT_BSTR:
			*binding.field = prop.ToString()
		case ole.VT_I4:
			*binding.field = fmt.Sprintf("%d", prop.Value())
		}

		// Cleanup VARIANT object immediately to prevent accumulation in loop
		prop.Clear()
	}

	return sc, nil
}

// Write creates a new shortcut (.lnk) file with the given properties
func Write(path string, sc Shortcut) error {
	// Validate input path
	if err := validatePath(path); err != nil {
		return err
	}

	// Set default values if not provided
	if sc.IconLocation == "" {
		sc.IconLocation = DefaultIconLocation
	}
	if sc.WindowStyle == "" {
		sc.WindowStyle = WindowStyleNormal
	}

	sh, err := newWshShell()
	if err != nil {
		return fmt.Errorf("failed to initialize shell: %w", err)
	}
	defer sh.Close()

	disp, err := sh.shortcut(path)
	if err != nil {
		return err
	}
	defer disp.Release()

	// Set all properties via the binding table. Property puts return no
	// value (the result VARIANT stays VT_EMPTY), so there is nothing to clear.
	for _, binding := range sc.propBindings() {
		if _, err := oleutil.PutProperty(disp, binding.name, *binding.field); err != nil {
			return fmt.Errorf("failed to set property %s: %w", binding.name, err)
		}
	}

	if _, err := oleutil.CallMethod(disp, "Save"); err != nil {
		return fmt.Errorf("failed to save shortcut: %w", err)
	}

	return nil
}

// validatePath validates the shortcut file path
func validatePath(path string) error {
	if strings.TrimSpace(path) == "" {
		return errors.New("path cannot be empty")
	}
	if !strings.HasSuffix(strings.ToLower(path), ".lnk") {
		return errors.New("path must have .lnk extension")
	}
	return nil
}

// HRESULT values not exported by go-ole (which defines only S_OK).
const (
	// sFalse: CoInitializeEx succeeded but the thread was already
	// initialized as STA. The reference count is still incremented.
	sFalse = 0x1
	// rpcEChangedMode: CoInitializeEx failed because the thread was
	// already initialized with a different concurrency model (MTA).
	rpcEChangedMode = 0x80010106
)

// wshShell wraps Windows Script Shell COM object
type wshShell struct {
	dispatch *ole.IDispatch
	// coInitialized tracks whether CoInitializeEx actually succeeded on this
	// thread. CoUninitialize must only be called when true; RPC_E_CHANGED_MODE
	// means initialization failed and there is no reference count to release.
	coInitialized bool
}

// newWshShell creates a new Windows Script Shell instance
func newWshShell() (*wshShell, error) {
	runtime.LockOSThread()

	coInitialized := true
	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY); err != nil {
		var oleErr *ole.OleError
		if !errors.As(err, &oleErr) {
			runtime.UnlockOSThread()
			return nil, fmt.Errorf("failed to initialize COM: %w", err)
		}

		hr := oleErr.Code()
		switch hr {
		case sFalse:
			// S_FALSE still increments the reference count, pairing is required.
		case rpcEChangedMode:
			// Initialization failed; must NOT call CoUninitialize later.
			coInitialized = false
		default:
			runtime.UnlockOSThread()
			return nil, fmt.Errorf("failed to initialize COM: %w", err)
		}
	}

	sh := &wshShell{coInitialized: coInitialized}

	unknown, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		sh.Close()
		return nil, fmt.Errorf("failed to create WScript.Shell object: %v", err)
	}

	dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		unknown.Release()
		sh.Close()
		return nil, fmt.Errorf("failed to query interface: %v", err)
	}
	// QueryInterface AddRef'ed the IDispatch, so the IUnknown reference
	// can be released here, before any possible CoUninitialize.
	unknown.Release()
	sh.dispatch = dispatch

	return sh, nil
}

// Close releases all COM resources and unlocks the OS thread locked by
// newWshShell. It must be called exactly once per successful newWshShell
// and is safe to call multiple times.
func (sh *wshShell) Close() {
	if sh.dispatch != nil {
		sh.dispatch.Release()
		sh.dispatch = nil
	}
	if sh.coInitialized {
		ole.CoUninitialize()
		sh.coInitialized = false
	}
	runtime.UnlockOSThread()
}

// shortcut returns the COM shortcut object for the given path.
// It does not create the .lnk file itself.
func (sh *wshShell) shortcut(path string) (*ole.IDispatch, error) {
	res, err := oleutil.CallMethod(sh.dispatch, "CreateShortcut", path)
	if err != nil {
		return nil, fmt.Errorf("failed to create shortcut object: %v", err)
	}
	return res.ToIDispatch(), nil
}
