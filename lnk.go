// Package lnk provides functionality to create and read Windows shortcut (.lnk) files.
// It uses Windows Script Shell COM object to interact with shortcut files.
//
// Example usage:
//
//	// Create a new shortcut
//	shortcut := lnk.Shortcut{
//		TargetPath:       "C:\\Program Files\\MyApp\\myapp.exe",
//		Description:      "My Application",
//		WorkingDirectory: "C:\\Program Files\\MyApp",
//	}
//	err := lnk.Make("C:\\Users\\Desktop\\MyApp.lnk", shortcut)
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
	DefaultWindowStyle  = "1"
	MaximizedWindow     = "3"
	MinimizedWindow     = "7"
)

// Error messages
var (
	ErrEmptyPath      = errors.New("path cannot be empty")
	ErrInvalidPath    = errors.New("path must have .lnk extension")
	ErrCreateObject   = errors.New("failed to create WScript.Shell object")
	ErrQueryInterface = errors.New("failed to query interface")
	ErrCreateShortcut = errors.New("failed to create shortcut object")
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
	// If empty, defaults to DefaultIconLocation
	IconLocation string

	// WindowStyle controls how the target window is displayed:
	// "1" (default) - normal window
	// "3" - maximized window
	// "7" - minimized window
	// Use the provided constants: DefaultWindowStyle, MaximizedWindow, MinimizedWindow
	WindowStyle string

	// WorkingDirectory is the initial working directory when the target is launched
	WorkingDirectory string
}

// properties returns a map of COM property names to their corresponding field pointers
// This method provides a centralized way to iterate over all shortcut properties
func (s *Shortcut) properties() map[string]*string {
	return map[string]*string{
		"TargetPath":       &s.TargetPath,
		"Arguments":        &s.Arguments,
		"Description":      &s.Description,
		"Hotkey":           &s.Hotkey,
		"IconLocation":     &s.IconLocation,
		"WindowStyle":      &s.WindowStyle,
		"WorkingDirectory": &s.WorkingDirectory,
	}
}

// wShell wraps Windows Script Shell COM object
type wShell struct {
	wshShellObject *ole.IUnknown
	wshShell       *ole.IDispatch
}

// newWShell creates a new Windows Script Shell instance
func newWShell() (*wShell, error) {
	runtime.LockOSThread()

	if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY); err != nil {
		runtime.UnlockOSThread()
		return nil, fmt.Errorf("failed to initialize COM: %w", err)
	}

	wshShellObject, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, fmt.Errorf("%w: %v", ErrCreateObject, err)
	}

	wshShell, err := wshShellObject.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		wshShellObject.Release()
		ole.CoUninitialize()
		runtime.UnlockOSThread()
		return nil, fmt.Errorf("%w: %v", ErrQueryInterface, err)
	}

	return &wShell{
		wshShellObject: wshShellObject,
		wshShell:       wshShell,
	}, nil
}

// Close properly releases all COM resources
func (w *wShell) Close() {
	if w.wshShell != nil {
		w.wshShell.Release()
	}
	if w.wshShellObject != nil {
		w.wshShellObject.Release()
	}
	ole.CoUninitialize()
	runtime.UnlockOSThread()
}

// createShortcut creates a shortcut COM object for the given path
func (w *wShell) createShortcut(path string) (*ole.IDispatch, error) {
	result, err := oleutil.CallMethod(w.wshShell, "CreateShortcut", path)
	if err != nil {
		return nil, fmt.Errorf("%w: %v", ErrCreateShortcut, err)
	}
	return result.ToIDispatch(), nil
}

// validatePath validates the shortcut file path
func validatePath(path string) error {
	if strings.TrimSpace(path) == "" {
		return ErrEmptyPath
	}
	if !strings.HasSuffix(strings.ToLower(path), ".lnk") {
		return ErrInvalidPath
	}
	return nil
}

// Read reads shortcut properties from a .lnk file
func Read(path string) (Shortcut, error) {
	var shortcut Shortcut

	// Validate input path
	if err := validatePath(path); err != nil {
		return shortcut, err
	}

	wsh, err := newWShell()
	if err != nil {
		return shortcut, fmt.Errorf("failed to initialize shell: %w", err)
	}
	defer wsh.Close()

	idispatch, err := wsh.createShortcut(path)
	if err != nil {
		return shortcut, err
	}
	defer idispatch.Release()

	// Read all properties using the shortcut's properties map
	for propName, fieldPtr := range shortcut.properties() {
		property, err := oleutil.GetProperty(idispatch, propName)
		if err != nil {
			return shortcut, fmt.Errorf("failed to get property %s: %w", propName, err)
		}
		if property.VT == ole.VT_BSTR {
			*fieldPtr = property.ToString()
		}
	}

	return shortcut, nil
}

// Make creates a new shortcut (.lnk) file with the given properties
func Make(path string, shortcut Shortcut) error {
	// Validate input path
	if err := validatePath(path); err != nil {
		return err
	}

	// Set default values if not provided
	if shortcut.IconLocation == "" {
		shortcut.IconLocation = DefaultIconLocation
	}
	if shortcut.WindowStyle == "" {
		shortcut.WindowStyle = DefaultWindowStyle
	}

	wsh, err := newWShell()
	if err != nil {
		return fmt.Errorf("failed to initialize shell: %w", err)
	}
	defer wsh.Close()

	idispatch, err := wsh.createShortcut(path)
	if err != nil {
		return err
	}
	defer idispatch.Release()

	// Set all properties using the shortcut's properties map
	for propName, fieldPtr := range shortcut.properties() {
		if _, err := oleutil.PutProperty(idispatch, propName, *fieldPtr); err != nil {
			return fmt.Errorf("failed to set property %s: %w", propName, err)
		}
	}

	if _, err := oleutil.CallMethod(idispatch, "Save"); err != nil {
		return fmt.Errorf("failed to save shortcut: %w", err)
	}

	return nil
}
