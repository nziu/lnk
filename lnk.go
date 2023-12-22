package lnk

import (
	"reflect"
	"runtime"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// Shortcut the shortcut (.lnk file) property struct
type Shortcut struct {
	// Shortcut target: a file path or a website
	TargetPath string
	// Arguments of shortcut
	Arguments string
	// Description of shortcut
	Description string
	// Hotkey of shortcut
	Hotkey string
	// Shortcut icon path, default: "%SystemRoot%\\System32\\SHELL32.dll,0"
	IconLocation string
	// WindowStyle, "1"(default) for default size and location; "3" for maximized window; "7" for minimized window
	WindowStyle string
	// Working directory of shortcut
	WorkingDirectory string
}

type wShell struct {
	wshShellObject *ole.IUnknown
	wshShell       *ole.IDispatch
}

func newWShell() (*wShell, error) {
	runtime.LockOSThread()
	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	wshShellObject, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		return nil, err
	}
	wshShell, err := wshShellObject.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		defer runtime.UnlockOSThread()
		defer ole.CoUninitialize()
		defer wshShellObject.Release()
		return nil, err
	}
	return &wShell{wshShellObject: wshShellObject, wshShell: wshShell}, nil
}

func (wsh *wShell) Close() {
	defer runtime.UnlockOSThread()
	defer ole.CoUninitialize()
	defer wsh.wshShellObject.Release()
	defer wsh.wshShell.Release()
}

func Read(path string) (shortcut Shortcut, err error) {
	wsh, err := newWShell()
	if err != nil {
		return shortcut, err
	}
	defer wsh.Close()

	createShortcut, err := oleutil.CallMethod(wsh.wshShell, "CreateShortcut", path)
	if err != nil {
		return shortcut, err
	}
	idispatch := createShortcut.ToIDispatch()
	defer idispatch.Release()

	typeOfShortcut := reflect.TypeOf(shortcut)
	valueOfShortcut := reflect.ValueOf(&shortcut).Elem()

	for i := 0; i < typeOfShortcut.NumField(); i++ {
		fieldName := typeOfShortcut.Field(i).Name
		property, err := oleutil.GetProperty(idispatch, fieldName)
		if err != nil {
			return shortcut, err
		}
		valueOfProperty := reflect.ValueOf(property.ToString())
		valueOfShortcut.FieldByName(fieldName).Set(valueOfProperty)
	}

	return shortcut, nil
}

func Make(path string, shortcut Shortcut) error {
	wsh, err := newWShell()
	if err != nil {
		return err
	}
	defer wsh.Close()

	createShortcut, err := oleutil.CallMethod(wsh.wshShell, "CreateShortcut", path)
	if err != nil {
		return err
	}
	idispatch := createShortcut.ToIDispatch()
	defer idispatch.Release()

	if shortcut.IconLocation == "" {
		shortcut.IconLocation = "%SystemRoot%\\System32\\SHELL32.dll,0"
	}
	if shortcut.WindowStyle == "" {
		shortcut.WindowStyle = "1"
	}

	typeOfShortcut := reflect.TypeOf(shortcut)
	valueOfShortcut := reflect.ValueOf(&shortcut).Elem()

	for i := 0; i < typeOfShortcut.NumField(); i++ {
		fieldName := typeOfShortcut.Field(i).Name
		fieldValue := valueOfShortcut.Field(i).String()
		_, err := oleutil.PutProperty(idispatch, fieldName, fieldValue)
		if err != nil {
			return err
		}
	}
	_, err = oleutil.CallMethod(idispatch, "Save")
	if err != nil {
		return err
	}
	return nil
}
