package autostart

import (
	"os"
	"path/filepath"

	ole "github.com/go-ole/go-ole"
	oleutil "github.com/go-ole/go-ole/oleutil"
)

var startupDir string

func init() {
	startupDir = filepath.Join(os.Getenv("USERPROFILE"), "AppData", "Roaming", "Microsoft", "Windows", "Start Menu", "Programs", "Startup")
}

func (a *App) path() string {
	return filepath.Join(startupDir, a.Name+".lnk")
}

func (a *App) IsEnabled() bool {
	_, err := os.Stat(a.path())
	return err == nil
}

func (a *App) Enable() error {
	if _, err := os.Lstat(startupDir); err != nil && os.IsNotExist(err) {
		if err := os.MkdirAll(startupDir, 0777); err != nil {
			return err
		}
	} else if err != nil {
		return err
	}

	ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED|ole.COINIT_SPEED_OVER_MEMORY)
	oleShellObject, err := oleutil.CreateObject("WScript.Shell")
	if err != nil {
		return err
	}
	defer oleShellObject.Release()
	wshell, err := oleShellObject.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return err
	}
	defer wshell.Release()
	cs, err := oleutil.CallMethod(wshell, "CreateShortcut", a.path())
	if err != nil {
		return err
	}
	idispatch := cs.ToIDispatch()
	oleutil.PutProperty(idispatch, "TargetPath", a.Exec[0])
	oleutil.CallMethod(idispatch, "Save")

	return nil
}

func (a *App) Disable() error {
	return os.Remove(a.path())
}
