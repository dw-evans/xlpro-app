package main

import (
	"os"
	"os/exec"
	"syscall"
	"unsafe"
	// "syscall"
	// "unsafe"
)

func enableVirtualTerminalProcessing() {
	const (
		STD_OUTPUT_HANDLE                 = -11
		ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
	)

	kernel32 := syscall.NewLazyDLL("kernel32.dll")
	getStdHandle := kernel32.NewProc("GetStdHandle")
	getConsoleMode := kernel32.NewProc("GetConsoleMode")
	setConsoleMode := kernel32.NewProc("SetConsoleMode")

    tmp:=int32(STD_OUTPUT_HANDLE)
	hOut, _, _ := getStdHandle.Call(uintptr(tmp))

	var mode uint32
	getConsoleMode.Call(hOut, uintptr(unsafe.Pointer(&mode)))
	setConsoleMode.Call(hOut, uintptr(mode|ENABLE_VIRTUAL_TERMINAL_PROCESSING))
}

// func setConsoleTitle(title string) {
// 	cmd := exec.Command("cmd", "/C", "title", title)
// 	cmd.Run()
// }

func main() {
	// os.Args = []string{
    //     "xlpro-server.exe",
    //     "C:/Users/Daniel Evans/projects/xlpro/.venv/Scripts/xlpro-cli.exe",
    //     "init",
    //     "C:/Users/Daniel Evans/projects/xlpro/xlpro_examples/xlpro-ex01-basics.xlsx",
	// }
    
    enableVirtualTerminalProcessing()

    // // conHostPath := "C:/Program Files/WindowsApps/Microsoft.WindowsTerminal_1.22.11141.0_x64__8wekyb3d8bbwe/wt.exe"
    // conHostPath := "conhost.exe"

    
    // if len(os.Args) < 2 {
    //     println("Usage: wrapper <command>")
    //     os.Exit(1)
    // }

    
	// First arg is the command; rest are arguments
    cmd := exec.Command(os.Args[1], os.Args[2:]...)

    // Optional: Inherit or capture stdout/stderr
    cmd.Stdout = os.Stdout
    cmd.Stderr = os.Stderr
    cmd.Stdin = os.Stdin

    err := cmd.Run()
    if err != nil {
        os.Exit(1)
    }
}
