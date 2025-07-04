package main

import (
	"os"
	"os/exec"
)




func main() {
	// os.Args = []string{
	// 	"xlpro-server.exe",
	// 	"C:\\Users\\Daniel Evans\\.xlpro\\envs\\910a7684-8979-4d80-9add-88e23a63365a_3.13.2\\.venv\\scripts\\python.exe",
	// 	"-m",
	// 	"xlpro.run_server",
	// 	"--debugpy_port=5679",
	// 	"--workbook_path=C:\\Users\\Daniel Evans\\projects\\xlpro\\xlpro_examples\\xlpro-ex01-basics.xlsx",
	// }


    if len(os.Args) < 2 {
        println("Usage: wrapper <command>")
        os.Exit(1)
    }


	// First arg is the command; rest are arguments
    cmd := exec.Command(os.Args[1], os.Args[2:]...)

    // Optional: Inherit or capture stdout/stderr
    cmd.Stdout = os.Stdout
    cmd.Stderr = os.Stderr
    cmd.Stdin = os.Stdin


    // // Join all args as a single command
    // cmdLine := strings.Join(os.Args[1:], " ")

    // // Run the command via shell
    // cmd := exec.Command("cmd", "/K", cmdLine) // Windows-specific

    // // Optional: inherit stdout/stderr
    // cmd.Stdout = os.Stdout
    // cmd.Stderr = os.Stderr

    err := cmd.Run()
    if err != nil {
        os.Exit(1)
    }
}
