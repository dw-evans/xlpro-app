package main

import (
	"fmt"
	"log"
	"os/exec"

	"github.com/BurntSushi/toml"
)

// Config struct to hold the TOML data
type Config struct {
	PythonPath string `toml:"python_path"`
	RunServerPath string `toml:"run_server_path"`
}

func main() {
	// Define the path to the config.toml file relative to the executable file.
	configPath := "../config.toml"

	// Read and parse the TOML file
	var config Config
	_, err := toml.DecodeFile(configPath, &config)
	if err != nil {
		log.Fatalf("Error reading TOML file: %v", err)
	}

	// Print the extracted python_interpreter value
	fmt.Printf("Using Python interpreter: %s\n", config.PythonPath)
	fmt.Printf("%s %s\n", config.PythonPath, config.RunServerPath)

	// Define the path to the Python script you want to run
	// pythonScriptPath := "path/to/your/script.py"

	// Construct the command to run the Python interpreter with the script
	cmd := exec.Command(config.PythonPath, config.RunServerPath)

	// Start the command (instead of Run, which waits for completion)
	err = cmd.Start()
	if err != nil {
		log.Fatal(err)
	}

	// Get the PID of the Python process
	pythonPid := cmd.Process.Pid
	fmt.Printf("Python process started with PID: %d\n", pythonPid)

	// Optionally wait for the process to finish (if you need to block until done)
	err = cmd.Wait()
	if err != nil {
		log.Fatal(err)
	}

	if err != nil {
		log.Fatalf("Error running Python script: %v\n", err)
	}
}
