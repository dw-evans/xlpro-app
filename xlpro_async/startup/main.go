package main

import (
	"fmt"
	"log"
	"os"
	"os/exec"

	"github.com/BurntSushi/toml"
)

// Config struct to hold the TOML data
type Config struct {
	PythonPath string `toml:"python_path"`
	RunServerPath string `toml:"run_server_path"`
}

func main() {
	// Define the path to the config.toml file
	configPath := "../config.toml"

	// Read and parse the TOML file
	var config Config
	_, err := toml.DecodeFile(configPath, &config)
	if err != nil {
		log.Fatalf("Error reading TOML file: %v", err)
	}

	// Print the extracted python_interpreter value
	fmt.Printf("Using Python interpreter: %s\n", config.PythonPath)

	// Define the path to the Python script you want to run
	// pythonScriptPath := "path/to/your/script.py"

	// Construct the command to run the Python interpreter with the script
	cmd := exec.Command(config.PythonPath, config.RunServerPath)

	// Set up to capture the command's output
	cmd.Stdout = os.Stdout
	cmd.Stderr = os.Stderr

	// Run the command
	err = cmd.Run()
	if err != nil {
		log.Fatalf("Error running Python script: %v", err)
	}
}
