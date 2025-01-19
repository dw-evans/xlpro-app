package main

import (
	"bufio"
	"fmt"
	"log"
	"os"
	"os/exec"
	"strconv"
	"strings"

	"github.com/BurntSushi/toml"
)

// Config struct to hold the TOML data
type Config struct {
	PythonPath string `toml:"python_path"`
	RunServerPath string `toml:"run_server_path"`
}

func main() {
	// Print the working directory
	path, err := os.Getwd()
	if err != nil {
		log.Println(err)
	}
	fmt.Printf("xlpro.exe Working Directory: %s\n", path)
	
	// Define the path to the config.toml file relative to the executable file.
	configPath := "../config.toml"
	
	// Read and parse the TOML file
	var config Config
	_, err = toml.DecodeFile(configPath, &config)
	if err != nil {
		log.Fatalf("Error reading TOML file: %v", err)
	}
	fmt.Printf("Config file being read from: %s\n", configPath)


	// Construct the command to run the Python interpreter with the script
	cmdArgs := []string{
		config.PythonPath, 
		"-Xfrozen_modules=off", 
		config.RunServerPath, 
		"--parent_pid", 
		strconv.Itoa(os.Getpid()),
	}

	fmt.Printf("Running command: %s\n", strings.Join(cmdArgs, " "))
	cmd := exec.Command(cmdArgs[0], cmdArgs[1:]...)
	// cmd := exec.Command("cmd", "/C", "start",config.PythonPath, config.RunServerPath, "--parent_pid", strconv.Itoa(os.Getpid()))

	// Capture the stdout pipe
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		log.Fatalf("Failed to get stdout: %v", err)
	}

	// Capture the stderr pipe (optional, in case you want to handle Python errors)
	stderr, err := cmd.StderrPipe()
	if err != nil {
		log.Fatalf("Failed to get stderr: %v", err)
	}

	// Start the command (instead of Run, which waits for completion)
	err = cmd.Start()
	if err != nil {
		log.Fatal(err)
	}

	// Get the PID of the Python process
	pythonPid := cmd.Process.Pid
	fmt.Printf("Python process started with PID: %d\n", pythonPid)

	// Create goroutines to handle stdout and stderr
	go func() {
		scanner := bufio.NewScanner(stdout)
		for scanner.Scan() {
			log.Printf("[STDOUT] %s", scanner.Text())
		}
		if err := scanner.Err(); err != nil {
			log.Printf("Error reading stdout: %v", err)
		}
	}()

	go func() {
		scanner := bufio.NewScanner(stderr)
		for scanner.Scan() {
			log.Printf("[STDERR] %s", scanner.Text())
		}
		if err := scanner.Err(); err != nil {
			log.Printf("Error reading stderr: %v", err)
		}
	}()


	// Optionally wait for the process to finish (if you need to block until done)
	err = cmd.Wait()
	if err != nil {
		log.Fatal(err)
	}

	if err != nil {
		log.Fatalf("Error running Python script: %v\n", err)
	}
}
