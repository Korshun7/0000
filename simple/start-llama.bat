@echo off
REM Copy llama-server.exe and the GGUF next to this file. No network flags.
llama-server.exe -m Qwen3-8B-Instruct-Q5_K_M.gguf --host 127.0.0.1 --port 8081 --jinja -c 8192 -ngl 99
