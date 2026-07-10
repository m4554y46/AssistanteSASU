@echo off
llama-server -m "C:\Users\micas\tokenforge\backend\spc\models\qwen2.5-1.5b-instruct-q4_k_m.gguf" -c 16384 --host 127.0.0.1 --port 8080
pause