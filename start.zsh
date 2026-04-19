#!/bin/zsh
source ~/tchop/bin/activate
cd "$(dirname "$0")"
uvicorn api_server:app --host 127.0.0.1 --port 6414
