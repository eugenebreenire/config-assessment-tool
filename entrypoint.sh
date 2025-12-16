#!/bin/bash
if [ "$1" = "backend" ]; then
  shift
  exec python backend/backend.py "$@"
else
  exec streamlit run frontend/frontend.py "$@"
fi