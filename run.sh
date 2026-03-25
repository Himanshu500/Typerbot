#!/bin/bash
# run.sh — Auto-restart bot on crash
echo "🤖 Starting bot with auto-restart..."
while true; do
    python bot.py
    echo "⚠️  Bot stopped. Restarting in 5 seconds..."
    sleep 5
done
