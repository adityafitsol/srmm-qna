#!/bin/bash

echo "================================================"
echo "  SRMM Scraper — Starting Both Agents"
echo "================================================"

# Kill any existing runs
if [ -f .agent1.pid ]; then
    OLD=$(cat .agent1.pid)
    kill "$OLD" 2>/dev/null && echo "Stopped old Agent1 (PID $OLD)" || true
fi
if [ -f .agent2.pid ]; then
    OLD=$(cat .agent2.pid)
    kill "$OLD" 2>/dev/null && echo "Stopped old Agent2 (PID $OLD)" || true
fi

# Start Agent 1
nohup python3 -u srmm_scraper.py > srmm_agent1.log 2>&1 &
A1=$!
echo "$A1" > .agent1.pid
echo "Agent1 started — PID: $A1 — log: srmm_agent1.log"

# Start Agent 2
nohup python3 -u srmm_scraper_agent2.py > srmm_agent2.log 2>&1 &
A2=$!
echo "$A2" > .agent2.pid
echo "Agent2 started — PID: $A2 — log: srmm_agent2.log"

echo ""
echo "Both running. Watch logs with: bash logs.sh"
echo "Stop all with:                 bash stop.sh"
