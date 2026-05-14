#!/bin/bash

echo "Stopping scrapers..."

if [ -f .agent1.pid ]; then
    kill "$(cat .agent1.pid)" 2>/dev/null && echo "Agent1 stopped" || echo "Agent1 already stopped"
    rm -f .agent1.pid
fi
if [ -f .agent2.pid ]; then
    kill "$(cat .agent2.pid)" 2>/dev/null && echo "Agent2 stopped" || echo "Agent2 already stopped"
    rm -f .agent2.pid
fi

echo "Done. Progress is saved — re-run 'bash run.sh' to resume."
