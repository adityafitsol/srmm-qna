#!/bin/bash

echo "================================================"
echo "  Tailing logs for Agent1 + Agent2"
echo "  Press Ctrl+C to stop watching (scrapers keep running)"
echo "================================================"

touch srmm_agent1.log srmm_agent2.log
tail -f srmm_agent1.log srmm_agent2.log
